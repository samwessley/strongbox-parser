import pandas as pd
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext
import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from copy import copy
import math
import sys

class StrongboxParser:
    def __init__(self):
        self.source_file = None
        self.output_dir = None
        self.start_date = None
        self.end_date = None
        self.source_data = {}
        self.template_data = {}
        self.root = None
        self.status_label = None
        self.progress_bar = None
        self.output_filename = None
        self.console_output = None  # Will hold the text widget for console output

    def print_and_log(self, message):
        """Print to console and also log to GUI if available"""
        print(message)  # Still print to console
        if self.console_output:
            try:
                self.console_output.insert(tk.END, message + "\n")
                self.console_output.see(tk.END)  # Auto-scroll to bottom
                if self.root:
                    self.root.update()  # Update GUI
            except:
                pass  # Ignore errors if GUI is not available

    def update_status(self, message, progress=None):
        """Update the status message and progress bar"""
        if self.status_label:
            self.status_label.config(text=message)
        if progress is not None and self.progress_bar:
            self.progress_bar['value'] = progress
        if self.root:
            self.root.update()

    def get_file_paths(self):
        """DEPRECATED: Use select_source_file and select_output_location instead"""
        pass

    def determine_date_range(self):
        """Automatically determine date range from TB tab"""
        self.update_status("Determining date range from source file...", 40)
        self.print_and_log("\nDetermining date range automatically from TB tab...")
        
        # Get TB date range
        date_row = pd.read_excel(self.source_file, sheet_name='TB', header=None, nrows=1, skiprows=3)
        date_row = date_row.iloc[0]
        
        # Convert dates to datetime
        date_columns = {}
        for col_idx, value in enumerate(date_row):
            try:
                if pd.notna(value):
                    # Try to handle different date formats
                    if isinstance(value, str):
                        # Try different date formats
                        for fmt in ['%m/%d/%Y', '%Y-%m-%d', '%m-%d-%Y', '%d-%b-%Y', '%d/%b/%Y']:
                            try:
                                date = datetime.strptime(value, fmt)
                                date_columns[date] = col_idx
                                break
                            except ValueError:
                                continue
                    else:
                        date = pd.to_datetime(value)
                        date_columns[date] = col_idx
            except Exception as e:
                self.print_and_log(f"Error processing date at column {col_idx}: {str(e)}")
                continue
        
        if not date_columns:
            raise Exception("No valid dates found in row 4 of the TB sheet. Please ensure dates are in a standard format.")
        
        # Find the earliest and latest dates in the TB sheet
        tb_dates = sorted(date_columns.keys())
        
        if len(tb_dates) < 2:
            raise Exception("Not enough dates found in TB sheet. Need at least two dates for beginning and ending balances.")
        
        tb_earliest = tb_dates[0]
        tb_latest = tb_dates[-1]
        
        self.print_and_log(f"TB date range: {tb_earliest.strftime('%Y-%m-%d')} to {tb_latest.strftime('%Y-%m-%d')}")
        
        # Set the date range based on TB dates
        # Beginning balance date is the earliest TB date (end of month)
        self.begin_balance_date = tb_earliest
        # Transaction start date is the day after the earliest TB date
        self.start_date = tb_earliest + relativedelta(days=1)
        # Ending date is the latest TB date (end of month)
        self.end_date = tb_latest
        
        # Print the final determined date range
        self.print_and_log(f"\nFinal determined date range:")
        self.print_and_log(f"Beginning Balance Date: {self.begin_balance_date.strftime('%Y-%m-%d')} (TB earliest)")
        self.print_and_log(f"Transaction Start Date: {self.start_date.strftime('%Y-%m-%d')} (day after beginning balance)")
        self.print_and_log(f"Ending Balance Date: {self.end_date.strftime('%Y-%m-%d')} (TB latest)")
        
        self.update_status(f"Date range determined: {self.start_date.strftime('%m/%d/%Y')} - {self.end_date.strftime('%m/%d/%Y')}", 45)
        
        return date_columns

    def get_last_day_of_month(self, date):
        """Return the last day of the month for a given date"""
        # If the date is already the last day, just return it
        next_month = date.replace(day=28) + relativedelta(days=4)  # Move to next month
        last_day = next_month - relativedelta(days=next_month.day)  # Subtract days to get last day
        return last_day

    def _initialize_excel_loading(self):
        """Initialize Excel file loading and get sheet names"""
        self.print_and_log("\nStarting to load source data...")
        self.update_status("Loading transaction data...", 40)
        
        self.print_and_log("Opening Excel file...")
        excel_file_pd = pd.ExcelFile(self.source_file)
        self.print_and_log(f"Found sheets: {excel_file_pd.sheet_names}")
        return excel_file_pd

    def _try_openpyxl_fallback(self, source_file_path, sheet_to_read, use_data_only):
        """Fallback method to read Excel sheets using openpyxl when pandas fails"""
        self.print_and_log(f"FALLBACK ATTEMPT with data_only={use_data_only} for sheet '{sheet_to_read}'.")
        fallback_workbook = None
        df_from_this_attempt = None
        try:
            fallback_workbook = openpyxl.load_workbook(source_file_path, data_only=use_data_only, read_only=True)
            if sheet_to_read not in fallback_workbook.sheetnames:
                self.print_and_log(f"FALLBACK WARNING: Sheet '{sheet_to_read}' not found in workbook (data_only={use_data_only}).")
                return None

            sheet_obj = fallback_workbook[sheet_to_read]
            data_rows = []
            header = []
            if sheet_obj.max_row > 0:
                header = [cell.value for cell in sheet_obj[1]]
            else:
                self.print_and_log(f"FALLBACK WARNING: Sheet '{sheet_to_read}' (data_only={use_data_only}) appears empty.")

            problematic_cell_count = 0
            for row_idx, row in enumerate(sheet_obj.iter_rows(min_row=2)):
                values = {}
                for col_idx, cell in enumerate(row):
                    cell_value = None
                    try:
                        cell_value = cell.value
                    except Exception as cell_err:
                        self.print_and_log(f"FALLBACK CELL ERROR (data_only={use_data_only}): Sheet '{sheet_to_read}', Row {row_idx + 2}, Col {col_idx + 1}. Error: {cell_err}. Using None.")
                        problematic_cell_count += 1
                    if col_idx < len(header):
                        if header[col_idx] is not None: 
                            values[header[col_idx]] = cell_value
                        else: 
                            values[f"Unknown_Header_Col_{col_idx+1}"] = cell_value 
                    else: 
                        values[f"Extra_Data_Col_{col_idx+1}"] = cell_value
                data_rows.append(values)
            
            df_from_this_attempt = pd.DataFrame(data_rows)
            if df_from_this_attempt.empty and not data_rows and header: 
                df_from_this_attempt = pd.DataFrame(columns=header)
            
            if problematic_cell_count > 0:
                self.print_and_log(f"FALLBACK INFO (data_only={use_data_only}): Encountered {problematic_cell_count} problematic cell(s) in '{sheet_to_read}'.")
            self.print_and_log(f"FALLBACK SUCCESS (data_only={use_data_only}): Successfully read and extracted data for '{sheet_to_read}'.")
            return df_from_this_attempt

        except Exception as current_attempt_err:
            self.print_and_log(f"FALLBACK ERROR (data_only={use_data_only}) for '{sheet_to_read}': {current_attempt_err} (Type: {type(current_attempt_err)})")
            raise current_attempt_err
        finally:
            if fallback_workbook:
                fallback_workbook.close()
                self.print_and_log(f"FALLBACK CLOSED (data_only={use_data_only}): Workbook for '{sheet_to_read}'.")

    def load_source_data(self):
        """Load data from source file"""
        try:
            # Load transaction data
            excel_file_pd = self._initialize_excel_loading()
            
            processed_txn_sheets_count = 0
            # workbook_openpyxl = None # No longer needed as a shared instance

            for sheet_name in excel_file_pd.sheet_names:
                if sheet_name.startswith('TXN-FY'):
                    self.print_and_log(f"\nAttempting to process sheet: {sheet_name}")
                    df = None
                    try:
                        # Attempt 1: Standard pandas read with lenient options
                        self.print_and_log(f"Reading sheet data for {sheet_name} using pandas default...")
                        df = pd.read_excel(
                            excel_file_pd, # Use the pd.ExcelFile object for potentially better performance
                            sheet_name=sheet_name,
                            engine='openpyxl',
                            na_filter=False,
                            keep_default_na=False
                        )
                        self.print_and_log(f"Successfully read sheet {sheet_name} using pandas default.")

                    except Exception as e:
                        self.print_and_log(f"Pandas default read failed for {sheet_name}: {str(e)} (Type: {type(e)})")
                        self.print_and_log(f"Attempting fallback read for {sheet_name} using openpyxl cell by cell...")
                        
                        # Main fallback execution flow
                        df = None
                        try:
                            df = self._try_openpyxl_fallback(self.source_file, sheet_name, use_data_only=True)
                        except Exception as e_data_true_attempt:
                            if "Value must be either numerical or a string containing a wildcard" in str(e_data_true_attempt):
                                self.print_and_log(f"Fallback with data_only=True failed with target error. Trying data_only=False for {sheet_name}.")
                                try:
                                    df = self._try_openpyxl_fallback(self.source_file, sheet_name, use_data_only=False)
                                except Exception as e_data_false_attempt:
                                    self.print_and_log(f"Fallback with data_only=False also failed for {sheet_name}: {str(e_data_false_attempt)}")
                                    # df remains None
                            else:
                                self.print_and_log(f"Fallback with data_only=True failed with an unexpected error for {sheet_name}, not retrying with data_only=False.")
                                # df remains None
                    
                    if df is None: # Check df, which would be None if all attempts failed
                        skipped_message = f"Sheet '{sheet_name}' could not be read by any method and will be SKIPPED. Please make sure to add these journal entries to the template manually."
                        self.print_and_log(skipped_message)
                        messagebox.showwarning("Sheet Read Error", skipped_message)
                        continue # To the next sheet in the outer loop

                    # Common processing for df (whether from pandas or openpyxl fallback)
                    self.print_and_log(f"Columns in {sheet_name}: {df.columns.tolist()}")
                    self.print_and_log(f"Number of rows in {sheet_name}: {len(df)}")
                    
                    df = self._apply_data_type_conversions(df, sheet_name)
                    
                    date_result = self._process_date_columns(df, sheet_name)
                    if date_result is None:
                        continue  # Skip this sheet if date processing failed
                    
                    df, has_fiscal_month = date_result
                    fiscal_month_dates = None
                    
                    if has_fiscal_month:
                        fiscal_month_dates = df['Fiscal Month'].copy()
                        
                    # Apply date range filtering and adjustments
                    filtered_df = self._apply_date_range_filtering(df, sheet_name, has_fiscal_month)
                    
                    if filtered_df is not None:
                        # Add the dataframe to our source data
                        self.source_data[sheet_name] = filtered_df
                        self.print_and_log(f"Successfully added filtered data from {sheet_name} to source_data.")
                        processed_txn_sheets_count += 1
                    else:
                        self.print_and_log(f"INFO: No data from {sheet_name} within the specified date range based on Fiscal Month. Sheet will not be in final output.")
            
            if processed_txn_sheets_count == 0 and any(s.startswith('TXN-FY') for s in excel_file_pd.sheet_names):
                message = "CRITICAL: No transaction (TXN-FY) sheets could be successfully processed after all attempts. The output may be incomplete or empty regarding journal entries. Please make sure to add all required journal entries to the template manually. Check the console logs for details on which sheets failed."
                self.print_and_log(message)
                messagebox.showerror("Critical Data Processing Error", message)
                self.update_status(message, 45)

            # Load trial balance data
            tb_data = self._load_trial_balance_data()
            self.source_data['TB'] = tb_data
            
            # Load TB-DATA sheet for balance information
            self.print_and_log("\nLoading TB-DATA sheet...")
            self.update_status("Loading TB-DATA sheet...", 75)
            try:
                tb_data_df = pd.read_excel(
                    excel_file_pd,
                    sheet_name='TB-DATA',
                    engine='openpyxl',
                    na_filter=False,
                    keep_default_na=False
                )
                self.source_data['TB-DATA'] = tb_data_df
                self.print_and_log(f"Successfully loaded TB-DATA sheet with {len(tb_data_df)} rows")
                self.print_and_log(f"TB-DATA columns: {tb_data_df.columns.tolist()}")
                
                # Show first few rows for debugging
                self.print_and_log("\nFirst few rows of TB-DATA:")
                self.print_and_log(tb_data_df.head())
                
            except Exception as e:
                self.print_and_log(f"Warning: Could not load TB-DATA sheet: {str(e)}")
                self.print_and_log("Will use default balance values (0) if TB-DATA is not available")
                self.source_data['TB-DATA'] = None

        except Exception as e:
            self.print_and_log(f"Error loading source data: {str(e)}")
            raise

    def create_journal_entries(self):
        """Create Journal Entries & Lines tab"""
        self.update_status("Creating journal entries...", 60)
        
        # Step 1: Get transaction sheets
        self.print_and_log("\nStep 1: Getting transaction sheets")
        transaction_sheets = {k: v for k, v in self.source_data.items() if k.startswith('TXN-FY')}
        self.print_and_log(f"Found sheets: {list(transaction_sheets.keys())}")
        
        # Step 2: Process each sheet individually
        self.print_and_log("\nStep 2: Processing individual sheets")
        processed_sheets = []
        for sheet_name, df in transaction_sheets.items():
            try:
                self.print_and_log(f"\nProcessing sheet: {sheet_name}")
                # Create a copy of the dataframe
                df_copy = df.copy()
                
                # Convert columns to appropriate types
                self.print_and_log("Converting columns...")
                df_copy['Transaction Id'] = df_copy['Transaction Id'].astype(str)
                df_copy['Memo'] = df_copy['Memo'].fillna('')
                df_copy['Doc/Ref No'] = df_copy['Doc/Ref No'].fillna('')
                df_copy['Account Id'] = df_copy['Account Id'].astype(str)
                
                # Convert numeric columns
                self.print_and_log("Converting numeric columns...")
                df_copy['Debit'] = pd.to_numeric(df_copy['Debit'], errors='coerce').fillna(0)
                df_copy['Credit'] = pd.to_numeric(df_copy['Credit'], errors='coerce').fillna(0)
                
                # Create the required columns
                self.print_and_log("Creating required columns...")
                processed_df = pd.DataFrame({
                    'Journal ID': df_copy['Transaction Id'],
                    'Journal Entry Description': df_copy['Doc/Ref No'],
                    'Posted Date': df_copy['Transaction Date'],
                    'Account ID': df_copy['Account Id'],
                    'Journal Line Description': df_copy['Memo'],
                    'Debit Amount': df_copy['Debit'],
                    'Credit Amount': df_copy['Credit']
                })
                
                processed_sheets.append(processed_df)
                self.print_and_log(f"Successfully processed sheet: {sheet_name}")
                
            except Exception as e:
                self.print_and_log(f"Error processing sheet {sheet_name}: {str(e)}")
                self.print_and_log("DataFrame info:")
                self.print_and_log(df.info())
                self.print_and_log("\nSample data:")
                self.print_and_log(df.head())
                raise
        
        # Step 3: Combine all processed sheets
        self.print_and_log("\nStep 3: Combining processed sheets")
        try:
            if not processed_sheets:
                raise Exception("No sheets were successfully processed")
            
            journal_entries = pd.concat(processed_sheets, ignore_index=True)
            self.print_and_log("Successfully combined all sheets")
            return journal_entries
            
        except Exception as e:
            self.print_and_log(f"Error combining sheets: {str(e)}")
            raise

    def create_trial_balance(self):
        """Create Comparative Trial Balances tab"""
        self.update_status("Creating trial balance...", 70)
        tb_data = self.source_data['TB']
        
        # Only filter out exact header matches, not all rows containing "account"
        tb_data = tb_data[~((tb_data['Account Id'].str.lower() == "account id") | 
                           (tb_data['Account Id'].str.lower() == "account"))]
        
        # Create new dataframe with required columns, but we'll update balances from TB-DATA
        trial_balance = pd.DataFrame({
            'Account ID': tb_data['Account Id'],
            'Account Name': tb_data['Account Name'],
            'Beginning Balance \n(Prior Period Balance)': 0.0,  # Will be populated from TB-DATA
            'Ending Balance': 0.0,  # Will be populated from TB-DATA
            'Account Type \n(see Mapping Categories tab)': '',
            'Account Mapping \n(see Mapping Categories tab)': '',
            'Account Description': tb_data['Financial Statement Classification']
        })
        
        # Update balances from TB-DATA sheet
        self.print_and_log("\nUpdating balances from TB-DATA sheet...")
        self.update_status("Updating balances from TB-DATA...", 72)
        
        updated_accounts = 0
        total_begin_from_tbdata = 0.0
        total_end_from_tbdata = 0.0
        
        for idx, row in trial_balance.iterrows():
            account_id = row['Account ID']
            begin_balance, end_balance = self._extract_balances_from_tb_data(
                account_id, self.begin_balance_date, self.end_date
            )
            
            trial_balance.at[idx, 'Beginning Balance \n(Prior Period Balance)'] = begin_balance
            trial_balance.at[idx, 'Ending Balance'] = end_balance
            
            total_begin_from_tbdata += begin_balance
            total_end_from_tbdata += end_balance
            
            if begin_balance != 0.0 or end_balance != 0.0:
                updated_accounts += 1
        
        self.print_and_log(f"Updated balances for {updated_accounts} accounts from TB-DATA")
        self.print_and_log(f"Total balances extracted from TB-DATA: Begin={total_begin_from_tbdata}, End={total_end_from_tbdata}")
        
        # Apply the class method to populate the Account Type column
        trial_balance['Account Type \n(see Mapping Categories tab)'] = trial_balance['Account Description'].apply(self.determine_account_type)
        
        # Count how many accounts were classified for each type
        account_type_counts = trial_balance['Account Type \n(see Mapping Categories tab)'].value_counts()
        self.print_and_log("\nAccount Type classification summary:")
        for account_type, count in account_type_counts.items():
            if account_type != '':
                self.print_and_log(f"  {account_type}: {count} accounts")
        
        # Fix f-string syntax error by extracting column name with backslash
        account_type_col = 'Account Type \n(see Mapping Categories tab)'
        unclassified_count = (trial_balance[account_type_col] == '').sum()
        self.print_and_log(f"  Unclassified: {unclassified_count} accounts")
        
        # Debug info
        self.print_and_log("\nTrial Balance DataFrame - first few rows:")
        self.print_and_log(f"Column names: {trial_balance.columns.tolist()}")
        self.print_and_log("First 5 rows:")
        self.print_and_log(trial_balance.head(5))
        self.print_and_log(f"Total rows in trial balance: {len(trial_balance)}")
        
        # Calculate sum of beginning and ending balances
        begin_sum = trial_balance['Beginning Balance \n(Prior Period Balance)'].sum()
        end_sum = trial_balance['Ending Balance'].sum()
        self.print_and_log(f"Sum of Beginning Balances: {begin_sum}")
        self.print_and_log(f"Sum of Ending Balances: {end_sum}")
        
        # Show a few sample balances to verify they're being set correctly
        self.print_and_log("\nSample balance values from trial balance:")
        for idx in range(min(5, len(trial_balance))):
            account_id = trial_balance.at[idx, 'Account ID']
            begin_bal = trial_balance.at[idx, 'Beginning Balance \n(Prior Period Balance)']
            end_bal = trial_balance.at[idx, 'Ending Balance']
            self.print_and_log(f"Account {account_id}: Begin={begin_bal}, End={end_bal}")
        
        # Check data types
        self.print_and_log(f"\nData types:")
        self.print_and_log(f"Beginning Balance column type: {trial_balance['Beginning Balance \n(Prior Period Balance)'].dtype}")
        self.print_and_log(f"Ending Balance column type: {trial_balance['Ending Balance'].dtype}")
        
        # Only remove rows that are definitely headers (contain exactly "Account ID")
        headers_to_remove = []
        for idx, row in trial_balance.iterrows():
            account_id = str(row['Account ID']).lower() if pd.notna(row['Account ID']) else ""
            if account_id == "account id" or account_id == "account":
                headers_to_remove.append(idx)
                self.print_and_log(f"Removing header row: {row['Account ID']}")
        
        if headers_to_remove:
            trial_balance = trial_balance.drop(headers_to_remove)
        
        # Verify there are no empty account IDs but don't filter out other accounts
        trial_balance = trial_balance[trial_balance['Account ID'].notna() & (trial_balance['Account ID'] != '')]
        
        # Reset the index after filtering
        trial_balance = trial_balance.reset_index(drop=True)
        
        return trial_balance

    def _setup_output_file_path(self):
        """Setup the output file path"""
        output_file = os.path.join(
            self.output_dir,
            f"{self.output_filename}_{self.start_date.strftime('%Y%m%d')}_{self.end_date.strftime('%Y%m%d')}.xlsx"
        )
        return output_file

    def _prepare_output_data(self):
        """Prepare trial balance and journal entries data for output"""
        # Load the trial balance data
        self.update_status("Creating trial balance...", 85)
        trial_balance = self.create_trial_balance()
        
        # Process the Journal Entries & Lines
        self.update_status("Creating journal entries...", 90)
        journal_entries = self.create_journal_entries()
        
        # Add accounts from journal entries that are missing from trial balance
        self.print_and_log("\nChecking for accounts in journal entries that are missing from trial balance...")
        self.update_status("Adding missing accounts to trial balance...", 87)
        
        # Get all account IDs in trial balance
        tb_account_ids = set(trial_balance['Account ID'].astype(str))
        
        # Get all unique accounts from journal entries with their account names
        je_accounts = {}
        if 'Account ID' in journal_entries.columns and len(journal_entries) > 0:
            for sheet_name, df in self.source_data.items():
                if sheet_name.startswith('TXN-FY') and 'Account Id' in df.columns and 'Account Name' in df.columns:
                    for _, row in df.iterrows():
                        account_id = str(row['Account Id'])
                        account_name = str(row['Account Name']) if pd.notna(row['Account Name']) else ''
                        je_accounts[account_id] = account_name
        
        # Find accounts in journal entries but not in trial balance
        missing_accounts = []
        for account_id, account_name in je_accounts.items():
            if account_id not in tb_account_ids and account_id.strip() != '':
                self.print_and_log(f"Found missing account in journal entries: {account_id} - {account_name}")
                
                # Get Financial Statement Classification if available
                fin_statement_class = ''
                account_type = ''
                
                # Try to find Financial Statement Classification from transaction data
                for sheet_name, df in self.source_data.items():
                    if sheet_name.startswith('TXN-FY') and 'Account Id' in df.columns and 'Financial Statement Classification' in df.columns:
                        matching_rows = df[df['Account Id'] == account_id]
                        if not matching_rows.empty and not pd.isna(matching_rows['Financial Statement Classification'].iloc[0]):
                            fin_statement_class = matching_rows['Financial Statement Classification'].iloc[0]
                            account_type = self.determine_account_type(fin_statement_class)
                            break
                
                missing_accounts.append({
                    'Account ID': account_id,
                    'Account Name': account_name,
                    'Beginning Balance \n(Prior Period Balance)': 0.0,
                    'Ending Balance': 0.0,
                    'Account Type \n(see Mapping Categories tab)': account_type,
                    'Account Mapping \n(see Mapping Categories tab)': '',
                    'Account Description': fin_statement_class
                })
        
        # Add missing accounts to trial balance
        if missing_accounts:
            missing_df = pd.DataFrame(missing_accounts)
            trial_balance = pd.concat([trial_balance, missing_df], ignore_index=True)
            self.print_and_log(f"Added {len(missing_accounts)} missing accounts to trial balance")
        
        return trial_balance, journal_entries

    def _clean_data_for_excel(self, trial_balance, journal_entries):
        """Apply ultra-aggressive data cleaning to prevent Excel corruption"""
        # EXTREME data cleaning - replace everything potentially problematic
        def ultra_clean_value(value):
            """Ultra-aggressive cleaning to remove any potential Excel corruption"""
            if pd.isna(value):
                return ""
            elif isinstance(value, (int, float)):
                if math.isinf(value) or math.isnan(value) or abs(value) > 1e15:
                    return 0.0
                return float(value)
            elif isinstance(value, pd.Timestamp):
                try:
                    # Return the datetime object itself, not a string
                    # Excel will handle the formatting
                    return value.to_pydatetime()
                except:
                    return ""
            else:
                try:
                    # Convert to string and encode/decode to clean any encoding issues
                    str_val = str(value).encode('ascii', errors='ignore').decode('ascii')
        
                    # Remove any remaining problematic characters - be extremely aggressive
                    import re
                    # Keep only letters, numbers, spaces, basic punctuation
                    cleaned = re.sub(r'[^\w\s\-\.\,\(\)\:\;\$\%\@\#\!\?\/\\]', ' ', str_val)
                    
                    # Replace multiple spaces with single space
                    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
                    
                    # Limit length
                    if len(cleaned) > 1000:  # Much shorter limit
                        cleaned = cleaned[:1000]
                    
                    return cleaned
                except:
                    return "DATA_ERROR"

        self.print_and_log("Ultra-cleaning all data...")
        
        # Clean trial balance
        for col in trial_balance.columns:
            trial_balance[col] = trial_balance[col].apply(ultra_clean_value)
        
        # Clean journal entries  
        for col in journal_entries.columns:
            journal_entries[col] = journal_entries[col].apply(ultra_clean_value)

        return trial_balance, journal_entries

    def _create_excel_workbook(self):
        """Create Excel workbook with basic styling setup"""
        self.print_and_log("Creating Excel workbook with openpyxl...")
        workbook = openpyxl.Workbook()
        
        # Define styles
        header_font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
        blue_fill = PatternFill(start_color='0070C0', end_color='0070C0', fill_type='solid')
        gray_fill = PatternFill(start_color='999999', end_color='999999', fill_type='solid')
        dark_blue_fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        styles = {
            'header_font': header_font,
            'blue_fill': blue_fill,
            'gray_fill': gray_fill,
            'dark_blue_fill': dark_blue_fill,
            'center_alignment': center_alignment
        }
        
        # Remove default sheet
        if 'Sheet' in workbook.sheetnames:
            workbook.remove(workbook['Sheet'])
        
        return workbook, styles

    def _create_other_sheets(self, workbook, styles):
        """Create and format all other sheets in the specified order"""
        # Create Instructions sheet
        instructions_sheet = workbook.create_sheet('Instructions')
        instructions_sheet.append(['Content for Instructions'])
        
        # Create Data Validation Tests sheet
        validation_sheet = workbook.create_sheet('Data Validation Tests')
        validation_sheet.append(['Content for Data Validation Tests'])
        
        # Create Notes sheet
        notes_sheet = workbook.create_sheet('Notes')
        notes_sheet.append(['Content for Notes'])
        
        # Create Banking Accts sheet with specific formatting
        banking_accts_sheet = workbook.create_sheet('Banking Accts')
        
        # Add headers for Banking Accts
        banking_accts_sheet.append(['Required', 'Required', 'Optional', 'Optional'])
        banking_accts_sheet.append(['Account Number', 'Account Name', 'Institution', 'Currency'])
        
        # Style row 1 headers for Banking Accts
        for col in range(1, 3):  # Columns A-B (Required)
            cell = banking_accts_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['blue_fill']
            cell.alignment = styles['center_alignment']
        
        for col in range(3, 5):  # Columns C-D (Optional)
            cell = banking_accts_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['gray_fill']
            cell.alignment = styles['center_alignment']
        
        # Style row 2 headers for Banking Accts
        for col in range(1, 5):  # All columns A-D
            cell = banking_accts_sheet.cell(row=2, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['dark_blue_fill']
            cell.alignment = styles['center_alignment']
        
        # Create Banking Txn sheet with specific formatting
        banking_txn_sheet = workbook.create_sheet('Banking Txn')
        
        # Add headers for Banking Txn
        banking_txn_sheet.append(['Required', 'Required', 'Required', 'Required'])
        banking_txn_sheet.append(['Posted Date', 'Description', 'Amount', 'Account Number'])
        
        # Style row 1 headers for Banking Txn (all Required)
        for col in range(1, 5):  # Columns A-D (all Required)
            cell = banking_txn_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['blue_fill']
            cell.alignment = styles['center_alignment']
        
        # Style row 2 headers for Banking Txn
        for col in range(1, 5):  # All columns A-D
            cell = banking_txn_sheet.cell(row=2, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['dark_blue_fill']
            cell.alignment = styles['center_alignment']
        
        # Create Mapping Categories sheet
        mapping_sheet = workbook.create_sheet('Mapping Categories')
        mapping_sheet.append(['Content for Mapping Categories'])

    def _handle_excel_creation_error(self, e, output_file, trial_balance, journal_entries):
        """Handle Excel creation errors with CSV fallback"""
        self.print_and_log(f"openpyxl failed: {e}")
        
        # Fallback to CSV files
        csv_dir = os.path.dirname(output_file)
        csv_base = os.path.splitext(os.path.basename(output_file))[0]
        
        tb_csv = os.path.join(csv_dir, f"{csv_base}_TrialBalance.csv")
        je_csv = os.path.join(csv_dir, f"{csv_base}_JournalEntries.csv")
        
        # Write as CSV files
        trial_balance.to_csv(tb_csv, index=False, encoding='utf-8')
        journal_entries.to_csv(je_csv, index=False, encoding='utf-8')
        
        self.print_and_log(f"Created CSV files instead:")
        self.print_and_log(f"Trial Balance: {tb_csv}")
        self.print_and_log(f"Journal Entries: {je_csv}")
            
        # Try to create a minimal Excel file
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "READ_ME"
            sheet['A1'] = "Excel file creation failed due to data corruption issues."
            sheet['A2'] = "Please use the CSV files created instead:"
            sheet['A3'] = f"Trial Balance: {os.path.basename(tb_csv)}"
            sheet['A4'] = f"Journal Entries: {os.path.basename(je_csv)}"
            workbook.save(output_file)
            workbook.close()
        except:
            pass

    def create_output_file(self):
        """Create output Excel file with both tabs"""
        self.update_status("Creating output file...", 80)
        
        # Create output filename
        output_file = self._setup_output_file_path()
        
        # Prepare output data
        trial_balance, journal_entries = self._prepare_output_data()
        
        # Clean data for Excel
        trial_balance, journal_entries = self._clean_data_for_excel(trial_balance, journal_entries)
        
        # Use openpyxl with aggressive error handling
        try:
            workbook, styles = self._create_excel_workbook()
            
            # Create main data sheets first (in desired order)
            tb_sheet = self._create_trial_balance_sheet(workbook, trial_balance, styles)
            je_sheet = self._create_journal_entries_sheet(workbook, journal_entries, styles)
            
            # Create other sheets after main data sheets
            self._create_other_sheets(workbook, styles)
            
            # Save with openpyxl
            workbook.save(output_file)
            self.print_and_log("File created successfully with openpyxl, aggressive data cleaning, and styling")
            
        except Exception as e:
            self._handle_excel_creation_error(e, output_file, trial_balance, journal_entries)
        
        self.update_status("Saving workbook...", 99)
        return output_file

    def process_data(self):
        """Process the data and create output file"""
        try:
            # Determine date range automatically
            self.date_columns = self.determine_date_range()
            
            # Load source data using the determined date range
            self.load_source_data()
            
            # Create output file
            output_file = self.create_output_file()
            
            self.update_status(f"File created successfully at:\n{output_file}", 100)
            messagebox.showinfo("Success", f"File created successfully at:\n{output_file}")
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            messagebox.showerror("Error", str(e))

    def run(self):
        """Run the parser"""
        try:
            # Create main window
            self.root = tk.Tk()
            self.root.title("Strongbox Parser")
            self.root.geometry("900x700")  # Made wider and taller to accommodate console output

            # Create main frame
            main_frame = ttk.Frame(self.root, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Create title label
            title_label = ttk.Label(main_frame, text="Strongbox Parser", font=("Helvetica", 16, "bold"))
            title_label.pack(pady=20)

            # Create source file selection button
            source_button = ttk.Button(main_frame, text="Select Source File", command=self.select_source_file)
            source_button.pack(pady=20)

            # Create status label
            self.status_label = ttk.Label(main_frame, text="Ready to start...", wraplength=500)
            self.status_label.pack(pady=10)

            # Create progress bar
            self.progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate')
            self.progress_bar.pack(pady=10)

            # Create console output section
            console_frame = ttk.LabelFrame(main_frame, text="Console Output", padding="10")
            console_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))

            # Create scrolled text widget for console output
            self.console_output = scrolledtext.ScrolledText(
                console_frame, 
                height=20, 
                width=80,
                font=("Consolas", 9),
                state=tk.NORMAL,
                wrap=tk.WORD
            )
            self.console_output.pack(fill=tk.BOTH, expand=True)

            # Add initial message to console
            self.print_and_log("Strongbox Parser initialized. Ready to process files.")

            # Start the main event loop
            self.root.mainloop()
        except Exception as e:
            if self.root:
                self.update_status(f"Error: {str(e)}", 0)
                messagebox.showerror("Error", str(e))
            else:
                messagebox.showerror("Error", str(e))
                
    def select_source_file(self):
        """Select source file using file dialog"""
        self.update_status("Selecting source file...", 10)
        self.source_file = filedialog.askopenfilename(
            title="Select Strongbox File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not self.source_file:
            self.update_status("Operation cancelled: No source file selected", 0)
            return
            
        self.update_status(f"Selected source file: {os.path.basename(self.source_file)}", 20)
        # Proceed to next step - select output file
        self.select_output_location()
    
    def select_output_location(self):
        """Select output directory and filename at once using save file dialog"""
        self.update_status("Selecting output location...", 30)
        
        # Use a save file dialog to get both the directory and filename at once
        output_file = filedialog.asksaveasfilename(
            title="Save Output File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Audit_Sight_Output.xlsx"
        )
        
        if not output_file:
            self.update_status("Operation cancelled: No output location selected", 0)
            return
            
        # Split the output_file into directory and filename
        self.output_dir = os.path.dirname(output_file)
        self.output_filename = os.path.splitext(os.path.basename(output_file))[0]
        
        self.update_status(f"Selected output: {self.output_filename} in {self.output_dir}", 35)
        
        # Proceed to determine date range and process data
        try:
            # Process the data
            self.process_data()
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            messagebox.showerror("Error", str(e))

    def determine_account_type(self, fs_classification):
        """Determine Account Type based on Financial Statement Classification Path"""
        if pd.isna(fs_classification) or fs_classification == '':
            return ''
            
        fs_classification = str(fs_classification).strip()
        
        if fs_classification.startswith('Total Assets'):
            return 'Assets'
        elif fs_classification.startswith('Total Liabilities and Equity → Total Liabilities'):
            return 'Liabilities'
        elif fs_classification.startswith('Total Liabilities and Equity → Total Equity'):
            return 'Equity'
        elif fs_classification.startswith('Net Income → Operating Profit → Gross Profit → Total Net Sales'):
            return 'Income'
        elif fs_classification.startswith('Net Income → Operating Profit → Gross Profit → Total COGS/COS'):
            return 'Expense'
        elif fs_classification.startswith('Net Income → Operating Profit → Total Operating Expenses'):
            return 'Expense'
        else:
            return ''

    def create_test_file(self):
        """Create a minimal test Excel file to diagnose corruption issues"""
        output_file = os.path.join(
            self.output_dir,
            f"TEST_{self.output_filename}_{self.start_date.strftime('%Y%m%d')}_{self.end_date.strftime('%Y%m%d')}.xlsx"
        )
        
        # Create minimal test data
        test_data = [
            ['Account ID', 'Account Name', 'Beginning Balance', 'Ending Balance'],
            ['100', 'Cash', 1000.00, 1500.00],
            ['200', 'Accounts Receivable', 5000.00, 4500.00],
            ['300', 'Inventory', 2000.00, 2200.00]
        ]
        
        # Try the most basic Excel creation possible
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = 'Test Data'
            
            for row_data in test_data:
                sheet.append(row_data)
            
            workbook.save(output_file)
            self.print_and_log(f"Test file created: {output_file}")
            return output_file
        except Exception as e:
            self.print_and_log(f"Error creating test file: {str(e)}")
            return None

    def _apply_data_type_conversions(self, df, sheet_name):
        """Apply data type conversions to DataFrame columns"""
        self.print_and_log(f"Converting data types for {sheet_name}...")
        
        # Convert string columns
        for col in ['Transaction Id', 'Account Id', 'Memo', 'Doc/Ref No']:
            if col in df.columns:
                df[col] = df[col].astype(str)
            else:
                self.print_and_log(f"Warning: Column '{col}' not found in sheet {sheet_name}. Skipping conversion.")
        
        # Convert numeric columns
        for col in ['Debit', 'Credit']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                self.print_and_log(f"Warning: Column '{col}' not found in sheet {sheet_name}. Skipping conversion.")
        
        return df

    def _process_date_columns(self, df, sheet_name):
        """Process and validate date columns in the DataFrame"""
        # Check if both Transaction Date and Fiscal Month columns exist
        if 'Transaction Date' not in df.columns:
            self.print_and_log(f"WARNING: 'Transaction Date' column NOT FOUND in sheet {sheet_name}. Cannot filter by date.")
            self.print_and_log(f"Skipping sheet {sheet_name} due to missing 'Transaction Date' column.")
            return None
            
        # Process the Transaction Date and Fiscal Month columns
        self.print_and_log(f"INFO: 'Transaction Date' column found in {sheet_name}.")
        self.print_and_log(f"Sample raw 'Transaction Date' values in {sheet_name} before pd.to_datetime:")
        self.print_and_log(df['Transaction Date'].head(10) if len(df) > 0 else "Sheet is empty or has no dates")
        
        # Convert Transaction Date to datetime
        df['Transaction Date'] = pd.to_datetime(df['Transaction Date'], errors='coerce')
        
        # Check if Fiscal Month column exists for date validation
        has_fiscal_month = 'Fiscal Month' in df.columns
        
        if has_fiscal_month:
            self.print_and_log(f"INFO: 'Fiscal Month' column found in {sheet_name}.")
            self.print_and_log(f"Sample raw 'Fiscal Month' values in {sheet_name} before pd.to_datetime:")
            self.print_and_log(df['Fiscal Month'].head(10) if len(df) > 0 else "Sheet is empty or has no fiscal months")
            
            # Convert Fiscal Month to datetime for comparison
            df['Fiscal Month'] = pd.to_datetime(df['Fiscal Month'], errors='coerce')
            
            # Count NaT values in Fiscal Month
            fiscal_nat_count = df['Fiscal Month'].isnull().sum()
            self.print_and_log(f"INFO: Found {fiscal_nat_count} NaT values in 'Fiscal Month' for {sheet_name} after conversion.")
        
        # Count NaT values in Transaction Date
        transaction_nat_count = df['Transaction Date'].isnull().sum()
        self.print_and_log(f"INFO: Found {transaction_nat_count} NaT values in 'Transaction Date' for {sheet_name} after conversion.")
        
        # Filter out rows where date conversion failed (NaT) for Transaction Date
        df = df[pd.notna(df['Transaction Date'])]
        if has_fiscal_month:
            # Keep rows only where Fiscal Month is valid
            df = df[pd.notna(df['Fiscal Month'])]
        
        self.print_and_log(f"INFO: Rows in {sheet_name} after filtering NaT dates: {len(df)}")
        return df, has_fiscal_month

    def _apply_date_range_filtering(self, df, sheet_name, has_fiscal_month):
        """Apply date range filtering and adjust transaction dates if needed"""
        if len(df) == 0:
            self.print_and_log(f"INFO: No valid dates found in {sheet_name} after NaT filtering. Sheet will not be in final output.")
            return None
        
        # Check if the sheet has any rows in our date range
        # Use Fiscal Month for range filtering if available
        if has_fiscal_month:
            # Filter based on Fiscal Month for date range eligibility
            fiscal_month_mask = (df['Fiscal Month'] >= self.start_date) & (df['Fiscal Month'] <= self.end_date)
            df_in_range = df[fiscal_month_mask].copy()
            self.print_and_log(f"INFO: Rows in {sheet_name} with Fiscal Month in date range [{self.start_date.strftime('%Y-%m-%d')} - {self.end_date.strftime('%Y-%m-%d')}]: {len(df_in_range)}")
            
            if len(df_in_range) > 0:
                df_in_range = self._adjust_transaction_dates(df_in_range, sheet_name)
                return df_in_range
            else:
                self.print_and_log(f"INFO: No data from {sheet_name} within the specified date range based on Fiscal Month. Sheet will not be in final output.")
                return None
        else:
            # If no Fiscal Month, fall back to Transaction Date filtering
            transaction_date_mask = (df['Transaction Date'] >= self.start_date) & (df['Transaction Date'] <= self.end_date)
            df_filtered = df[transaction_date_mask]
            self.print_and_log(f"INFO: Rows in {sheet_name} after applying date range to Transaction Date [{self.start_date.strftime('%Y-%m-%d')} - {self.end_date.strftime('%Y-%m-%d')}]: {len(df_filtered)}")
            
            if len(df_filtered) > 0:
                return df_filtered
            else:
                self.print_and_log(f"INFO: No data from {sheet_name} within the specified date range based on Transaction Date. Sheet will not be in final output.")
                return None

    def _adjust_transaction_dates(self, df_in_range, sheet_name):
        """Adjust transaction dates that don't match their fiscal month"""
        # Identify and fix Transaction Dates that don't match their Fiscal Month
        mismatched_dates = 0
        adjusted_dates = []
        
        for idx, row in df_in_range.iterrows():
            transaction_date = row['Transaction Date']
            fiscal_month = row['Fiscal Month']
            
            # Check if Transaction Date's month/year matches Fiscal Month's month/year
            if (transaction_date.year != fiscal_month.year) or (transaction_date.month != fiscal_month.month):
                # Get the last day of the Fiscal Month
                last_day_of_month = self.get_last_day_of_month(fiscal_month)
                adjusted_dates.append((idx, transaction_date, last_day_of_month))
                # Update Transaction Date to the last day of the Fiscal Month
                df_in_range.at[idx, 'Transaction Date'] = last_day_of_month
                mismatched_dates += 1
        
        if mismatched_dates > 0:
            self.print_and_log(f"INFO: Adjusted {mismatched_dates} Transaction Dates to match their Fiscal Month in {sheet_name}")
            if len(adjusted_dates) > 0 and len(adjusted_dates) <= 10:
                self.print_and_log("Sample of adjusted dates (idx, original_date, new_date):")
                for adj in adjusted_dates[:10]:
                    self.print_and_log(f"  Row {adj[0]}: {adj[1].strftime('%Y-%m-%d')} -> {adj[2].strftime('%Y-%m-%d')}")
        
        return df_in_range

    def _load_trial_balance_data(self):
        """Load trial balance data from TB sheet"""
        self.print_and_log("\nLoading trial balance data...")
        self.update_status("Loading trial balance data...", 50)
        
        # Get the date_columns that were determined in determine_date_range
        date_columns = self.date_columns
        
        # Find the closest TB dates to our calculated range
        available_tb_dates = sorted(date_columns.keys())
        
        # Find closest beginning date (on or before our begin_balance_date)
        closest_begin_date = None
        for tb_date in available_tb_dates:
            if tb_date <= self.begin_balance_date:
                closest_begin_date = tb_date
            else:
                break
        
        # If no TB date is on or before our begin_balance_date, use the earliest TB date
        if closest_begin_date is None:
            closest_begin_date = available_tb_dates[0]
            self.print_and_log(f"WARNING: No TB date found on or before calculated begin balance date {self.begin_balance_date.strftime('%Y-%m-%d')}, using earliest TB date {closest_begin_date.strftime('%Y-%m-%d')}")
        
        # Find closest ending date (on or after our end_date)
        closest_end_date = None
        for tb_date in reversed(available_tb_dates):
            if tb_date >= self.end_date:
                closest_end_date = tb_date
            else:
                break
        
        # If no TB date is on or after our end_date, use the latest TB date
        if closest_end_date is None:
            closest_end_date = available_tb_dates[-1]
            self.print_and_log(f"WARNING: No TB date found on or after calculated end date {self.end_date.strftime('%Y-%m-%d')}, using latest TB date {closest_end_date.strftime('%Y-%m-%d')}")
        
        # Use the beginning and ending dates already identified
        begin_date = self.begin_balance_date
        
        self.print_and_log(f"\nTarget dates:")
        self.print_and_log(f"Begin date: {begin_date}")
        self.print_and_log(f"End date: {self.end_date}")
        self.print_and_log(f"Begin balance date (exact): {closest_begin_date}")
        self.print_and_log(f"End balance date (exact): {closest_end_date}")
        
        return self._extract_trial_balance_data(date_columns, closest_begin_date, closest_end_date)

    def _extract_trial_balance_data(self, date_columns, closest_begin_date, closest_end_date):
        """Extract data from the TB sheet using openpyxl"""
        try:
            self.print_and_log("Opening workbook for trial balance data...")
            self.update_status("Opening workbook for trial balance...", 52)
            wb = openpyxl.load_workbook(self.source_file, data_only=True, read_only=True)
            self.print_and_log("Workbook opened successfully")
            
            self.print_and_log("Accessing TB sheet...")
            self.update_status("Accessing TB sheet...", 54)
            tb_sheet = wb['TB']
            self.print_and_log(f"TB sheet accessed. Sheet has {tb_sheet.max_row} rows and {tb_sheet.max_column} columns")
            
            # Find Financial Statement Classification Path column
            self.print_and_log("Finding Financial Statement Classification column...")
            self.update_status("Finding classification column...", 56)
            fin_statement_col = self._find_financial_classification_column(tb_sheet)
            
            # Print column indices for debugging
            self.print_and_log(f"\nColumn indices:")
            self.print_and_log(f"Begin date column index: {date_columns[closest_begin_date]}")
            self.print_and_log(f"End date column index: {date_columns[closest_end_date]}")
            self.print_and_log(f"Financial Statement Classification column index: {fin_statement_col}")
            
            self.print_and_log("Starting to extract data from TB sheet...")
            self.update_status("Extracting trial balance data...", 58)
            
            data = []
            rows_processed = 0
            
            # Limit the number of rows to process to avoid hanging on very large sheets
            max_row_to_process = min(tb_sheet.max_row, 10000)  # Limit to 10,000 rows max
            
            for row_idx in range(8, max_row_to_process + 1):
                rows_processed += 1
                
                # Update progress every 100 rows
                if rows_processed % 100 == 0:
                    self.print_and_log(f"Processed {rows_processed} rows...")
                    progress = 58 + (rows_processed / max_row_to_process) * 10  # 58-68%
                    self.update_status(f"Processing TB row {rows_processed}...", progress)
                
                try:
                    account_id = tb_sheet.cell(row=row_idx, column=4).value
                    if account_id is not None:
                        account_id = str(account_id).strip()
                        # Only skip rows with exact header matches, not all rows containing "account"
                        if account_id.lower() == "account id" or account_id.lower() == "account":
                            if rows_processed <= 20:  # Only print for first 20 rows to avoid spam
                                self.print_and_log(f"Skipping header row {row_idx} with account_id: {account_id}")
                            continue
                            
                        account_name = tb_sheet.cell(row=row_idx, column=6).value
                        begin_cell = tb_sheet.cell(row=row_idx, column=date_columns[closest_begin_date] + 1)
                        end_cell = tb_sheet.cell(row=row_idx, column=date_columns[closest_end_date] + 1)
                        
                        # Get Financial Statement Classification using the identified column
                        fin_statement_class = tb_sheet.cell(row=row_idx, column=fin_statement_col).value
                        
                        # Debug output for Financial Statement Classification for the first few rows
                        if row_idx < 12:
                            self.print_and_log(f"Row {row_idx}, Account: {account_id}, Financial Statement Classification (Column {fin_statement_col}): {fin_statement_class}")
                        
                        # Handle different types appropriately
                        if fin_statement_class is None:
                            fin_statement_class = ''
                        else:
                            fin_statement_class = str(fin_statement_class).strip()
                        
                        try:
                            begin_balance = float(begin_cell.value) if begin_cell.value is not None else 0.0
                        except (ValueError, TypeError):
                            begin_balance = 0.0
                            
                        try:
                            end_balance = float(end_cell.value) if end_cell.value is not None else 0.0
                        except (ValueError, TypeError):
                            end_balance = 0.0
                        
                        data.append({
                            'Account Id': account_id,
                            'Account Name': str(account_name).strip() if account_name is not None else '',
                            'Beginning Balance': begin_balance,
                            'Ending Balance': end_balance,
                            'Financial Statement Classification': fin_statement_class
                        })
                except Exception as e:
                    if rows_processed <= 20:  # Only print errors for first 20 rows to avoid spam
                        self.print_and_log(f"Error processing TB sheet row {row_idx}: {str(e)}. Skipping row.")
                    continue
            
            self.print_and_log(f"Completed processing {rows_processed} rows from TB sheet")
            self.print_and_log(f"Extracted {len(data)} valid account records")
            self.update_status("Converting TB data to DataFrame...", 68)
            
            # Convert to DataFrame
            tb_data = pd.DataFrame(data)
            self.print_and_log("\nFirst few rows of TB data:")
            self.print_and_log(tb_data.head())
            
            return tb_data
            
        except Exception as e:
            self.print_and_log(f"Error in _extract_trial_balance_data: {str(e)}")
            raise
        finally:
            # Close the workbook if it exists
            try:
                if 'wb' in locals():
                    wb.close()
                    self.print_and_log("Workbook closed")
            except:
                pass

    def _find_financial_classification_column(self, tb_sheet):
        """Find the Financial Statement Classification column in the TB sheet"""
        self.print_and_log("Searching for Financial Statement Classification column...")
        fin_statement_col = None
        
        # Search more efficiently - check fewer rows and columns first
        for row_idx in range(1, 6):  # Check first 5 rows for headers
            self.print_and_log(f"Checking row {row_idx} for headers...")
            for col_idx in range(1, 15):  # Check first 14 columns (reduced from 25)
                try:
                    cell_value = tb_sheet.cell(row=row_idx, column=col_idx).value
                    if cell_value:
                        cell_text = str(cell_value).lower()
                        if 'financial statement classification' in cell_text or 'fin statement classification' in cell_text:
                            fin_statement_col = col_idx
                            self.print_and_log(f"Found Financial Statement Classification header in row {row_idx}, column {col_idx}: {cell_value}")
                            return fin_statement_col
                except Exception as e:
                    # Silently continue if there's an error reading a cell
                    continue
            if fin_statement_col:
                break
        
        if not fin_statement_col:
            # Default to column C (3) as specified
            fin_statement_col = 3
            self.print_and_log(f"No Financial Statement Classification header found, using default column C (3)")
        
        return fin_statement_col

    def _create_trial_balance_sheet(self, workbook, trial_balance, styles):
        """Create and format the Comparative Trial Balances sheet"""
        # Create trial balance sheet
        tb_sheet = workbook.create_sheet('Comparative Trial Balances')
        
        # Set column widths for trial balance (converting px to Excel units)
        tb_sheet.column_dimensions['A'].width = 14.3  # 100px
        tb_sheet.column_dimensions['B'].width = 34.3  # 240px
        tb_sheet.column_dimensions['C'].width = 15.7  # 110px
        tb_sheet.column_dimensions['D'].width = 15.7  # 110px
        tb_sheet.column_dimensions['E'].width = 15.7  # 110px
        tb_sheet.column_dimensions['F'].width = 34.3  # 240px
        tb_sheet.column_dimensions['G'].width = 34.3  # 240px
        
        # Write headers
        tb_sheet.append(['Required'] * 5 + ['Optional'] * 2)
        tb_sheet.append(list(trial_balance.columns))
        
        # Style row 1 headers (Required/Optional)
        for col in range(1, 6):  # Columns A-E
            cell = tb_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['blue_fill']
            cell.alignment = styles['center_alignment']
        
        for col in range(6, 8):  # Columns F-G
            cell = tb_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['gray_fill']
            cell.alignment = styles['center_alignment']
        
        # Style row 2 headers (column names) and set row height
        tb_sheet.row_dimensions[2].height = 25  # 33px ≈ 25 points
        for col in range(1, 8):  # All columns
            cell = tb_sheet.cell(row=2, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['dark_blue_fill']
            cell.alignment = styles['center_alignment']
        
        # Write trial balance data with error handling
        for _, row in trial_balance.iterrows():
            row_values = []
            for col in trial_balance.columns:
                value = row[col]
                try:
                    # Double-check the value is clean
                    if isinstance(value, str) and len(value) > 1000:
                        value = value[:1000]
                    row_values.append(value)
                except:
                    row_values.append("ERROR")
            tb_sheet.append(row_values)
        
        return tb_sheet

    def _create_journal_entries_sheet(self, workbook, journal_entries, styles):
        """Create and format the Journal Entries & Lines sheet"""
        # Create journal entries sheet
        je_sheet = workbook.create_sheet('Journal Entries & Lines')
        
        # Set column widths for journal entries (converting px to Excel units)
        je_sheet.column_dimensions['A'].width = 14.3  # 100px
        je_sheet.column_dimensions['B'].width = 48.6  # 340px
        je_sheet.column_dimensions['C'].width = 15.7  # 110px
        je_sheet.column_dimensions['D'].width = 20.0  # 140px
        je_sheet.column_dimensions['E'].width = 35.7  # 250px
        je_sheet.column_dimensions['F'].width = 15.7  # 110px
        je_sheet.column_dimensions['G'].width = 15.7  # 110px
        
        # Write headers
        je_sheet.append(['Required', 'Optional', 'Required', 'Required', 'Optional', 'Required', 'Required'])
        je_sheet.append(list(journal_entries.columns))
        
        # Style row 1 headers (Required/Optional)
        required_cols = [1, 3, 4, 6, 7]  # Columns A, C, D, F, G
        optional_cols = [2, 5]  # Columns B, E
        
        for col in required_cols:
            cell = je_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['blue_fill']
            cell.alignment = styles['center_alignment']
        
        for col in optional_cols:
            cell = je_sheet.cell(row=1, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['gray_fill']
            cell.alignment = styles['center_alignment']
        
        # Style row 2 headers (column names)
        for col in range(1, 8):  # All columns
            cell = je_sheet.cell(row=2, column=col)
            cell.font = styles['header_font']
            cell.fill = styles['dark_blue_fill']
            cell.alignment = styles['center_alignment']
        
        # Write journal entries data with error handling
        for _, row in journal_entries.iterrows():
            row_values = []
            for col in journal_entries.columns:
                value = row[col]
                try:
                    # Double-check the value is clean
                    if isinstance(value, str) and len(value) > 1000:
                        value = value[:1000]
                    row_values.append(value)
                except:
                    row_values.append("ERROR")
            je_sheet.append(row_values)
        
        # Apply date formatting to column C (Posted Date) in Journal Entries & Lines
        for row_num in range(3, je_sheet.max_row + 1):  # Start from row 3 (after headers)
            cell = je_sheet.cell(row=row_num, column=3)
            cell.number_format = 'M/D/YYYY'
        
        return je_sheet

    def _extract_balances_from_tb_data(self, account_id, begin_date, end_date):
        """Extract beginning and ending balances from TB-DATA sheet for a specific account"""
        if 'TB-DATA' not in self.source_data or self.source_data['TB-DATA'] is None:
            return 0.0, 0.0
        
        tb_data = self.source_data['TB-DATA']
        
        # Debug: Show the first few rows and column structure for the first account only
        if not hasattr(self, '_tb_data_debug_shown'):
            self.print_and_log("\nDEBUG: TB-DATA structure for balance extraction:")
            self.print_and_log(f"TB-DATA columns: {tb_data.columns.tolist()}")
            self.print_and_log("First 3 rows of TB-DATA:")
            for i in range(min(3, len(tb_data))):
                row_data = []
                for j in range(min(8, len(tb_data.columns))):  # Show first 8 columns
                    row_data.append(f"Col{j}: {tb_data.iloc[i, j]}")
                self.print_and_log(f"Row {i}: {', '.join(row_data)}")
            self._tb_data_debug_shown = True
        
        # Find the account in TB-DATA
        # Try different possible account ID column positions
        account_data = None
        account_col_found = None
        
        # Check columns B through D (indices 1-3) for Account ID
        for col_idx in [1, 2, 3]:
            try:
                if col_idx < len(tb_data.columns):
                    matching_rows = tb_data[tb_data.iloc[:, col_idx].astype(str) == str(account_id)]
                    if not matching_rows.empty:
                        account_data = matching_rows
                        account_col_found = col_idx
                        break
            except Exception:
                continue
        
        if account_data is None or account_data.empty:
            # Account not found in TB-DATA
            return 0.0, 0.0
        
        begin_balance = 0.0
        end_balance = 0.0
        matches_found = 0
        
        # Find balances matching our target dates
        for _, row in account_data.iterrows():
            try:
                # Column A (index 0) should contain fiscal month
                fiscal_month_val = row.iloc[0]
                if pd.isna(fiscal_month_val):
                    continue
                
                # Try to parse the fiscal month date
                fiscal_month = pd.to_datetime(fiscal_month_val)
                
                # Column E (index 4) is Starting Account Balance  
                starting_balance_val = row.iloc[4] if len(row) > 4 else 0.0
                starting_balance = 0.0
                if pd.notna(starting_balance_val):
                    try:
                        starting_balance = float(starting_balance_val)
                    except (ValueError, TypeError):
                        starting_balance = 0.0
                
                # Column G (index 6) is Ending Account Balance
                ending_balance_val = row.iloc[6] if len(row) > 6 else 0.0
                ending_balance = 0.0
                if pd.notna(ending_balance_val):
                    try:
                        ending_balance = float(ending_balance_val)
                    except (ValueError, TypeError):
                        ending_balance = 0.0
                
                # Check if fiscal month matches our target dates
                if fiscal_month.date() == begin_date.date():
                    begin_balance = ending_balance  # Changed from starting_balance to ending_balance
                    matches_found += 1
                    
                if fiscal_month.date() == end_date.date():
                    end_balance = ending_balance
                    matches_found += 1
                    
            except Exception as e:
                continue
        
        # Debug for first few accounts
        if matches_found > 0 and not hasattr(self, f'_balance_debug_{account_id}'):
            self.print_and_log(f"Found balances for account {account_id}: Begin={begin_balance}, End={end_balance}")
            setattr(self, f'_balance_debug_{account_id}', True)
        
        return begin_balance, end_balance

if __name__ == "__main__":
    parser = StrongboxParser()
    parser.run() 