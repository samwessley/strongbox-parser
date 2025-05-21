import pandas as pd
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import calendar
import openpyxl
from copy import copy

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

    def update_status(self, message, progress=None):
        """Update the status message and progress bar"""
        self.status_label.config(text=message)
        if progress is not None:
            self.progress_bar['value'] = progress
        self.root.update()

    def get_file_paths(self):
        """Get source file and output directory using GUI"""
        self.update_status("Selecting source file...", 10)
        self.source_file = filedialog.askopenfilename(
            title="Select Strongbox File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not self.source_file:
            raise Exception("No source file selected")

        self.update_status("Selecting output directory...", 20)
        self.output_dir = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if not self.output_dir:
            raise Exception("No output directory selected")

    def get_date_range(self):
        """Get date range using GUI"""
        self.update_status("Enter date range...", 30)
        
        # Create date entry widgets
        start_frame = ttk.Frame(self.root, padding="10")
        start_frame.grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(start_frame, text="Start Date:").grid(row=0, column=0)
        start_date = ttk.Entry(start_frame)
        start_date.grid(row=0, column=1)
        start_date.insert(0, "MM/DD/YYYY")

        end_frame = ttk.Frame(self.root, padding="10")
        end_frame.grid(row=1, column=0, padx=5, pady=5)
        ttk.Label(end_frame, text="End Date:").grid(row=0, column=0)
        end_date = ttk.Entry(end_frame)
        end_date.grid(row=0, column=1)
        end_date.insert(0, "MM/DD/YYYY")

        # Add filename entry
        filename_frame = ttk.Frame(self.root, padding="10")
        filename_frame.grid(row=2, column=0, padx=5, pady=5)
        ttk.Label(filename_frame, text="Output Filename:").grid(row=0, column=0)
        filename_entry = ttk.Entry(filename_frame)
        filename_entry.grid(row=0, column=1)
        filename_entry.insert(0, "Audit_Sight_Output")

        def validate_inputs():
            try:
                start = datetime.strptime(start_date.get(), "%m/%d/%Y")
                end = datetime.strptime(end_date.get(), "%m/%d/%Y")
                if start > end:
                    messagebox.showerror("Error", "Start date must be before end date")
                    return
                self.start_date = start
                self.end_date = end
                self.output_filename = filename_entry.get()
                if not self.output_filename:
                    messagebox.showerror("Error", "Please enter an output filename")
                    return
                # Hide the input widgets
                start_frame.grid_remove()
                end_frame.grid_remove()
                filename_frame.grid_remove()
                validate_button.grid_remove()
                # Start processing
                self.process_data()
            except ValueError:
                messagebox.showerror("Error", "Invalid date format. Use MM/DD/YYYY")

        validate_button = ttk.Button(self.root, text="Start Processing", command=validate_inputs)
        validate_button.grid(row=3, column=0, pady=10)

    def load_source_data(self):
        """Load data from source file"""
        try:
            print("\nStarting to load source data...")
            self.update_status("Loading transaction data...", 40)
            
            # Load transaction data
            print("Opening Excel file...")
            excel_file_pd = pd.ExcelFile(self.source_file) # For initial sheet name listing and attempts
            print(f"Found sheets: {excel_file_pd.sheet_names}")
            
            processed_txn_sheets_count = 0
            # workbook_openpyxl = None # No longer needed as a shared instance

            for sheet_name in excel_file_pd.sheet_names:
                if sheet_name.startswith('TXN-FY'):
                    print(f"\nAttempting to process sheet: {sheet_name}")
                    df = None
                    try:
                        # Attempt 1: Standard pandas read with lenient options
                        print(f"Reading sheet data for {sheet_name} using pandas default...")
                        df = pd.read_excel(
                            excel_file_pd, # Use the pd.ExcelFile object for potentially better performance
                            sheet_name=sheet_name,
                            engine='openpyxl',
                            na_filter=False,
                            keep_default_na=False
                        )
                        print(f"Successfully read sheet {sheet_name} using pandas default.")

                    except Exception as e:
                        print(f"Pandas default read failed for {sheet_name}: {str(e)} (Type: {type(e)})")
                        print(f"Attempting fallback read for {sheet_name} using openpyxl cell by cell...")
                        
                        # Helper function for the openpyxl fallback logic
                        def _try_openpyxl_fallback(source_file_path, sheet_to_read, use_data_only):
                            print(f"FALLBACK ATTEMPT with data_only={use_data_only} for sheet '{sheet_to_read}'.")
                            fallback_workbook = None
                            df_from_this_attempt = None
                            try:
                                fallback_workbook = openpyxl.load_workbook(source_file_path, data_only=use_data_only, read_only=True)
                                if sheet_to_read not in fallback_workbook.sheetnames:
                                    print(f"FALLBACK WARNING: Sheet '{sheet_to_read}' not found in workbook (data_only={use_data_only}).")
                                    return None # Sheet not found in this mode

                                sheet_obj = fallback_workbook[sheet_to_read]
                                data_rows = []
                                header = []
                                if sheet_obj.max_row > 0:
                                    header = [cell.value for cell in sheet_obj[1]]
                                else:
                                    print(f"FALLBACK WARNING: Sheet '{sheet_to_read}' (data_only={use_data_only}) appears empty.")

                                problematic_cell_count = 0
                                for row_idx, row in enumerate(sheet_obj.iter_rows(min_row=2)):
                                    values = {}
                                    for col_idx, cell in enumerate(row):
                                        cell_value = None
                                        try:
                                            cell_value = cell.value
                                        except Exception as cell_err:
                                            print(f"FALLBACK CELL ERROR (data_only={use_data_only}): Sheet '{sheet_to_read}', Row {row_idx + 2}, Col {col_idx + 1}. Error: {cell_err}. Using None.")
                                            problematic_cell_count += 1
                                        if col_idx < len(header):
                                            if header[col_idx] is not None: values[header[col_idx]] = cell_value
                                            else: values[f"Unknown_Header_Col_{col_idx+1}"] = cell_value 
                                        else: values[f"Extra_Data_Col_{col_idx+1}"] = cell_value
                                    data_rows.append(values)
                                
                                df_from_this_attempt = pd.DataFrame(data_rows)
                                if df_from_this_attempt.empty and not data_rows and header: 
                                    df_from_this_attempt = pd.DataFrame(columns=header)
                                
                                if problematic_cell_count > 0:
                                    print(f"FALLBACK INFO (data_only={use_data_only}): Encountered {problematic_cell_count} problematic cell(s) in '{sheet_to_read}'.")
                                print(f"FALLBACK SUCCESS (data_only={use_data_only}): Successfully read and extracted data for '{sheet_to_read}'.")
                                return df_from_this_attempt

                            except Exception as current_attempt_err:
                                # This catches errors from load_workbook OR cell iteration for the current attempt
                                print(f"FALLBACK ERROR (data_only={use_data_only}) for '{sheet_to_read}': {current_attempt_err} (Type: {type(current_attempt_err)})")
                                raise current_attempt_err # Re-raise to be handled by the caller
                            finally:
                                if fallback_workbook:
                                    fallback_workbook.close()
                                    print(f"FALLBACK CLOSED (data_only={use_data_only}): Workbook for '{sheet_to_read}'.")
                        
                        # Main fallback execution flow
                        df = None
                        try:
                            df = _try_openpyxl_fallback(self.source_file, sheet_name, use_data_only=True)
                        except Exception as e_data_true_attempt:
                            if "Value must be either numerical or a string containing a wildcard" in str(e_data_true_attempt):
                                print(f"Fallback with data_only=True failed with target error. Trying data_only=False for {sheet_name}.")
                                try:
                                    df = _try_openpyxl_fallback(self.source_file, sheet_name, use_data_only=False)
                                except Exception as e_data_false_attempt:
                                    print(f"Fallback with data_only=False also failed for {sheet_name}: {str(e_data_false_attempt)}")
                                    # df remains None
                            else:
                                print(f"Fallback with data_only=True failed with an unexpected error for {sheet_name}, not retrying with data_only=False.")
                                # df remains None
                    
                    if df is None: # Check df, which would be None if all attempts failed
                        skipped_message = f"Sheet '{sheet_name}' could not be read by any method and will be SKIPPED. Please make sure to add these journal entries to the template manually."
                        print(skipped_message)
                        messagebox.showwarning("Sheet Read Error", skipped_message)
                        continue # To the next sheet in the outer loop

                    # Common processing for df (whether from pandas or openpyxl fallback)
                    print(f"Columns in {sheet_name}: {df.columns.tolist()}")
                    print(f"Number of rows in {sheet_name}: {len(df)}")
                    
                    print(f"Converting data types for {sheet_name}...")
                    # Convert string columns
                    for col in ['Transaction Id', 'Account Id', 'Memo', 'Doc/Ref No']:
                        if col in df.columns:
                            df[col] = df[col].astype(str)
                        else:
                            print(f"Warning: Column '{col}' not found in sheet {sheet_name}. Skipping conversion.")
                    
                    # Convert numeric columns
                    for col in ['Debit', 'Credit']:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        else:
                            print(f"Warning: Column '{col}' not found in sheet {sheet_name}. Skipping conversion.")
                    
                    # Convert Transaction Date to datetime
                    if 'Transaction Date' in df.columns:
                        print(f"INFO: 'Transaction Date' column found in {sheet_name}.")
                        print(f"Sample raw 'Transaction Date' values in {sheet_name} before pd.to_datetime:")
                        print(df['Transaction Date'].head(10) if len(df) > 0 else "Sheet is empty or has no dates")
                        
                        df['Transaction Date'] = pd.to_datetime(df['Transaction Date'], errors='coerce')
                        
                        print(f"Sample 'Transaction Date' values in {sheet_name} after pd.to_datetime (before NaT filter):")
                        print(df['Transaction Date'].head(10) if len(df) > 0 else "Sheet is empty or has no dates")
                        
                        # Count NaT values before filtering
                        nat_count = df['Transaction Date'].isnull().sum()
                        print(f"INFO: Found {nat_count} NaT (Not a Time) values in 'Transaction Date' for {sheet_name} after conversion.")
                        
                        # Filter out rows where date conversion failed (NaT)
                        df = df[pd.notna(df['Transaction Date'])] 
                        print(f"INFO: Rows in {sheet_name} after filtering NaT dates: {len(df)}")
                        
                        if len(df) > 0:
                            mask = (df['Transaction Date'] >= self.start_date) & (df['Transaction Date'] <= self.end_date)
                            df_filtered = df[mask]
                            print(f"INFO: Rows in {sheet_name} after applying date range [{self.start_date.strftime('%Y-%m-%d')} - {self.end_date.strftime('%Y-%m-%d')}]: {len(df_filtered)}")
                            if len(df_filtered) > 0:
                                self.source_data[sheet_name] = df_filtered
                                print(f"Successfully added filtered data from {sheet_name} to source_data.")
                                processed_txn_sheets_count += 1
                            else:
                                print(f"INFO: No data from {sheet_name} within the specified date range. Sheet will not be in final output.")
                        else:
                            print(f"INFO: No valid dates found in {sheet_name} after NaT filtering. Sheet will not be in final output.")
                    else:
                        print(f"WARNING: 'Transaction Date' column NOT FOUND in sheet {sheet_name}. Cannot filter by date.")
                        print(f"Skipping sheet {sheet_name} due to missing 'Transaction Date' column.")
            
            if processed_txn_sheets_count == 0 and any(s.startswith('TXN-FY') for s in excel_file_pd.sheet_names):
                message = "CRITICAL: No transaction (TXN-FY) sheets could be successfully processed after all attempts. The output may be incomplete or empty regarding journal entries. Please make sure to add all required journal entries to the template manually. Check the console logs for details on which sheets failed."
                print(message)
                messagebox.showerror("Critical Data Processing Error", message)
                self.update_status(message, 45)

            print("\nLoading trial balance data...")
            self.update_status("Loading trial balance data...", 50)
            
            # First read row 4 to get the dates
            print("Reading date row from TB sheet...")
            date_row = pd.read_excel(self.source_file, sheet_name='TB', header=None, nrows=1, skiprows=3)
            date_row = date_row.iloc[0]
            print(f"Date row values: {date_row.values}")
            
            # Convert dates to datetime and find the closest dates to our target dates
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
                                    print(f"Found date column: {date} at index {col_idx}")
                                    break
                                except ValueError:
                                    continue
                        else:
                            date = pd.to_datetime(value)
                            date_columns[date] = col_idx
                            print(f"Found date column: {date} at index {col_idx}")
                except Exception as e:
                    print(f"Error processing date at column {col_idx}: {str(e)}")
                    continue
            
            if not date_columns:
                raise Exception("No valid dates found in row 4 of the TB sheet. Please ensure dates are in a standard format (MM/DD/YYYY, YYYY-MM-DD, etc.)")
            
            # Find the closest dates to our target dates
            begin_date = (self.start_date - relativedelta(days=1)).replace(day=calendar.monthrange(
                (self.start_date - relativedelta(days=1)).year,
                (self.start_date - relativedelta(days=1)).month
            )[1])
            
            closest_begin_date = min(date_columns.keys(), key=lambda x: abs((x - begin_date).days))
            closest_end_date = min(date_columns.keys(), key=lambda x: abs((x - self.end_date).days))
            
            print(f"\nTarget dates:")
            print(f"Begin date: {begin_date}")
            print(f"End date: {self.end_date}")
            print(f"Closest begin date: {closest_begin_date}")
            print(f"Closest end date: {closest_end_date}")
            
            # Now read the actual data with data_only=True to evaluate formulas
            # Open the workbook
            try:
                wb = openpyxl.load_workbook(self.source_file, data_only=True)
                tb_sheet = wb['TB']
                
                # Find Financial Statement Classification Path column
                fin_statement_col = None
                for row_idx in range(1, 8):  # Check first few rows for headers
                    for col_idx in range(1, 25):  # Check first several columns
                        cell_value = tb_sheet.cell(row=row_idx, column=col_idx).value
                        if cell_value:
                            cell_text = str(cell_value).lower()
                            if 'financial statement classification' in cell_text or 'fin statement classification' in cell_text:
                                fin_statement_col = col_idx
                                print(f"Found Financial Statement Classification header in row {row_idx}, column {col_idx}: {cell_value}")
                                break
                    if fin_statement_col:
                        break
                
                if not fin_statement_col:
                    # Default to column C (3) as specified
                    fin_statement_col = 3
                    print(f"No Financial Statement Classification header found, using default column C (3)")
                
                # Print column indices for debugging
                print(f"\nColumn indices:")
                print(f"Begin date column index: {date_columns[closest_begin_date]}")
                print(f"End date column index: {date_columns[closest_end_date]}")
                print(f"Financial Statement Classification column index: {fin_statement_col}")
                
                data = []
                for row_idx in range(8, tb_sheet.max_row + 1):
                    try:
                        account_id = tb_sheet.cell(row=row_idx, column=4).value
                        if account_id is not None:
                            account_id = str(account_id).strip()
                            # Only skip rows with exact header matches, not all rows containing "account"
                            if account_id.lower() == "account id" or account_id.lower() == "account":
                                print(f"Skipping header row {row_idx} with account_id: {account_id}")
                                continue
                                
                            account_name = tb_sheet.cell(row=row_idx, column=6).value
                            begin_cell = tb_sheet.cell(row=row_idx, column=date_columns[closest_begin_date] + 1)
                            end_cell = tb_sheet.cell(row=row_idx, column=date_columns[closest_end_date] + 1)
                            
                            # Get Financial Statement Classification using the identified column
                            fin_statement_class = tb_sheet.cell(row=row_idx, column=fin_statement_col).value
                            
                            # Debug output for Financial Statement Classification for the first few rows
                            if row_idx < 12:
                                print(f"Row {row_idx}, Account: {account_id}, Financial Statement Classification (Column {fin_statement_col}): {fin_statement_class}")
                            
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
                        print(f"Error processing TB sheet row {row_idx}: {str(e)}. Skipping row.")
                        continue
                
                # Convert to DataFrame
                tb_data = pd.DataFrame(data)
                print("\nFirst few rows of TB data:")
                print(tb_data.head())
                
                self.source_data['TB'] = tb_data
                
            finally:
                # No need to explicitly close workbook with openpyxl as it doesn't keep file handles open
                pass

        except Exception as e:
            print(f"Error loading source data: {str(e)}")
            raise

    def create_journal_entries(self):
        """Create Journal Entries & Lines tab"""
        self.update_status("Creating journal entries...", 60)
        
        # Step 1: Get transaction sheets
        print("\nStep 1: Getting transaction sheets")
        transaction_sheets = {k: v for k, v in self.source_data.items() if k.startswith('TXN-FY')}
        print(f"Found sheets: {list(transaction_sheets.keys())}")
        
        # Step 2: Process each sheet individually
        print("\nStep 2: Processing individual sheets")
        processed_sheets = []
        for sheet_name, df in transaction_sheets.items():
            try:
                print(f"\nProcessing sheet: {sheet_name}")
                # Create a copy of the dataframe
                df_copy = df.copy()
                
                # Convert columns to appropriate types
                print("Converting columns...")
                df_copy['Transaction Id'] = df_copy['Transaction Id'].astype(str)
                df_copy['Memo'] = df_copy['Memo'].fillna('')
                df_copy['Doc/Ref No'] = df_copy['Doc/Ref No'].fillna('')
                df_copy['Account Id'] = df_copy['Account Id'].astype(str)
                
                # Convert numeric columns
                print("Converting numeric columns...")
                df_copy['Debit'] = pd.to_numeric(df_copy['Debit'], errors='coerce').fillna(0)
                df_copy['Credit'] = pd.to_numeric(df_copy['Credit'], errors='coerce').fillna(0)
                
                # Create the required columns
                print("Creating required columns...")
                processed_df = pd.DataFrame({
                    'Journal ID': df_copy['Transaction Id'],
                    'Journal Entry Description': df_copy['Doc/Ref No'],
                    'Posted Date': df_copy['Transaction Date'],
                    'Account': df_copy['Account Id'],
                    'Journal Line Description': df_copy['Memo'],
                    'Debit': df_copy['Debit'],
                    'Credit': df_copy['Credit']
                })
                
                processed_sheets.append(processed_df)
                print(f"Successfully processed sheet: {sheet_name}")
                
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {str(e)}")
                print("DataFrame info:")
                print(df.info())
                print("\nSample data:")
                print(df.head())
                raise
        
        # Step 3: Combine all processed sheets
        print("\nStep 3: Combining processed sheets")
        try:
            if not processed_sheets:
                raise Exception("No sheets were successfully processed")
            
            journal_entries = pd.concat(processed_sheets, ignore_index=True)
            print("Successfully combined all sheets")
            return journal_entries
            
        except Exception as e:
            print(f"Error combining sheets: {str(e)}")
            raise

    def create_trial_balance(self):
        """Create Comparative Trial Balances tab"""
        self.update_status("Creating trial balance...", 70)
        tb_data = self.source_data['TB']
        
        # Only filter out exact header matches, not all rows containing "account"
        tb_data = tb_data[~((tb_data['Account Id'].str.lower() == "account id") | 
                           (tb_data['Account Id'].str.lower() == "account"))]
        
        # Create new dataframe with required columns
        trial_balance = pd.DataFrame({
            'Account ID': tb_data['Account Id'],
            'Account Name': tb_data['Account Name'],
            'Beginning Balance \n(Prior Period Balance)': tb_data['Beginning Balance'],
            'Ending Balance': tb_data['Ending Balance'],
            'Account Type \n(see Mapping Categories tab)': '',
            'Account Mapping \n(see Mapping Categories tab)': '',
            'Account Description': tb_data['Financial Statement Classification']
        })
        
        # Debug info
        print("\nTrial Balance DataFrame - first few rows:")
        print(f"Column names: {trial_balance.columns.tolist()}")
        print("First 5 rows:")
        print(trial_balance.head(5))
        print(f"Total rows in trial balance: {len(trial_balance)}")
        
        # Calculate sum of beginning and ending balances
        begin_sum = trial_balance['Beginning Balance \n(Prior Period Balance)'].sum()
        end_sum = trial_balance['Ending Balance'].sum()
        print(f"Sum of Beginning Balances: {begin_sum}")
        print(f"Sum of Ending Balances: {end_sum}")
        
        # Only remove rows that are definitely headers (contain exactly "Account ID")
        headers_to_remove = []
        for idx, row in trial_balance.iterrows():
            account_id = str(row['Account ID']).lower() if pd.notna(row['Account ID']) else ""
            if account_id == "account id" or account_id == "account":
                headers_to_remove.append(idx)
                print(f"Removing header row: {row['Account ID']}")
        
        if headers_to_remove:
            trial_balance = trial_balance.drop(headers_to_remove)
        
        # Verify there are no empty account IDs but don't filter out other accounts
        trial_balance = trial_balance[trial_balance['Account ID'].notna() & (trial_balance['Account ID'] != '')]
        
        # Reset the index after filtering
        trial_balance = trial_balance.reset_index(drop=True)
        
        return trial_balance

    def create_output_file(self):
        """Create output Excel file with both tabs"""
        self.update_status("Creating output file...", 80)
        # Create output filename
        output_file = os.path.join(
            self.output_dir,
            f"{self.output_filename}_{self.start_date.strftime('%Y%m%d')}_{self.end_date.strftime('%Y%m%d')}.xlsx"
        )
        
        # Create a new workbook from scratch with no external links
        workbook = openpyxl.Workbook()
        
        # Define the tab order
        tab_order = [
            'Instructions', 
            'Data Validation Tests', 
            'Notes', 
            'Comparative Trial Balances',
            'Journal Entries & Lines', 
            'Banking Accts', 
            'Banking Txns', 
            'Mapping Categories'
        ]
        
        # Remove default sheet
        if 'Sheet' in workbook.sheetnames:
            default_sheet = workbook['Sheet']
            workbook.remove(default_sheet)
        
        # Create each sheet in the desired order
        for sheet_name in tab_order:
            workbook.create_sheet(sheet_name)
        
        # Set tab colors
        tab_colors = {
            'Instructions': '002060',
            'Data Validation Tests': '002060',
            'Notes': '002060',
            'Mapping Categories': '0070C0'
        }
        
        for sheet_name, color_code in tab_colors.items():
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                sheet.sheet_properties.tabColor = color_code
        
        # Load the trial balance data
        self.update_status("Creating trial balance...", 85)
        trial_balance = self.create_trial_balance()
        
        # Process the Journal Entries & Lines
        self.update_status("Creating journal entries...", 90)
        journal_entries = self.create_journal_entries()
        
        # Add accounts from journal entries that are missing from trial balance
        print("\nChecking for accounts in journal entries that are missing from trial balance...")
        self.update_status("Adding missing accounts to trial balance...", 87)
        
        # Get all account IDs in trial balance
        tb_account_ids = set(trial_balance['Account ID'].astype(str))
        
        # Get all unique accounts from journal entries with their account names
        je_accounts = {}
        if 'Account' in journal_entries.columns and len(journal_entries) > 0:
            # Get unique account ID and account name pairs from transaction data
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
                print(f"Found missing account in journal entries: {account_id} - {account_name}")
                missing_accounts.append({
                    'Account ID': account_id,
                    'Account Name': account_name,
                    'Beginning Balance \n(Prior Period Balance)': 0.0,
                    'Ending Balance': 0.0,
                    'Account Type \n(see Mapping Categories tab)': '',
                    'Account Mapping \n(see Mapping Categories tab)': '',
                    'Account Description': ''  # No Financial Statement Classification for missing accounts
                })
        
        # Add missing accounts to trial balance
        if missing_accounts:
            missing_df = pd.DataFrame(missing_accounts)
            trial_balance = pd.concat([trial_balance, missing_df], ignore_index=True)
            print(f"Added {len(missing_accounts)} missing accounts to trial balance")
        
        # Process the TB sheet
        tb_sheet = workbook['Comparative Trial Balances']
        
        # Add required/optional labels to row 1
        for col in range(1, 6):  # Columns A-E
            tb_sheet.cell(row=1, column=col, value='Required')
        for col in range(6, 8):  # Columns F-G
            tb_sheet.cell(row=1, column=col, value='Optional')
        
        # Add column headers to row 2
        for col_idx, column_name in enumerate(trial_balance.columns, 1):
            tb_sheet.cell(row=2, column=col_idx, value=column_name)
        
        # Add data starting at row 3
        for row_idx, row in enumerate(trial_balance.iterrows(), 3):
            for col_idx, col_name in enumerate(trial_balance.columns, 1):
                tb_sheet.cell(row=row_idx, column=col_idx, value=row[1][col_name])
        
        # Set column widths for Comparative Trial Balances
        column_widths = {
            'A': 148/7,  # Approximate conversion from pixels to Excel units
            'B': 294/7,
            'C': 154/7,
            'D': 154/7,
            'E': 294/7,
            'F': 302/7,
            'G': 162/7
        }
        
        for col_letter, width in column_widths.items():
            tb_sheet.column_dimensions[col_letter].width = width
        
        # Process the Journal Entries & Lines sheet
        je_sheet = workbook['Journal Entries & Lines']
        
        # Add required/optional labels to row 1
        required_cols = [1, 3, 4, 6, 7]  # A, C, D, F, G
        optional_cols = [2, 5]  # B, E
        for col in required_cols:
            je_sheet.cell(row=1, column=col, value='Required')
        for col in optional_cols:
            je_sheet.cell(row=1, column=col, value='Optional')
        
        # Add column headers to row 2
        for col_idx, column_name in enumerate(journal_entries.columns, 1):
            je_sheet.cell(row=2, column=col_idx, value=column_name)
        
        # Add data starting at row 3
        for row_idx, row in enumerate(journal_entries.iterrows(), 3):
            for col_idx, col_name in enumerate(journal_entries.columns, 1):
                je_sheet.cell(row=row_idx, column=col_idx, value=row[1][col_name])
        
        # Copy templates for the other tabs - SAFELY with no external links
        self.update_status("Copying templates for other tabs...", 95)
        template_wb = openpyxl.load_workbook('Audit Sight Template.xlsx', data_only=True, keep_links=False)
        
        # Copy content from template without formulas or links
        for sheet_name in tab_order:
            if sheet_name not in ['Comparative Trial Balances', 'Journal Entries & Lines']:
                if sheet_name in template_wb.sheetnames:
                    template_sheet = template_wb[sheet_name]
                    output_sheet = workbook[sheet_name]
                    
                    # Copy content but not formulas, links or connections
                    for row in template_sheet.rows:
                        for cell in row:
                            # Get the cell value only, ignoring formulas
                            cell_value = cell.value
                            
                            # Avoid copying hyperlinks or external references
                            if isinstance(cell_value, str) and (cell_value.startswith('http') or cell_value.startswith('=')):
                                if cell_value.startswith('='):
                                    # Replace formula with its value or empty string
                                    cell_value = ""
                            
                            # Set the cell value directly without links
                            output_cell = output_sheet.cell(row=cell.row, column=cell.column, value=cell_value)
                            
                            # Copy only safe styling elements
                            if cell.has_style:
                                output_cell.font = copy(cell.font)
                                output_cell.border = copy(cell.border)
                                output_cell.fill = copy(cell.fill)
                                output_cell.number_format = cell.number_format
                                output_cell.alignment = copy(cell.alignment)
        
        # Apply safe header formatting
        self.update_status("Applying formatting...", 98)
        if 'Comparative Trial Balances' in template_wb.sheetnames:
            template_tb = template_wb['Comparative Trial Balances']
            for col in range(1, 8):  # Columns A-G
                if col <= len(trial_balance.columns):
                    template_cell = template_tb.cell(row=2, column=col)
                    tb_cell = tb_sheet.cell(row=2, column=col)
                    if template_cell.has_style:
                        tb_cell.font = copy(template_cell.font)
                        tb_cell.fill = copy(template_cell.fill)
        
        if 'Journal Entries & Lines' in template_wb.sheetnames:
            template_je = template_wb['Journal Entries & Lines']
            for col in range(1, 8):  # Columns A-G
                if col <= len(journal_entries.columns):
                    template_cell = template_je.cell(row=2, column=col)
                    je_cell = je_sheet.cell(row=2, column=col)
                    if template_cell.has_style:
                        je_cell.font = copy(template_cell.font)
                        je_cell.fill = copy(template_cell.fill)
        
        # Ensure no external links exist
        if hasattr(workbook, 'external_links'):
            workbook.external_links.clear()
            
        # Turn off external data connections
        workbook.properties.externalReferences = False
        
        # Save the workbook with properties to avoid external content warnings
        self.update_status("Saving workbook...", 99)
        workbook.save(output_file)
        
        return output_file

    def process_data(self):
        """Process the data and create output file"""
        try:
            self.load_source_data()
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
            self.root.geometry("600x400")

            # Create status label
            self.status_label = ttk.Label(self.root, text="Initializing...", wraplength=500)
            self.status_label.grid(row=4, column=0, padx=10, pady=10)

            # Create progress bar
            self.progress_bar = ttk.Progressbar(self.root, length=400, mode='determinate')
            self.progress_bar.grid(row=5, column=0, padx=10, pady=10)

            self.get_file_paths()
            self.get_date_range()
            
            # Start the main event loop
            self.root.mainloop()
        except Exception as e:
            if self.root:
                self.update_status(f"Error: {str(e)}", 0)
                messagebox.showerror("Error", str(e))
            else:
                messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    parser = StrongboxParser()
    parser.run() 