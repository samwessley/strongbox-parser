# Strongbox Parser - Installation & Usage Guide

## 📥 Download & Installation

### For Windows Users (PC)

1. **Download the executable:**
   - Go to the [Releases page](https://github.com/samwessley/strongbox-parser/releases)
   - Download the latest `StrongboxParser-Windows.exe` file
   - Or download from the latest [Actions build artifacts](https://github.com/samwessley/strongbox-parser/actions)

2. **Installation:**
   - No installation required! Just download and run
   - Save the `.exe` file to a folder of your choice (e.g., Desktop, Documents)
   - Double-click `StrongboxParser-Windows.exe` to run

3. **First Run:**
   - Windows may show a security warning (this is normal for unsigned executables)
   - Click "More info" → "Run anyway" to proceed
   - The application will open with a GUI interface

### For Mac Users

1. **Download the executable:**
   - Go to the [Releases page](https://github.com/samwessley/strongbox-parser/releases)  
   - Download the latest `StrongboxParser-macOS` file

2. **Installation:**
   - No installation required! Just download and run
   - Save the file to Applications or Desktop
   - Right-click the file → "Open" (first time only due to security settings)

## 🚀 How to Use

1. **Launch the Application:**
   - Double-click the executable file
   - A window will open with "Strongbox Parser" title

2. **Process Your Files:**
   - Click "Select Source File" button
   - Choose your Strongbox Excel file (.xlsx)
   - Choose where to save the output file
   - The tool will automatically process your data

3. **Monitor Progress:**
   - Watch the console output box for real-time status updates
   - Progress bar shows current processing stage
   - Status messages provide detailed feedback

4. **Results:**
   - Output Excel file will be created with properly formatted tabs:
     - Comparative Trial Balances
     - Journal Entries & Lines  
     - Instructions, Data Validation Tests, Notes
     - Banking Accts, Banking Txn, Mapping Categories

## 📋 Requirements

- **Windows:** Windows 10 or later
- **Mac:** macOS 10.13 or later
- **Input File:** Strongbox Excel file with TB and TXN-FY sheets
- **No Python installation required** - executables include everything needed

## 🔧 Troubleshooting

### Windows Issues

**"Windows protected your PC" warning:**
- This is normal for unsigned executables
- Click "More info" → "Run anyway"
- Consider adding the executable to Windows Defender exclusions

**Excel file won't open:**
- Ensure you have Excel or Excel viewer installed
- Try opening the generated CSV files if Excel creation fails

### Mac Issues

**"Cannot open because developer cannot be verified":**
- Right-click the file → "Open" → "Open" again
- Or go to System Preferences → Security & Privacy → General → "Open Anyway"

**Permission denied:**
- Open Terminal and run: `chmod +x /path/to/StrongboxParser-macOS`
- Then double-click to run

### General Issues

**Large file processing is slow:**
- This is normal for files with hundreds of thousands of transactions
- The console output shows progress - the tool is still working
- Processing 700K+ transactions may take 5-10 minutes

**Missing data in output:**
- Check console output for any skipped sheets or errors
- Ensure your Strongbox file has properly formatted TB and TXN-FY sheets

## 📞 Support

If you encounter issues:
1. Check the console output for error messages
2. Ensure your Strongbox file follows the expected format
3. Contact the development team with:
   - Screenshots of any error messages
   - Description of your input file structure
   - Operating system version

## 🔄 Updates

New versions are automatically built when updates are made to the code:
- Check the [Releases page](https://github.com/samwessley/strongbox-parser/releases) for new versions
- Download the latest version to get bug fixes and new features
- No need to uninstall old versions - just replace the executable file 