# XLSX Merge Framework

A utility for managing Excel workbooks in version control by unpacking them into XML components and repacking them back into XLSX files. This framework makes it easier to track changes in Excel files and manage merge conflicts.

## Setup

1. Clone the repository:
```bash
git clone https://github.com/Shubham-SK/xlsx-merge-framework.git
cd xlsx-merge-framework
```

2. Create and activate a virtual environment:
```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# On macOS/Linux:
source .venv/bin/activate
# On Windows:
.\.venv\Scripts\activate
```

## Directory Structure

```
xlsx-merge-framework/
├── assets/
│   ├── workbooks/     # Store your .xlsx files here
│   └── xmls/          # Unpacked XML components will be stored here
└── src/
    ├── xlsx_manager.py    # Main utility for packing/unpacking
    └── xlsx_merger.py     # Merger utility (coming soon)
```

## Usage

### Basic Workflow

1. **Place your Excel file** in the `assets/workbooks/` directory.

2. **Unpack the Excel file** before making changes or committing:
```bash
python src/xlsx_manager.py unpack your_workbook.xlsx
```
This will create a directory in `assets/xmls/` containing the unpacked XML components.

3. **Commit the XML files** to version control:
```bash
git add assets/xmls/your_workbook/
git commit -m "Add/Update your_workbook components"
git push
```

4. **After pulling changes**, pack the XML files back into an XLSX:
```bash
python src/xlsx_manager.py pack your_workbook
```
This will create/update the XLSX file in `assets/workbooks/`.

### Advanced Usage

#### Custom Output Names
You can specify a custom output name when packing:
```bash
python src/xlsx_manager.py pack your_workbook -o new_name.xlsx
```

#### Custom Directories
Override default directories if needed:
```bash
python src/xlsx_manager.py unpack workbook.xlsx --workbooks-dir custom/workbooks --xmls-dir custom/xmls
```

### CLI Reference

```
usage: xlsx_manager.py [-h] [-o OUTPUT] [--workbooks-dir WORKBOOKS_DIR]
                      [--xmls-dir XMLS_DIR]
                      {pack,unpack} name

XLSX file unpacker and packer utility

positional arguments:
  {pack,unpack}         Action to perform: pack or unpack XLSX file
  name                  For unpack: name of the XLSX file in workbooks directory.
                       For pack: name of the directory in xmls directory

optional arguments:
  -h, --help           show this help message and exit
  -o OUTPUT            Output name (optional). For pack: output XLSX filename.
                      For unpack: output directory name
  --workbooks-dir DIR  Directory containing XLSX files (default: assets/workbooks)
  --xmls-dir DIR       Directory for XML files (default: assets/xmls)
```

## Best Practices

1. **Always unpack before committing**: This ensures changes to Excel files are tracked in a diff-friendly format.

2. **Pack after pulling**: When pulling changes from the repository, pack the XML files to get the updated XLSX.

3. **Version Control**:
   - Commit the XML files in `assets/xmls/`
   - Consider adding `assets/workbooks/*.xlsx` to `.gitignore` to avoid conflicts
   - Keep the XML files as the source of truth

4. **XML Formatting**: All XML files are automatically formatted for better readability and version control diff tracking.

## Features

- ✅ Unpack XLSX files into XML components
- ✅ Pack XML components back into valid XLSX files
- ✅ Automatic XML formatting for better diff tracking
- ✅ Command-line interface
- ✅ Custom output names and directory support
- 🔄 Overwrite protection and error handling
- 📝 Formatted XML for better version control

## Contributing

Contributions are welcome! Please feel free to submit pull requests, create issues, or suggest improvements.

## License

This project is licensed under the MIT License - see the LICENSE file for details.