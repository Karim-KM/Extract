# Extract Files from folder using excel workbook
## Requirements
- pandas 
- shutil

## usage
    usage: Extract.py [-h] [-p PATH] [-ext EXTENTION] [-xl EXCEL] [--ignore-existing]

    File Extractor based on workbook list Create Excel file with name 'List' and two columns 'Name' and 'Number'

    options:
    -h, --help            show this help message and exit
    -p PATH, --path PATH  Folder that contins the files (default: None)
    -ext EXTENTION, --extention EXTENTION
                            Folder that contins the files (default: None)
    -xl EXCEL, --excel EXCEL
                            Folder that contins the files (default: None)
    --ignore-existing     skip files that exist (default: False)

## To-do's
- [ ] Validate on files that exists.
Hope this script proves useful to somebody else!
