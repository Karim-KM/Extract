# Importing Moduels
import pandas as pd
from pathlib import Path
import os,shutil,sys,argparse



def main():
    # Setting variables.
    shtname = 'List'
    move_dir = 'Output'
    # Setting arguments
    parser = argparse.ArgumentParser(description="""File Extractor based on workbook list
                                    Create Excel file with name 'List' and two columns 'Name' and 'Number' """,
                                    formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-p", "--path", help="Folder that contins the files")
    parser.add_argument("-ext", "--extention",
                        help="Folder that contins the files")
    parser.add_argument("-xl", "--excel", help="Folder that contins the files")
    parser.add_argument("--ignore-existing",
                        action="store_true", help="skip files that exist")
    args = parser.parse_args(args=None if sys.argv[1:] else ['--help'])
    fldr_dir = args.path # directory of the files.
    extract(args, fldr_dir, shtname, move_dir)


def extract(args, fldr_dir, shtname, move_dir):

    # list all accounts Python: Select Linter
    df = pd.read_excel(args.excel, shtname)
    # loop through folders and sub folders
    i = 0
    for index, row in df.iterrows():
        Name = row["Name"]
        Num = row["Number"]
        try:
            try:
                # Searching for the file by number
                filepath = max(Path(fldr_dir).rglob(
                    '*' + Num + '*.' + args.extention), key=os.path.getmtime)
                # todo:Validating if the file already exists in the same folder

                # Creating folder in same script folder if it's not exists.
                os.makedirs(move_dir, exist_ok=True)

                # Copy file to path
                shutil.copy(filepath, move_dir)
                print('File Found ' + Num, "|", Name)
                i += 1
            except:
                # Searching for the file by name
                filepath = max(Path(fldr_dir).rglob(
                    '*' + Name + '*.' + args.extention), key=os.path.getmtime)
                # todo:Validating if the file already exists in the same folder

                # Creating folder in same script folder if it's not exists.
                os.makedirs(move_dir, exist_ok=True)

                # Copy file to path
                shutil.copy(filepath, move_dir)
                print('File copied' + Num, "|", Name)
                i += 1
        except:
            print('file not found ' + Num, "|", Name)
    print("Total found files " + str(i))


if __name__ == "__main__":
    main()
