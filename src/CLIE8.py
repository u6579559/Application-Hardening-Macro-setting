import sys
import argparse
import os.path
import time
import xml.etree.ElementTree as ET

# Hacky solution as src could not be found in the path
sys.path.insert(0, '../')
from main import Report

# test_string => python CLIE8.py --gpo "C:\Users\Liam\Desktop\Tech Launcher\Sem1 - 2020\techlauncher-essential-8\tests\GroupPolicies\ChromeAndIE.xml" -o "D:\TechLauncher Dev"
# test string 2 => python CLIE8.py --gpo "C:\Users\Liam\Desktop\Tech Launcher\Sem1 - 2020\techlauncher-essential-8\tests\GroupPolicies\ChromeAndIE.xml" "C:\Users\Liam\Desktop\Tech Launcher\Sem1 - 2020\techlauncher-essential-8\tests\GroupPolicies\3BrowserReport.xml" -o "D:\TechLauncher Dev"

'''
To run the command line interface the command will be "python CLIE8.py --gpo 1.xml 2.xml 3.xml 4.xml -o pathtodirectory -n nameoffile"
We can later extend this if necessary to include other files or features (by appending a new option/arguments)
-n is an optional parameter.
'''
# Command Line Interface
_opts = [opt for opt in sys.argv[1:] if opt.startswith("-")]
_args = [arg for arg in sys.argv[1:] if not arg.startswith("-")]


def clean_file_name(file_name):
    return file_name.replace(":", "_")


def print_command():
    print('Usage: python CLIE8.py --gpo 1.xml 2.xml 3.xml 4.xml -o pathtodirectory -n "Name of file"')
    print('Please be aware that directories with spaces will be treated as separate arguments unless enclosed with ""')
    print('For example C:/Users/Some name will be treated as [C:/Users/Some, name]')
    print('But "C:/Users/Some name" will be treated as [C:/Users/Some name]')


def is_valid(filename):
    is_valid = True
    try:
        ET.parse(filename)
        is_valid = True
    except Exception as e:
        is_valid = False
    return is_valid


def main():
    if "--gpo" not in _opts:
        print("There was no gpo option specified")
        print_command()
        return
    if "-o" not in _opts:
        print("No output directory specified")
        print_command()
        return
    if "-n" not in _opts:
        print("No name specified, defaulting to Essential8Report " + clean_file_name(str(time.asctime(time.localtime()))))

    parser = argparse.ArgumentParser(description='Create a report on the Essential Eight using GPO files.')
    parser.add_argument('--gpo', default=[], nargs='*', help='Input arguments (GPO files)')
    parser.add_argument('-o', help="Output file directory")
    parser.add_argument('-n', default='Essential8Report ' + time.asctime(time.localtime()), help='Name of the report')
    args = vars(parser.parse_args())
    print(args.get('gpo'))
    gpo_list = args.get('gpo')
    save_directory = args.get('o')
    existing_gpo = []
    # print(args.get('o'))
    report_name = clean_file_name(args.get('n'))
    print(report_name)
    '''Getting all of the valid GPO's provided'''
    for gpo in gpo_list:
        if os.path.isfile(gpo):
            if is_valid(gpo):
                existing_gpo.append(gpo)
            else:
                print("The file " + gpo + " exists but is not a valid xml file")
        else:
            print("The file " + gpo + " does not exist.")

    if len(existing_gpo) == 0:
        print("No valid GPO files were provided, exiting.")
        return

    '''Checking to see if the directory provided is valid'''
    if not os.path.isdir(save_directory):
        print("The directory you have specified to save to '" + save_directory + "' is not accessible, exiting")
        return

    ext = 'docx'
    report = Report(existing_gpo)
    full_path = os.path.join(save_directory, str(report_name) + "." + ext)
    report.export_report(full_path)
    # print('valid xml: ' + str(existing_gpo))


if __name__ == "__main__":
    main()
