from openpyxl import Workbook
from openpyxl import load_workbook
import json
import yaml 
import sys
import zipfile
import pathlib
from lxml import etree 
import os
import base64
import shutil
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML
import gzip
from sys import exit
import csv
import time

# Global setting to add an extra level of compression to binaries internal to the format
setting_compress_binary = False

"""
Configuration function for creating a temporary folder, which
will store the output of the Excel workbooks as they are 
decompressed. 
"""
def set_temp_folder():
    return 'output_dir'

# Helper function for testing to compress files 
def compress_file(input_path):
    with open(input_path, 'rb') as f_in:
        with gzip.open('{0}.gz'.format(input_path), 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)

# Helper function for testing file decompression
def decompress_file(input_path, output_path):
    with gzip.open(input_path, 'rb') as f_in:
        with open(output_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)

# Compress a file and place the output in the same location as the input file 
def compress_file_in_place(input_path):
    output_path = None
    with open(input_path, 'rb') as f_in:
        output_path = '{0}.gz'.format(input_path)
        with gzip.open(output_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)
    if os.path.exists(input_path):
        os.remove(input_path)
    os.rename(output_path, input_path)

# Decompress a file and place the output in the same location as the input file 
def decompress_file_in_place(input_path):
    output_path = None
    with gzip.open(input_path, 'rb') as f_in:
        output_path = '{0}.gz'.format(input_path)
        with open(output_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)
    if os.path.exists(input_path):
        os.remove(input_path)
    os.rename(output_path, input_path)

"""
Confirm if VBA code is actually present within a module. Modules
will be returned by default for all XLSM workbooks, but we only
want to output a separate VBA file and VBA section within the YML
file if we find VBA within the module.
"""
def screen_for_vba(vba_code):
    contains_vba = False
    for line in vba_code.splitlines():
        if not line.strip().startswith('Attribute'):
            contains_vba = True
    return contains_vba
    
"""
We want to only delete a workbook file and recreate it if the user 
does not already have it open in Excel.  We attempt to delete the 
file and will return an error if permission is not granted. This 
will cause the script to skip over creating the file in question 
and return an error to the log and to standard output.
"""
def delete_file_safe(input_file):
    if not os.path.exists(input_file):
        return True
    try:
        os.remove(input_file)
        return True
    except OSError:
        return False

"""
Function used to convert a workbook either XLSX or XLSM into YML.
VBA code is parsed and separated into a secondary file if that 
option has been selected in the configuration file. This would be
called by a git pre-commit hook to convert spreadsheets before they
are stored in the repo.
"""
def write_workbook_to_yml(workbook_path, vba_convert):

    fpath, extension  = os.path.splitext(workbook_path)
    output_path = set_temp_folder()

    ymlFilename = '{0}{1}.yml'.format(fpath, extension)
    # Remove previously created YML file     
    if not delete_file_safe(ymlFilename):
        return False

    vbaFilename = '{0}.vba'.format(fpath)
    vbaCodeList = []
    # Remove previously created VBA file    
    if vba_convert: 
        if not delete_file_safe(vbaFilename):
            return False

    with zipfile.ZipFile(workbook_path,"r") as zip_ref:
        zip_ref.extractall(output_path)

    directory = pathlib.Path(output_path)

    with open(ymlFilename, 'w', encoding="utf-8") as output_yml:
        output_yml.write('options: ' + "\n")
        output_yml.write('  extension: "{0}"\n'.format(extension))

        if extension in ['.xlsm']:
            vbaparser = VBA_Parser(workbook_path)
            if vbaparser.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
                    if screen_for_vba(vba_code):
                        vbaC1 = 'vba: ' + "\n"
                        vbaC2 = '  filename: "{0}"\n'.format(vba_filename)
                        vbaC3 = '  code: |' + "\n"
                        output_yml.write(vbaC1)
                        output_yml.write(vbaC2)
                        output_yml.write(vbaC3)
                        vbaCodeList += [vbaC1, vbaC2, vbaC3]

                        for line in vba_code.splitlines():
                            vbaC4 = '    {0}\n'.format(line)
                            output_yml.write(vbaC4)
                            vbaCodeList += [vbaC4]

        for file_path in directory.rglob("*"):
            file_name, file_extension = os.path.splitext(file_path)
            if file_path.is_dir():
                continue

            if file_extension in ['.xml', '.vml', '.rels'] or file_path.stem in ['.rels']:
                
                temp = etree.parse(file_path) 
                new_xml = etree.tostring(temp, pretty_print = True, encoding = str) # https://www.geeksforgeeks.org/pretty-printing-xml-in-python/
                
                output_yml.write(str(file_path) + ': |' + "\n")
                for line in new_xml.splitlines():
                    output_yml.write('  ' + line.strip() + "\n")
            else:
                if setting_compress_binary:
                    compress_file_in_place(file_path)

                with open(file_path, 'rb') as binary_file:
                    binary_file_data = binary_file.read()
                    base64_encoded_data = base64.b64encode(binary_file_data)
                    base64_message = base64_encoded_data.decode('utf-8')

                output_yml.write(str(file_path) + ': |' + "\n")
                output_yml.write('  ' + base64_message + "\n")

    # Cleanup the temporary directory
    shutil.rmtree(output_path)  

    if vba_convert: 
        with open(vbaFilename, 'w', encoding="utf-8") as output_vba:
            for codeLine in vbaCodeList:
                output_vba.write(codeLine)

    return True


"""
Function used to convert YML back into a workbook either either XLSX or XLSM 
as specified by the YML file. This would be called by a git post-checkout 
hook to convert spreadsheets before they are stored in the repo.
"""
def convert_yml_to_workbook(inputFile):
    temp_folder = set_temp_folder()
    fpath0, extension0 = os.path.splitext(inputFile)
    fpath, extension = os.path.splitext(fpath0)

    with open(inputFile, encoding="utf-8") as file:
        inputYML = yaml.safe_load(file)

    outputExtension = inputYML['options']['extension']
    outputFilePath = '{0}{1}'.format(fpath, outputExtension)

    if not delete_file_safe(outputFilePath):
        return False

    for key in inputYML:
        if key.startswith(temp_folder):
            
            path_list = key.split(os.path.sep)[1:]
            new_path = os.path.join(temp_folder, *path_list) 
            
            os.makedirs(os.path.dirname(new_path), exist_ok=True)

            if str(new_path).endswith(tuple(['.xml', '.vml', '.rels'])):
                with open(new_path, "w", encoding="utf-8") as f:
                    f.write(inputYML[key])
            else:
                with open(new_path, 'wb') as binary_file:
                    encoded = inputYML[key].encode('utf-8')
                    file_bytes = base64.b64decode(encoded)
                    binary_file.write(file_bytes)

                if setting_compress_binary:
                    decompress_file_in_place(new_path)
            
    # https://realpython.com/python-zipfile/#building-a-zip-file-from-a-directory
    directory = pathlib.Path(temp_folder + '/')
    with zipfile.ZipFile(outputFilePath, mode="w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for file_path in directory.rglob("*"):
            archive.write(
                file_path,
                arcname=file_path.relative_to(directory)
            )

    # Cleanup the temporary directory
    shutil.rmtree(temp_folder)  

    return True

"""    
This function confirms if the type of file that we have found when scanning 
across the directories is a file that we can convert, and that the settings
specified in the configuration file request that we convert the file
"""
def validate_file_path(conversion_type, setting_convert_xlsx, setting_convert_xlsm, excluded_list, folder, file):
    for x in excluded_list:
        if folder.startswith(x):
            return False

    xlsx_ext = '.xlsx'
    xlsm_ext = '.xlsm'
    if conversion_type == 'convert_to_excel':
        xlsx_ext = '.xlsx.yml'
        xlsm_ext = '.xlsm.yml'

    if file.endswith(xlsx_ext) and setting_convert_xlsx:
        return True
    if file.endswith(xlsm_ext) and setting_convert_xlsm:
        return True
    
    return False


"""
The main loop of the application after we have parsed input arguments.
Here we read in the settings specified in the configuration file, and 
then loop across all files in the repo to find files that need to be 
converted in format. If we cannot convert some files due to permission
locks, we report this to the calling function as an error. The results 
of individual file operations are reported through a logging file if 
this is enabled in the configuration file.
"""
def entry_point(conversion_type):
    with open('version_sheet_settings.yml', encoding="utf-8") as file:
        sheetSettings = yaml.safe_load(file)

    # Read configuration settings
    setting_enabled = sheetSettings['options']['enabled']
    setting_convert_xlsx = sheetSettings['options']['convert_xlsx']
    setting_convert_xlsm = sheetSettings['options']['convert_xlsm']
    setting_convert_vba = sheetSettings['options']['convert_vba_separate_file']
    setting_enable_logging = sheetSettings['options']['enable_logging']
    setting_logfile = sheetSettings['options']['logfile']
    setting_exclude_directories = sheetSettings['exclude_directories']

    if not setting_enabled:
        return 0

    # Append the root directory to the path we retrieved from settings
    rootdir = '.'
    exclude_dir = [os.path.join(rootdir, x) for x in setting_exclude_directories]

    if setting_enable_logging:
        logfile = open(setting_logfile, 'a', encoding="utf-8")

    convertFailureCount = 0

    for subdir, dirs, files in os.walk(rootdir):
        for file in files:
            if validate_file_path(conversion_type, setting_convert_xlsx, setting_convert_xlsm, exclude_dir, subdir, file):
                
                filepath = os.path.join(subdir, file)
                print(filepath)

                start_time = time.time()
                if conversion_type == 'convert_to_excel':
                    convertResult = convert_yml_to_workbook(filepath)
                else:
                    convertResult = write_workbook_to_yml(filepath, setting_convert_vba)

                end_time = time.time()
                elapsed_time = round(end_time - start_time,3)
                exec_time = time.ctime()
                convertResultString = 'Success' if convertResult else 'Failure'
                log_output = '{0} | {1} | {2} | Execution time: {3} seconds | {4}\n'.format(exec_time, conversion_type, convertResultString, elapsed_time, filepath)

                if not convertResult:
                    convertFailureCount += 1

                if setting_enable_logging:
                    logfile.write(log_output)

    if setting_enable_logging:
        logfile.close()

    if convertFailureCount > 0:
        print('Could not convert {0} locked files.'.format(convertFailureCount))
        return 1
    return 0


""" 
The global entry point for the application code is specified here. We parse 
the input arguments and determnine the requested operation. This is then passed
to the downstream functions for execution. If an error code is returned (indicated
by a sys.exit code other than zero) then the git hook will also produce an error 
and display an error message to the user. 
"""
if len(sys.argv) != 2:
    sys.exit(1)

input_arg = sys.argv[1]
if input_arg not in ['convert_to_excel', 'convert_to_yml']:
    print('Invalid arguments')
    sys.exit(1)

result = entry_point(input_arg)
sys.exit(result)




