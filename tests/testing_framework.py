import os 
import pyautogui
import time
import shutil
import subprocess
import csv
import hashlib

"""
Hardcoded name of the local folder which contains the repo. This 
will be accessed by the script during testing to convert files and 
and commit them into the repo for size comparison. 
"""
def name_repo_folder():
	return 'pythonxlsx_testing2'

"""
The authentication key which will be used during testing. This should 
be replaced with either a valid password or the repo access token.
"""
def name_repo_password():
	return ''

"""
The location of the hosted repo which will be cloned to the local folder.
"""
def name_repo():
	return 'https://replace_this@bitbucket.org/replace_this/replace_this.git'

"""
A sequence of commands which will clone the repo in it's current state into 
the local folder
"""
def clone_repo():
	time.sleep(5)
	pyautogui.typewrite('git clone {0}'.format(name_repo()))
	pyautogui.press("enter")
	time.sleep(2)
	pyautogui.typewrite(name_repo_password())
	pyautogui.press("enter")
	time.sleep(5)

"""
A sequence of commands which will repack the repo to save space and achieve
the optimial levels of compression
"""
def git_repack():
	time.sleep(1)
	pyautogui.typewrite('git repack')
	pyautogui.press("enter")
	time.sleep(5)
	pyautogui.typewrite('git prune-packed')
	pyautogui.press("enter")
	time.sleep(5)
	pyautogui.typewrite('git gc --aggressive')
	pyautogui.press("enter")
	time.sleep(5)

"""
A sequence of commands which will commit changed files to the local copy of the 
repo along with the provided commit message
"""
def git_commit(commit_message):
	time.sleep(1)
	pyautogui.typewrite('git add -A')
	pyautogui.press("enter")
	time.sleep(3)
	pyautogui.typewrite('git commit -m "{0}"'.format(commit_message))
	pyautogui.press("enter")
	time.sleep(5)

"""
A sequence of commands which will push the updated version of the repo to the
cloud hosted repo. 
"""
def git_push():
	time.sleep(1)
	pyautogui.typewrite('git push')
	pyautogui.press("enter")
	time.sleep(2)
	pyautogui.typewrite(name_repo_password())
	pyautogui.press("enter")
	time.sleep(5)

"""
A sequence of commands which changes the current directly of the terminal
session into the repo folder
"""
def cd_into_repo():
	pyautogui.typewrite('cd {0}'.format(name_repo_folder()))
	pyautogui.press("enter")
	time.sleep(1)

"""
A sequence of commands which exits the terminal session from the repo folder
"""
def cd_out_repo():
	pyautogui.typewrite('cd ../')
	pyautogui.press("enter")
	time.sleep(1)

"""
A sequence of commands which deletes the local repo folder and all contents
"""
def wipe_repo():
	dirpath = name_repo_folder()
	delete_folder(dirpath)
	time.sleep(1)

"""
Implements the deletion of the repo folder and all subdirectories 
"""
def delete_folder(inputFolder):
	if os.path.exists(inputFolder) and os.path.isdir(inputFolder):
		os.system('rmdir {0} /S /Q'.format(inputFolder))

"""
Copys files from the specified source directory into the repo folder
"""
def copy_into_repo(sDir):
	dDir = name_repo_folder()
	for item in os.listdir(sDir):
		sItem = os.path.join(sDir, item)
		dItem = os.path.join(dDir, item)
		if not os.path.isdir(sItem):
			shutil.copy(sItem, dItem)

"""
Copies all files between a specified source folder and the destination
"""
def copy_between_folders(sFolder, dFolder):
	for item in os.listdir(sFolder):
		sItem = os.path.join(sFolder, item)
		dItem = os.path.join(dFolder, item)
		if not os.path.isdir(sItem):
			shutil.copy(sItem, dItem)

"""
Runs an external command to determine the size of the repo on a windows system
"""
def record_size_repo():
	repoDir = name_repo_folder()
	resultBytes = subprocess.check_output('folder-size.cmd {0}'.format(repoDir));
	result = resultBytes.decode("utf-8")
	result = result.strip()
	repo_size_bytes = result.split(' ')[-2].replace(',', '')
	return repo_size_bytes

"""
Runs the versioning script to convert files in the source folder into YML format
"""
def convert_to_yml(sFolder): # Converts a folder from Excel to YML
	for item in os.listdir(sFolder):
		if not os.path.isdir(item):
			cmd1 = 'version_xlsx.exe convert_to_yml_in_place {0}'.format(os.path.join(sFolder,item))
			os.system(cmd1)

"""
Parses through all folders to determine if there is a valid before and after folder.
The before folder should contain the initial version of the spreadsheet and the 
after folder will contain a modified version of the same sheet. This script will then
convert both into YML in two new folders which are created by the script.
"""
def parse_folders():
	print('Parsing folders...')
	beforeXLSX = 'before_excel_'
	afterXLSX = 'after_excel_'
	beforeYML = 'before_yml_'
	afterYML = 'after_yml_'

	results = {}

	for item in os.listdir():
		if os.path.isdir(item):
			if item.startswith((beforeXLSX, afterXLSX)):
				name = item.split('_')[2]
				if name not in results:
					results[name] = []
				if len(os.listdir(item)) == 1: #Ensure only one file in directory
					results[name].append(item)
			elif item.startswith((beforeYML, afterYML)):
				delete_folder(item)

	for sheetName in results:
		if len(results[sheetName]) != 2:
			print('Did not find two valid directories for: {0}'.format(sheetName))
			results = {}
			return results
		bFolderIn = beforeXLSX + sheetName
		bFolderOut = beforeYML + sheetName 
		aFolderIn = afterXLSX + sheetName
		aFolderOut =  afterYML + sheetName
		os.mkdir(bFolderOut)
		os.mkdir(aFolderOut)
		copy_between_folders(bFolderIn, bFolderOut)
		copy_between_folders(aFolderIn, aFolderOut)

		convert_to_yml(bFolderOut)
		convert_to_yml(aFolderOut)
		
		print('Parsed {0}'.format(sheetName))
	
	print('folder parse complete.')
	return results




"""
Run the size comparison operation across the available folders with the necessary 
prefixes. Each set of file is committed to the repo and then the size recorded so
a comparison can be performed.
"""
def run_size_compare():
	folderTree = parse_folders()
	if len(folderTree) == 0:
		return

	beforeXLSX = 'before_excel_'
	afterXLSX = 'after_excel_'
	beforeYML = 'before_yml_'
	afterYML = 'after_yml_'

	with open('results_.csv', 'w', newline='') as f:
		writer = csv.writer(f)
		writer.writerow(['Sheet Name', 'File Type', 'Event Name', 'Repo Size (Bytes)'])
		for sheetName in folderTree:
			for fType in [beforeXLSX, afterXLSX, beforeYML, afterYML]:
				print(fType + sheetName)
				writer.writerow([sheetName, fType, 'clone_initial', record_size_repo()])
				copy_into_repo(fType + sheetName)
				cd_into_repo()
				git_commit(fType + sheetName)
				writer.writerow([sheetName, fType, 'commit', record_size_repo()])
				git_repack()
				writer.writerow([sheetName, fType, 'repack', record_size_repo()])
				git_push()
				writer.writerow([sheetName, fType, 'push', record_size_repo()])
				cd_out_repo()
				save_file_sizes('file_sizes_' + fType + sheetName + '.csv')
			
		
"""
A testing function which logs the size of individual files within the repo for 
additional testing.
"""
def save_file_sizes(output_path):

	dirpath = name_repo_folder()
	output_list = []

	total_size = 0

	for dirpath, dirnames, filenames in os.walk(dirpath):
		for f in filenames:
			fp = os.path.join(dirpath, f)
			# skip if it is symbolic link
			if not os.path.islink(fp):
				fSize = os.path.getsize(fp)
				total_size += fSize
				output_list.append([fp, hashlib.md5(fp.encode()).hexdigest(), fSize])

	with open(output_path, 'w', newline='') as f:
		writer = csv.writer(f)
		writer.writerow(['File Path', 'File Hash', 'Size (Bytes)'])

		for row in output_list:
			writer.writerow(row)
	


"""
The entry point for the application. We wait for 5 seconds to allow the
user to switch to the secondary terminal where commands will be input.
We then proceed with the execution of the testing operations. 
"""
time.sleep(5)
run_size_compare()

exit()



