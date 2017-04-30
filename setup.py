import os, shutil
from setuptools import setup, find_packages

def readme():
	with open('README.rst') as f:
		return f.read()
		
def create_user_files():
	# Create a new directory, or use the existing one
	src_dir = os.getcwd()
	dest_dir = os.path.expanduser('~/Documents/SAIDI and SAIFI/')
	new_dirs = ["", "Stats", "Templates", "Scripts"]
	for dir in new_dirs:
		if not os.path.exists(os.path.join(dest_dir, dir)):
			os.makedirs(os.path.join(dest_dir, dir))
	
	# Copy database/src files from the Data folder for the user
	file_names = ["SAIDI SAIFI Calculator.xlsm", "ICP_Search_Prog_a2k3.mde", "EXTRA RECORDS.csv"]
	for filename in file_names:
		try:
			shutil.copy2(os.path.join(src_dir, "Data", filename), os.path.join(dest_dir, filename))
			print "File %s successfully copied" % filename
		except IOError:
			if os.path.isfile(os.path.join(dest_dir, filename)):
				print "The file %s is in use and cannot be replaced/updated" % filename
			else:
				print "The file %s is missing and cannot be coppied" % filename
	
	# Copy all the template files
	for file in os.listdir(os.path.join(src_dir, "Data")):
		if file.endswith(".crtx") or file == "Templates.xlsx" or file == "Weekly Report Template.docx":
			shutil.copy2(os.path.join(src_dir, "Data", file), os.path.join(dest_dir, "Templates"))
		elif file.endswith(".py"):
			shutil.copy2(os.path.join(src_dir, "Data", file), os.path.join(dest_dir, "Scripts"))

		

if __name__ == "__main__":
	setup(name='SAIDI/SAIFI (ORS)',
		  version='0.1',
		  description='A package to assist with SAIDI/SAIFI calculations and data presentation',
		  long_description=readme(),
		  url='https://github.com/Top-Dog/SAIDI-SAIFI-Calculator',
		  author="Sean D. O'Connor",
		  author_email='sdo51@uclive.ac.nz',
		  license='MIT',
		  packages=['SAIDISAIFI', 'SAIDISAIFI.Parser', 'SAIDISAIFI.Output'],
		  # adding packages
		  #packages=find_packages('src'),
		  #package_dir = {'':'src'},
		  install_requires=['numpy', 'pypiwin32', 'pyodbc', 'six', 'wheel', 'virtualenv', 'tabulate'], #pywin32
		  
		  # Add static data files
		  include_package_data = True,
		  package_data = {
			'': ['*.txt'],
			'': ['static/*.txt'],
			'static': ['*.txt'],
			},
		  
		  dependency_links=[], 
		  #zip_safe=False,
		  #packages=find_packages()
		  )
	
	# Create some files for the user
	create_user_files()
