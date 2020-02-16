Elite Few War Planner
======================
This is a python script for determing war plans in Hero Wars. Currently this should handle your typical local server wars that don't require a Grand Arena team. 

**Note:** This script is currently a work in progress and you should review its output to see if you agree with what it has come up with.


## Installation Instructions
1. Install python 3.8 from [here ](https://www.python.org/downloads/release/python-381/ "Python Download Page")
2. Once installed, open command prompt (assuming windows, otherwise use terminal) and type the following commands:
> pip install xlrd
3. Download this project from github. Then take a look at the usage section.

# Script Usage:
### Setup: Create your version of the excel input file.
**Note:** The Name doesn't matter, you may want to create a copy of the sample provided and then change the values.

1. Use the first sheet to enter the power of your team
2. Use the 2nd sheet to enter the power and location of your opponent



### This Script can be run in 2 ways:

1. Drag and drop the excel file onto the startPlanning.py script.
2. run from the command line.
	* Open the containing folder and Shift + right click. Choose Open Powershell or Command Prompt window here
	* then type the following command (example from attached files, change the name if you create a new excel file etc.)
	>python startPlanning.py FMEvsLA2-13.xlsx
	* Using command line you can add a 3rd parameter set to 'True' if you want to have assignments seperated by Hero and Titan assignments
	>python startPlanning.py FMEvsLA2-13.xlsx True


# Screenshots:

![alt text](../resources/ScriptOutput.png)
