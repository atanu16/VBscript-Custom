pip install Office365-REST-Python-Client


For Power Automate Desktop:
	1.	Add the “Run Python script” or “Run a command line script” action.
2.	Set the command:

python "C:\Path\To\update_sharepoint.py" "Credit" "BluePrism" "AA" "Machone" "Completed" "BNA" "2:00 PM" "4:00 PM"

3.	Make sure:
	•	Python and office365-rest-python-client are installed.
	•	The .py file is in an accessible path.




For Automation Anywhere (A360):
	1.	Use the “Run Script” or “Run DOS Command” action.
	2.	Command:

python "C:\Path\To\update_sharepoint.py" "Creation" "Blum" "8TI" "7IK" "Completed" "BNA" "2:00 PM" "4:00 PM"

3.	Ensure:
	•	Python is installed and available in environment variables.
	•	Bot runner has access to the script and SharePoint.
