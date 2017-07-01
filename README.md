# ProfileComparisonTool

Description – This tool can be used for comparing two profiles with each other. It generates a excel sheet with various tabs showing all the differences.

How is the tool built– This tool is built in Java and uses XML DOM Parser for parsing the Profile XMLs and generate Excel sheet with comparison results. Excel sheet is generated using Apache POI API.

How to use this Tool – Dowload the zip file from the repo and then Follow below steps for using this tool –

1. Go to Tool folder and Use package.xml given with code to retrieve metadata from Workbench or you can use Force.com Migration tool. Just make sure all standard objects that we are using are listed in this xml.(Ideally you dont have to change this since most of the objects are already listed) - If you dont know how to retrieve profile using workbench, here are the instructions - https://sushilsfdc.blogspot.com/2017/04/retrieve-salesforce-profile-xml-using.html

2. Once you have retrieved the code, Open the profiles folder in retrieved package and Copy below files in that folder - 
CompareProfiles.bat
ProfileComparisonTool.jar

3. Open CompareProfiles.bat file in text editor and update profile names you want to compare.

4. Close it and then double click on CompareProfiles.bat. This should open command prompt and generate the profile comparison XLS.

What are permissions compared by this tool

Assigned Apps, Object Settings which includes - Object Access, Page Layout Assignments, Record Type Access, and Field level security, Tab Visibility, App Permissions, Apex Class Access, Visualforce Page Access, External Data Source Access, System/User Permissions, Custom Permissions


What are the permissions not compared by this tool -Assigned Connected App, Named Credential Access, Data Category Visibility, Desktop Client Access, Login Hours, Login IP Ranges, Service Providers, Session Timeout, Password Policies

Limitation - In case you have App Exchange Packages, Above xml will not retrieve profile permisssion for package components. You will have to retrieve app exchange package permission separately. 
