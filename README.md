- Digital solution to a Travel Agency Management System containing a simple, friendly user interface
- User interface includes account sign-up, log-in capability, and a home page with menu linking all forms
- Functionalities involve customer, vehicle, package and employee registrations (separate forms), package and vehicle booking and billing generation, automatic report and receipt generation on completion of booking/billing, fuel use record keeping, updating vehicle or booking information and backing-up data
- VB6 used for development, Ms Access used for database; complete with documentation
- Complete Documentation provided here:
https://raw.githubusercontent.com/priontu/Travel-Agency-Management-Project/main/Travel%20Agency%20Management%20System%20documentation(%20Main).pdf

- User Documentation (smaller) provided here:
https://github.com/priontu/Travel-Agency-Management-Project/blob/main/User%20Documentation.pdf




 
Installation instructions
The steps that need to be taken for the installation of the software are:
1.	i. Open “Travel Agency” folder from Project (D:) drive.
ii. Open folder named “Package”.
iii. Double click on Setup.exe file.

![image](https://user-images.githubusercontent.com/61733487/153757453-b427615f-efa1-42b5-94cc-1b30cd662f0b.png)

 
2.	Click “Ok”
3.	Click the button shown by the arrow.








	
 
4.	Click “Continue”.











5.	A progress bar is shown


 
6.	The installation is finished when the message shown below appears.






7.	Starting the system
The Log-in form appears when the following icon is clicked from the desktop window. The access to Main Menu form can only be gained if the correct User type, User ID and Password are provided.
 
The system basically has two parts. A part that is designed to interact with us and the other part is used to store the data we provide to the system and allow the computer system to make calculations and monitor the things happening. The part that was made to interact with the user is designed using Visual Basic 6.0 and the other part is made using Ms Access. Programming techniques have been used to link the two parts to function efficiently i.e. a connection has been set up between the database and the data environment of Visual Basic 6.0.

Logging into the system
1.	Enter the correct User type, User ID and Password.
2.	Click Log-in
3.	When the button is pressed, the Main Menu form will appear.




Adding a new User to the system
1.	Select “User Creation” form from the “User information” section of the Main Menu.
2.	Enter User type, User ID, Password and Confirm Password.
3.	The Password and Confirm Password must match.
4.	Click “Create Account” button.
5.	A message will appear that the account has been created successfully. 
6.	Click “Ok”






Changing the Password of an existing User
1.	Select “User Creation” form from the “User information” section of the Main Menu.
2.	Provide User type and User ID in the User creation form.
3.	Click “Search”.
4.	The previous Password and Confirm password will appear.
5.	Change the Password and Confirm Password both. It should be made sure that both of them match.
6.	Click “Update”.
7.	Confirmation message shown on successful update of data.
8.	Click “Ok”.







Deleting an existing User Account
1.	Select “User Creation” from the “User information” section of the Main Menu.
2.	Provide User type and User ID.
3.	Click “Search”.
4.	Click “Delete”.
5.	A confirmation message will appear for the deletion of account.
6.	Click “Ok”.

 
Adding a new Employee record 
1.	Select “Employee Entry” form from “Employee information” section of the Main Menu.
2.	Enter all the Employee information.
3.	Click  “Make Entry”
4.	Confirmation message appears for the successful addition of data.
5.	Click “Ok”.



 
Updating an existing Employee record 
1.	Select “Employee Entry” form from the “Employee information” section of the Main Menu.
2.	Select an Employee ID from the “Employee ID list” of the Update section of the form.
3.	Make changes to the information.
4.	Click “Update”.
5.	Confirmation message appears on successful update. Click “Ok”.














Deleting an existing Employee record 
1.	Select an Employee ID from the Employee ID list of the Update section of the form.
2.	Click “Delete”.
3.	Confirmation message appears on successful deletion of the record.
4.	Click “Ok”.



Adding a new Package
1.	Select “Package Entry” form from the “Package information” section of the Main Menu.
2.	Provide Package information, Number of reservations, itinerary and Package Details.
3.	Click “Save”.
4.	Confirmation message appears on successful addition of the record.
5.	Click “Ok”.






















	
 
Updating a Package
1.	Select “Package Entry” form from the “Package information” section of the Main Menu.
2.	Select the “Package type” from the “Update” section of the form.
3.	List of Package names is loaded.
4.	Select the Package name.
5.	The fields on the form are filled.
6.	Make required changes.
7.	Click “Update”.
8.	Confirmation message appears on update of information.
9.	Click “Ok”

















Deleting a Package
1.	Select “Package Entry” form from the “Package information” section of the Main Menu.
2.	Select a “Package type” from the “Update” section.
3.	List of Package names is loaded. Select a Package name.
4.	Package information for the Package is loaded.
5.	Click “Delete”.
6.	Confirmation message appears on successful deletion of the record.
7.	Click “Ok”.

 
Booking a Package
1.	Select “Package Booking” form from the “Package information” section Main Menu form.
2.	Select a Package type. List of Package names is loaded. Select a Package name.
3.	Check if there are any vacancies for the Package. Enter Number of People. Net total is automatically calculated.
4.	Enter Customer information.
5.	Click “Confirm” button.
6.	Confirmation message appears on successful booking made.
7.	Click “Ok”.

















Updating a Package
1.	Select “Package Booking Update” form from the “Package information” section of the Main Menu form.
2.	Provide Booking ID and click “Search” button.
3.	The fields on the form are filled.
4.	Make required changes.
5.	Click “Update” button.
6.	Confirmation message appears on successful update of record.
7.	Click “Ok”.
 
Deleting a Booking record
1.	Select “Package Booking Update” form from the “Package information” section of the Main Menu.
2.	Provide Booking ID of the record to be deleted.
3.	Click “Search” button.
4.	All the fields on the form are filled.
5.	Click “Delete” button.
6.	Confirmation message appears on successful deletion of record.
7.	Click “Ok”.


















	
Recording Package Expenses information
1.	Select “Package Expenses Entry” form from the “Package information” section of the Main Menu.
2.	Provide Package ID, Package name, Expenses and Amount.
3.	Click “Done”. The record is added to the table in “Overall Expenses data” section of the form.
4.	Repeat the process for the number of expenses present.
5.	After all the expenses data have been provided, click “Calculate” button.
6.	The total amount is calculated and total generated in the textbox.
7.	Click “Save”.
8.	Confirmation message appears for successful saving of data.
9.	Click “Ok”
 
Adding a record for the information of a new Vehicle
1.	Select “Vehicle Entry” form from the “Vehicle information” section of the Main Menu.
2.	A Vehicle ID is generated.
3.	Fill the form with all the Vehicle information.
4.	Press “Submit” button
5.	Confirmation message appears on addition of record to the database.
6.	Click “Ok”.







Updating the information of a record of a Vehicle 
1.	Select “Vehicle Update” form from the “Vehicle information” section of the Main Menu.
2.	Select Vehicle ID from the Vehicle ID list. The form is filled.
3.	Make required changes. 
4.	Click “Update” button.
5.	Confirmation message appears on successful update.
6.	Click “Ok”.

 
Deleting a record of information on a Vehicle
1.	Select “Vehicle Update” form from the “Vehicle information” section of the Main Menu form.
2.	Select the Vehicle ID of the record to be deleted from the Vehicle ID list.
3.	Click Delete.
4.	Confirmation message appears on successful deletion of the record.
5.	Click “Ok”.







Booking a Vehicle
1.	Select “Vehicle Booking” form from the “Vehicle information” section of the Main Menu form.
2.	“Vehicle Booking” form appears and a Booking ID is generated.
3.	Fill up the Customer information section of the form and provide Journey date and Drop-off date.
4.	Select Vehicle ID from the Vehicle information section. A list of Makes appears.
5.	On selection of a Make, a list of Models appears. A list of Vehicle IDs appears on selection of a Model. Finally, Registration number, Rate per kilometer and Number of seats appears on selection of a Vehicle ID
6.	 Enter Starting kilometer, Advance paid and Driver name.
7.	Click “Confirm” button.
8.	Confirmation message appears on successful booking.
9.	Click “Ok”
	
 





















	
 
Updating information of an existing booking
1.	Select “Vehicle Booking Update” form from the “Vehicle information” section of the Main Menu form.
2.	Provide Booking ID. Click “Search” button.
3.	The rest of the fields on the form are filled.
4.	Make changes to the information.
5.	Click “Update” button.
6.	Confirmation message appears on successful update of the record.
7.	Click “Ok”.















 
	Deleting an existing booking
1.	Select “Vehicle Booking Update” form from the Vehicle information section of the Main Menu form.
2.	Provide Booking ID. Click “Search” button.
3.	Click “Delete” button.
4.	Confirmation message appears on successful deletion.
5.	Click “Ok”.

 
Billing a Vehicle Rental
1.	Select the “Vehicle Billing” form from the “Vehicle information” section of the Main Menu form.
2.	Provide Booking ID. Click “Search” button.
3.	Provide “Ending kilometer”. The “Cost of travel” is calculated automatically.
4.	Provide Booking charge, Driver charge, Additional Expenditures and Discount.
5.	Click “Generate Bill” button.
6.	Click “Save” button.
7.	Click “Print” button.
8.	Click “Refresh” button.
		
 
Addition of Fuel Record data
1.	Select “Fuel Record” form from the “Vehicle information” section of the Main Menu form.
2.	Enter Vehicle ID, Fuel type, Quantity and Amount.
3.	Click “Save” button.
4.	Confirmation message shown on successful addition of data to the database.
5.	Click “Ok”.












Recording data on servicing of the Vehicles
1.	Select “Vehicle Servicing” form from the “Vehicle information” section of the Main Menu form.
2.	Select a Vehicle type. List of Vehicle IDs is loaded. 
3.	Select a Vehicle ID. Table of “Servicing History Preview” is loaded.
4.	Enter “Servicing Details” and “Amount”.
5.	Click “Save”.
6.	The record is added to the table and the textboxes are emptied.

 






















	
 
Generating Reports
1.	Select the name of the reports to be viewed from the “Reports” section of Main Menu form.
2.	The report is generated.











Printing Reports
1.	Select the report to be printed from the “Reports “section of the Main Menu form.
2.	Click on the button highlighted below.




3.	/	The following window appears. Enter number of pages that need to be printed.
	



4.	Click “Ok” to print the report.



Creating Backup
1.	Select “Backup” form from the “Tools” section of the Main Menu.
2.	Select directory of target file and the directory of backup file.
3.	Click “Backup files” button.
4.	A window appears asking for name of target file. Enter name of target folder and click “Ok”.
5.	Another window appears asking for name of backup file. Enter name of Backup file and click “Ok”.
6.	Confirmation message appears on successful creation of backup file.










The whole system is linked with Ms Access. All kinds of processes (such as adding, searching, editing, deleting etc.) require Ms Access to work on the background. Therefore it is an essential part for the functioning of the whole system.
	Security Measures
Since the program will include important data, it becomes crucial to take measures to keep the data safe. Special care must be taken to make sure anyone who is not authorized does not get access to the system. 
Physical protection: The computer must be shut down and locked away when it is not being used. When data is carried through flash drives they should be kept safe and measures should be taken so that they are not lost. In case portable computers (such as laptops) are being used, they should be kept safe and out of reach of other people as it is easy to snatch them and run away.
Backing up:  The database information should be backed up periodically and frequently (as transactions are made on an hourly basis) to prevent loss of data and to compensate for accidental loss of data due to unforeseen reasons. It is also advisable to move the backed up information to a new storage medium regularly in case a computer system crashes. Furthermore, it is also advisable to keep a different storage media for the storage of files which could be kept away from the agency venue to allow retrieval of data in case of any unforeseen hazard (such as a fire).
Virus protection:  A licensed antivirus should be installed to the system which needs to be updated periodically in order to protect files from getting infected by a virus. An infected file may be corrupted which will result in loss of data. With an antivirus comes a strong firewall that will prevent illegal access to the system and protect the files in the system.
Password and encryption: Depending on the version being used, it is possible to set a windows password (for further information, see Windows “Help”). This is highly recommended as it prevents unknown people from accessing user account of the operating system. My system is already password protected for access to only users of the system. No one can use my system without a proper User type, User ID and Password. 
The data can also be encrypted so that no one can read the data even at gaining illegal access.




	
	On-screen help examples
Throughout the system, on-screen help is available. Whenever the mouse pointer hovers over a control, a message appears notifying the user of the function of the control.



















 
Tool-tip texts are provided throughout the system. Moreover, the customer can refer to the “Help” option in the Main Menu form in case they forget the next step to be taken in the middle of a process.
•	Troubleshooting
This section is a guide to what should be done in case any part of the system is not working as well as it should. It should help encounter small problems. However, any further problem faced requires the council of the programmer. 

What should I do if my report does not print?
	Thoroughly check if the printer is properly installed to the computer properly.
	If the printer is installed, check if the printer is turned on and if there are papers in the printer. Ensure proper connection of the printer to the computer.
	If the above does not work, reinstall the printer and try again.
What should I do if there is no action even when I press a button on the form?
	It is possible that the system has crashed. It is advisable to restart the computer and try again.
What should I do if I am not able to gain access to the system due to wrong password?
	It is advisable to check the password provided thoroughly. Sometimes access may not be granted due to case sensitivity, or leaving “Caps Lock” activated by accident.
I have installed the system but clicking on the icon on the screen does not start the system. What should I do?
	Probably the system was not installed properly. It is advisable to uninstall the system (from the “Programs and features” option of the control panel) and reinstall it again.
Pressing the power button is not turning on my computer. What should I do?
	Check if the computer is connected to the power supply properly. If it is connected properly and still the computer does not start, the user should consult with a technician to examine the power box of the System Unit.

 
	Glossary of terms
•	Antivirus:  Software designed to detect and deal with virus malware.
•	ASCII: American Standard Code of Information Interchange.
•	ASCII value: these are standard recognized ASCII values of different characters. For example, ‘A’ has an ASCII value of 65, which is taken as a standard identification of ‘A’ by computers. That is, when the computer receives an ASCII value of 65, it deems it as ‘A’. 

•	Backup: Copy of data stored for security reasons and used in the event of loss or corruption of live data.
•	Database: Collection of related data. A complete set of data is called a record, while the different types of data present in the tables are called fields. 

•	Encryption: Process whereby a message when transmitted can only be understood by sender and receiver. The coded format cannot be understood by anyone else who reads it. For example, the number “1121” can be written as “3343”, where each number is incremented by 1. The latter is the encrypted version of the former code.

•	Flash drive:  A type of portable storage media (also known as pen drive or memory stick).

•	Malware: Software that has been designed for mischievous or criminal purposes; it might slow down the system or cause deletion of files etc. 

•	Microsoft Access: database software that has been used to develop the system. It has a wide range of features that makes it efficient database software. It is compatible with Visual basic 6.0. It was made by the Microsoft Company.

•	Reports: Hard or soft copy outputs-usually providing summary information.

•	Tables: One of the many collections of data in form of fields and records which might be present in a database. 

•	Tool tip text: A function in Visual Basic that allows a text to be displayed when the mouse is paused over a button. 

•	Validation check: The data input to the system are checked for validity through a number of tests. These include length checks, presence checks and character type checks.

•	Virus: Malicious self-replicating software that can harm the computer system.

•	Visual Basic 6.0: An event driven programming language which has been used for the development of this program.

