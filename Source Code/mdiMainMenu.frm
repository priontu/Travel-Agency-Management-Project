VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MenuForm 
   BackColor       =   &H8000000C&
   Caption         =   "Main Menu"
   ClientHeight    =   9495
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15390
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMainMenu.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rpt8 
      Left            =   840
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt9 
      Left            =   840
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt10 
      Left            =   840
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt7 
      Left            =   840
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt6 
      Left            =   840
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt5 
      Left            =   840
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt4 
      Left            =   840
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt3 
      Left            =   840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt2 
      Left            =   840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   840
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnfile 
      Caption         =   "Tools"
      Begin VB.Menu mnBackup 
         Caption         =   "Backup"
      End
   End
   Begin VB.Menu mnPackage 
      Caption         =   "Package information"
      Begin VB.Menu mnpackageentry 
         Caption         =   "Package Entry"
      End
      Begin VB.Menu mnpackageexpensesentry 
         Caption         =   "Package Expenses Entry"
      End
      Begin VB.Menu mnpackagebooking 
         Caption         =   "Package Booking "
      End
      Begin VB.Menu mnpackagebookingupdate 
         Caption         =   "Package Booking Update"
      End
   End
   Begin VB.Menu mnVehicle 
      Caption         =   "Vehicle information"
      Begin VB.Menu mnvehicleentry 
         Caption         =   "Vehicle Entry"
      End
      Begin VB.Menu mnvehicleupdate 
         Caption         =   "Vehicle Update"
      End
      Begin VB.Menu mnVehicleBooking 
         Caption         =   "Vehicle Booking"
      End
      Begin VB.Menu mnvehiclebookingupdate 
         Caption         =   "Vehicle Booking Update"
      End
      Begin VB.Menu mnvehiclebilling 
         Caption         =   "Vehicle Billing"
      End
      Begin VB.Menu mnvehicleservicing 
         Caption         =   "Vehicle Servicing"
      End
      Begin VB.Menu mnFuelRecord 
         Caption         =   "Fuel Record"
      End
   End
   Begin VB.Menu mnUser 
      Caption         =   "User information"
      Begin VB.Menu mnusercreation 
         Caption         =   "User Creation"
      End
      Begin VB.Menu mnlogin 
         Caption         =   "User Login"
      End
   End
   Begin VB.Menu mnEmployee 
      Caption         =   "Employee information"
      Begin VB.Menu mnemployeeentry 
         Caption         =   "Employee Entry"
      End
   End
   Begin VB.Menu mnReports 
      Caption         =   "Reports"
      Begin VB.Menu rptEmployeeDetails 
         Caption         =   "Employee Details"
      End
      Begin VB.Menu rptFuelRecord 
         Caption         =   "Fuel Record"
      End
      Begin VB.Menu rptPackageList 
         Caption         =   "List of Packages"
      End
      Begin VB.Menu rptVehicleList 
         Caption         =   "List of Vehicles"
      End
      Begin VB.Menu rptPackageBookingCustomerList 
         Caption         =   "Customer list of Package Booking"
      End
      Begin VB.Menu rptVehicleBookingCustomerList 
         Caption         =   "Customer list of Vehicle Booking"
      End
      Begin VB.Menu rptPackageExpensesList 
         Caption         =   "Package Expenses List"
      End
      Begin VB.Menu rptTotalPackageExpenses 
         Caption         =   "List of Total Package Expenses"
      End
      Begin VB.Menu rptVehicleRentalTakings 
         Caption         =   "Takings from Vehicle Rentals"
      End
      Begin VB.Menu rptVehicleServicing 
         Caption         =   "Vehicle Servicing Details"
      End
   End
   Begin VB.Menu mnUserDoc 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFile_Click()
    FuelRecord.Show
End Sub

Private Sub mnBackup_Click()
    frmBackup.Show
End Sub

Private Sub mnemployeeentry_Click()
    EmployeeEntry.Show
End Sub

Private Sub mnFuelRecord_Click()
    FuelRecord.Show
End Sub

Private Sub mnlogin_Click()
    UserLogin.Show
End Sub

Private Sub mnpackagebooking_Click()
    PackageBooking.Show
End Sub

Private Sub mnpackagebookingupdate_Click()
    PackageBookingUpdate.Show
End Sub

Private Sub mnpackageentry_Click()
    PackageEntry.Show
End Sub

Private Sub mnpackageexpensesentry_Click()
    PackageExpensesEntry.Show
End Sub

Private Sub mnusercreation_Click()
    usercreation.Show
End Sub

Private Sub mnvehiclebilling_Click()
    VehicleBilling.Show
End Sub

Private Sub mnvehiclebooking_Click()
    VehicleBooking.Show
End Sub

Private Sub mnvehiclebookingupdate_Click()
    VehicleBookingUpdate.Show
End Sub

Private Sub mnvehicleentry_Click()
    VehicleEntry.Show
End Sub

Private Sub mnvehicleservicing_Click()
    VehicleServicing.Show
End Sub

Private Sub mnvehicleupdate_Click()
    VehicleUpdate.Show
End Sub


Private Sub rptEmployeeDetails_Click()
rpt1.ReportFileName = App.Path & "\Reports\EmployeeDetails.rpt"
'CrystalReport1.SelectionFormula = "{guestentry.ticketnumber}=" & txtticketno.Text & ""
rpt1.Action = 2
End Sub

Private Sub rptFuelRecord_Click()
rpt2.ReportFileName = App.Path & "\Reports\FuelRecord.rpt"
rpt2.Action = 2
End Sub

Private Sub rptPackageBookingCustomerList_Click()
rpt5.ReportFileName = App.Path & "\Reports\CustomerListOfPackageBooking.rpt"
rpt5.Action = 2
End Sub

Private Sub rptPackageExpensesList_Click()
rpt7.ReportFileName = App.Path & "\Reports\PackageExpensesList.rpt"
rpt7.Action = 2
End Sub

Private Sub rptPackageList_Click()
rpt3.ReportFileName = App.Path & "\Reports\ListOfPackages.rpt"
rpt3.Action = 2
End Sub

Private Sub rptTotalPackageExpenses_Click()
rpt3.ReportFileName = App.Path & "\Reports\ListOfTotalPackageExpenses.rpt"
rpt3.Action = 2
End Sub

Private Sub rptVehicleBookingCustomerList_Click()
rpt6.ReportFileName = App.Path & "\Reports\CustomerListOfVehicleBooking.rpt"
rpt6.Action = 2
End Sub

Private Sub rptVehicleList_Click()
rpt4.ReportFileName = App.Path & "\Reports\ListOfVehicles.rpt"
rpt4.Action = 2
End Sub

Private Sub rptVehicleRentalTakings_Click()
rpt9.ReportFileName = App.Path & "\Reports\ListOfTakingsFromVehicleRentals.rpt"
rpt9.Action = 2
End Sub

Private Sub rptVehicleServicing_Click()
rpt10.ReportFileName = App.Path & "\Reports\VehicleServicingDetailsList.rpt"
rpt10.Action = 2
End Sub
