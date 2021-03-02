VERSION 5.00
Begin VB.Form PackageBooking 
   Caption         =   "Package booking"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form6"
   ScaleHeight     =   9090
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Information on Package"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   3720
      TabIndex        =   25
      Top             =   840
      Width           =   3375
      Begin VB.Label lblVacancies 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   45
         Top             =   7560
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Number of Vacancies"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1200
         TabIndex        =   44
         Top             =   7320
         Width           =   1665
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Days"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   42
         Top             =   7560
         Width           =   615
      End
      Begin VB.Label lblAccommodation 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Label lblPackageDetails 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   2970
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Package Details"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Accommodation type"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   4920
         Width           =   1620
      End
      Begin VB.Label lblDuration 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   7560
         Width           =   375
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   7320
         Width           =   930
      End
      Begin VB.Label lblArrivalTime 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   6360
         Width           =   3015
      End
      Begin VB.Label lblStartingTime 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   6960
         Width           =   3045
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Arrival time"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   6120
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Starting time"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   6720
         Width           =   1275
      End
      Begin VB.Label lblJourneyDate 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   5760
         Width           =   3015
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Journey Date"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   5520
         Width           =   1275
      End
      Begin VB.Label lblItinerary 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Itinerary"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   7320
      TabIndex        =   7
      Top             =   840
      Width           =   3255
      Begin VB.TextBox txtContactNumber 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   3000
      End
      Begin VB.TextBox txtBookingID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3000
      End
      Begin VB.TextBox txtCustomerName 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3000
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Booking ID"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Contact number"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Package Specifications"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3375
      Begin VB.ListBox lstPackageName 
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         ItemData        =   "Package booking.frx":0000
         Left            =   120
         List            =   "Package booking.frx":0002
         TabIndex        =   46
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtPackageID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   5280
         Width           =   3015
      End
      Begin VB.ComboBox cboPackageType 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Package booking.frx":0004
         Left            =   120
         List            =   "Package booking.frx":0006
         TabIndex        =   23
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtNetTotal 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   7440
         Width           =   3000
      End
      Begin VB.TextBox txtNPeople 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   6720
         Width           =   3000
      End
      Begin VB.TextBox txtCostPerPerson 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   6000
         Width           =   3000
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "PackageID"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   38
         Top             =   5040
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Package Type"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Net total"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   7200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Number of people"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cost per person"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   5760
         Width           =   1530
      End
      Begin VB.Label Label5 
         Caption         =   "Package name"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   7320
      TabIndex        =   0
      Top             =   5400
      Width           =   3255
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   1260
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Package Booking"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3240
      TabIndex        =   43
      Top             =   120
      Width           =   3780
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9480
      TabIndex        =   5
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8880
      TabIndex        =   4
      Top             =   360
      Width           =   600
   End
End
Attribute VB_Name = "PackageBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Package name list click event.
Private Sub lstPackageName_Click()
    Set rs = New Recordset
'Clearing data from text boxes in case they need to be changed.
    txtNPeople.Text = ""
    txtNetTotal.Text = ""
'Package information loaded automatically from Package Entry database at the selection of a Package name.
    rs.Open "select * from Package_entry where Package_name = '" & lstPackageName.Text & "'", con, adOpenDynamic, adLockOptimistic

'The package information is loaded automatically at the selection of a package.
    txtCostPerPerson.Text = rs.Fields("Cost_per_person")
    lblItinerary.Caption = rs.Fields("Itinerary")
    lblPackageDetails.Caption = rs.Fields("Package_details")
    lblJourneyDate.Caption = rs.Fields("Journey_date")
    lblArrivalTime.Caption = rs.Fields("Arrival_time")
    lblStartingTime.Caption = rs.Fields("Starting_time")
    lblAccommodation.Caption = rs.Fields("Accommodation_type")
    lblDuration.Caption = rs.Fields("Duration/days")
    txtPackageID.Text = rs.Fields("PackageID")
    lblVacancies.Caption = rs.Fields("NumberOfReservations") - rs.Fields("NumberReserved")
End Sub
'Package type combobox click event.
Private Sub cboPackageType_Click()
'The Package name list is always cleared when the user changes the package type in case the old imformation is not overriden.
    lstPackageName.Clear
    txtNPeople.Text = ""
    txtNetTotal.Text = ""
    txtPackageID = ""
    lblItinerary.Caption = ""
    lblPackageDetails = ""
    lblAccommodation.Caption = ""
    lblJourneyDate.Caption = ""
    lblStartingTime.Caption = ""
    lblArrivalTime.Caption = ""
    lblDuration.Caption = ""

   
'Package names are laoded onto the Package names list on the selection of a particular package type. Only packages of this particular type are selected and loaded onto the list.
'Opening the Package Entry database.
   Set rs = New Recordset
'Opening the Package Entry database.
    rs.Open "select * from Package_entry where Package_type ='" & cboPackageType.Text & "'", con, adOpenDynamic, adLockOptimistic
'Loading the Package names of the particular type selected into the Package name list.
    Do While Not rs.EOF
        lstPackageName.AddItem (rs.Fields("Package_name"))
        rs.MoveNext
    Loop
    
End Sub
'Confirm button click event.
Private Sub cmdConfirm_Click()
'Validation check to make sure all the inportant fields are filled.
    If cboPackageType.Text = "" Or txtPackageID.Text = "" Or txtCostPerPerson.Text = "" Or txtNPeople = "" Or txtBookingID.Text = "" Or txtCustomerName.Text = "" Or txtAddress.Text = "" Or txtContactNumber.Text = "" Then
'Message to notify the user that all of the information need to be provided.
        MsgBox "All the Package information and customer information need to provided in order to proceed with the booking."
        Exit Sub
    End If
    
'Length check to make sure contact number is of 11 digits.
    If Len(txtContactNumber.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If
    
    Set rs = New Recordset
'Updating reservation information in the Package Entry database.
    rs.Open "select * from Package_entry where Package_name ='" & lstPackageName.Text & "'", con, adOpenDynamic, adLockOptimistic
        p = rs.Fields("NumberReserved")
        p = p + txtNPeople.Text
    If rs.Fields("NumberOfReservations") > p Then
        rs.Fields("NumberReserved") = p
        rs.Update
    Else
'Validation check to check if there are any vacancies.
'Message to notify the user that there are no vacancies that are available.
            MsgBox "Sorry, there are no vacancies for this package. All the slots have been booked. Please try another package."
            Exit Sub
        End If
    rs.Close
'Opening the Package booking databse.
    rs.Open "select * from Package_booking", con, adOpenDynamic, adLockOptimistic
'Validation check to make sure that the same ID cannot be used twice.
    Dim IsNewID As Boolean
    IsNewID = True
    If Not rs.EOF Then
            Do While Not rs.EOF
                If txtBookingID.Text = rs.Fields("PackageBookingID") Then
                    IsNewID = False
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
    
    If IsNewID = False Then
'Message to notify that the user if booking ID is already in use.
        MsgBox "This Booking ID is already in use. Please try again."
        txtBookingID.Text = ""
        Exit Sub
    End If

'Storage of Booking information into the database.
    rs.AddNew
        rs.Fields("PackageBookingID") = txtBookingID.Text
        rs.Fields("Customer_name") = txtCustomerName.Text
        rs.Fields("Customer_address") = txtAddress.Text
        rs.Fields("Package_type") = cboPackageType.Text
        rs.Fields("Journey_Date") = lblJourneyDate.Caption
        rs.Fields("Contact") = txtContactNumber.Text
        rs.Fields("Cost_per_person") = txtCostPerPerson.Text
        rs.Fields("Package_name") = lstPackageName.Text
        rs.Fields("Number_of_people") = txtNPeople.Text
        rs.Fields("Net_total") = txtNetTotal.Text
        rs.Fields("Date_of_booking") = lblDate.Caption
        rs.Fields("PackageID") = txtPackageID.Text
        rs.Fields("Accommodation_type") = lblAccommodation.Caption
        rs.Fields("Journey_date") = lblJourneyDate.Caption
        rs.Fields("Starting_time") = lblStartingTime.Caption
        rs.Fields("Arrival_time") = lblArrivalTime.Caption
        rs.Fields("Duration/days") = lblDuration.Caption
        rs.Fields("Itinerary") = lblItinerary.Caption
        rs.Fields("Package_details") = lblPackageDetails.Caption
    rs.Update
 'Message shown to notify the user that storage was sucessful.
    MsgBox "Booking has been made successfully."
    rs.Close
    
'The fields are emptied after the storage of a record.
    txtBookingID.Text = ""
    txtCustomerName.Text = ""
    txtAddress.Text = ""
    cboPackageType.Text = ""
    txtContactNumber.Text = ""
    txtCostPerPerson.Text = ""
    txtNPeople.Text = ""
    txtNetTotal.Text = ""
    txtPackageID = ""
    lblItinerary.Caption = ""
    lblPackageDetails = ""
    lblAccommodation.Caption = ""
    lblJourneyDate.Caption = ""
    lblStartingTime.Caption = ""
    lblArrivalTime.Caption = ""
    lblDuration.Caption = ""
    lstPackageName.Clear
    lstPackageName.Refresh
    
    Call Form_Load
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'The fields are emptied and the form is refreshed to allow the user to start anew.
    txtBookingID.Text = ""
    txtCustomerName.Text = ""
    txtAddress.Text = ""
    cboPackageType.Text = ""
    txtContactNumber.Text = ""
    txtCostPerPerson.Text = ""
    txtNPeople.Text = ""
    txtNetTotal.Text = ""
    txtPackageID = ""
    lblItinerary.Caption = ""
    lblPackageDetails = ""
    lblAccommodation.Caption = ""
    lblJourneyDate.Caption = ""
    lblStartingTime.Caption = ""
    lblArrivalTime.Caption = ""
    lblDuration.Caption = ""
    lstPackageName.Clear
    lstPackageName.Refresh
    
    Call Form_Load
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Showing current date on the form.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
    Set rs = New Recordset
    rs.Open "select * from Package_booking", con, adOpenDynamic, adLockOptimistic
        rs.MoveLast
        txtBookingID.Text = CInt(rs!PackageBookingID) + 1
    rs.Close
    
    Set rs = New Recordset
    rs.Open "select * from Package_entry", con, adOpenDynamic, adLockOptimistic
'Loading Package Types from Package Entry Database.
'A check is made to make sure that each type is not loaded twice.
    If rs.EOF = False Then
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To cboPackageType.ListCount - 1
                If cboPackageType.List(i) = rs.Fields("Package_type") Then
                    IsNewType = False
                End If
            Next
            If IsNewType = True Then
                cboPackageType.AddItem (rs.Fields("Package_type"))
            End If
            rs.MoveNext
        Loop
    End If
    
End Sub
'Contact Number keypress event.
Private Sub txtContactNumber_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
'Cost per person keypress event.
Private Sub txtCostPerPerson_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
   If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
'Net total keypress event.
Private Sub txtNetTotal_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
   If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
'Number of people keypress event.
Private Sub txtNPeople_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
   If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub

'Number of people keyup event.
Private Sub txtNPeople_KeyUp(KeyCode As Integer, Shift As Integer)
'The bar moves to booking ID textbox when enter is pressed after the number of people textbox is filled.
    If KeyCode = vbKeyReturn Then
        txtBookingID.SetFocus
    End If
    
End Sub
'Number of people lostfocus event.
Private Sub txtNPeople_LostFocus()
'Calculation of net total by multiplying Number of people with Cost per person.
    txtNetTotal.Text = txtNPeople.Text * txtCostPerPerson.Text
End Sub
