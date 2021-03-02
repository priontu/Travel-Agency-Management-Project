VERSION 5.00
Begin VB.Form PackageBookingUpdate 
   Caption         =   "Package Booking Update"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   9270
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
      TabIndex        =   31
      Top             =   960
      Width           =   3375
      Begin VB.Label lblVacancies 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   50
         Top             =   7560
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Number of vacancies"
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
         TabIndex        =   49
         Top             =   7320
         Width           =   1650
      End
      Begin VB.Label Label17 
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
         TabIndex        =   46
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblItinerary 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   45
         Top             =   600
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
         TabIndex        =   44
         Top             =   5520
         Width           =   1035
      End
      Begin VB.Label lblJourneyDate 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   5760
         Width           =   3015
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
         TabIndex        =   42
         Top             =   6720
         Width           =   1035
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
         TabIndex        =   41
         Top             =   6120
         Width           =   900
      End
      Begin VB.Label lblStartingTime 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   6960
         Width           =   3045
      End
      Begin VB.Label lblArrivalTime 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   6360
         Width           =   3015
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
         TabIndex        =   38
         Top             =   7320
         Width           =   690
      End
      Begin VB.Label lblDuration 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   7560
         Width           =   375
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
         TabIndex        =   35
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label lblPackageDetails 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   34
         Top             =   2880
         Width           =   2970
      End
      Begin VB.Label lblAccommodation 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   5160
         Width           =   3015
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
         TabIndex        =   32
         Top             =   7560
         Width           =   375
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
      Height          =   3255
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtBookingID 
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
         TabIndex        =   47
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   3000
      End
      Begin VB.TextBox txtContactNumber 
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
         TabIndex        =   19
         Top             =   2760
         Width           =   3000
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
         TabIndex        =   25
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   615
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
         TabIndex        =   23
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   870
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
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
         Top             =   1200
         Width           =   3000
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
      Height          =   4695
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   3255
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
         TabIndex        =   30
         Top             =   2400
         Width           =   3000
      End
      Begin VB.ComboBox cboPackageName 
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
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   3015
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
         TabIndex        =   26
         Top             =   3600
         Width           =   3000
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
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3015
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
         TabIndex        =   10
         Top             =   3000
         Width           =   3000
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
         TabIndex        =   9
         Top             =   4200
         Width           =   3000
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Package ID"
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
         TabIndex        =   29
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   1410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PackageType"
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
         TabIndex        =   16
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1140
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
         TabIndex        =   14
         Top             =   2760
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Date Booking made"
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
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblBookingDate 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   11
         Top             =   600
         Width           =   3000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   7320
      TabIndex        =   0
      Top             =   960
      Width           =   1815
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1500
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1500
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1500
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   1500
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   1500
      End
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Package Booking Update"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1800
      TabIndex        =   48
      Top             =   120
      Width           =   4530
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
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Width           =   480
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
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "PackageBookingUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global declaration of a string data and an integer data.
Dim BID As String
Dim OldNumberPeople As Integer
 
'Package name combobox click event.
Private Sub cboPackageName_Click()
    Set rs = New Recordset
'The fields Number of people and Net total are emptied if package name is changed in case they need to be updated as well.
    txtNPeople.Text = ""
    txtNetTotal.Text = ""
'Loading cost per person from Package Entry database.
    rs.Open "select * from Package_entry where Package_name = '" & cboPackageName.Text & "'", con, adOpenDynamic, adLockOptimistic
    txtPackageID.Text = rs.Fields("PackageID")
    txtCostPerPerson.Text = rs.Fields("Cost_per_person")
    lblVacancies.Caption = rs.Fields("NumberOfReservations") - rs.Fields("NumberReserved")
    lblItinerary.Caption = rs.Fields("Itinerary")
    lblPackageDetails.Caption = rs.Fields("Package_details")
    lblJourneyDate.Caption = rs.Fields("Journey_date")
    lblArrivalTime.Caption = rs.Fields("Arrival_time")
    lblStartingTime.Caption = rs.Fields("Starting_time")
    lblAccommodation.Caption = rs.Fields("Accommodation_type")
    lblDuration.Caption = rs.Fields("Duration/days")
    
End Sub
'Package type combobox click event.
Private Sub cboPackageType_Click()
'The Package name combobox is cleared when a new package type is selected.
    cboPackageName.Clear
    cboPackageName.Text = ""

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
    lblItinerary.Caption = ""
    lblPackageDetails.Caption = ""
    
'The Package names under the selected Package type are loaded.
    Set rs = New Recordset
    rs.Open "select * from Package_entry where Package_type='" & cboPackageType.Text & "'", con, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        cboPackageName.AddItem (rs.Fields("Package_name"))
        rs.MoveNext
    Loop
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Delete button click event.
Private Sub cmdDelete_Click()
'Validation check to make sure none of the fields are left empty.
    If txtBookingID.Text = Empty Or lblName.Caption = Empty Or txtAddress.Text = Empty Or cboPackageType.Text = Empty Or txtContactNumber.Text = Empty Or txtCostPerPerson.Text = Empty Or cboPackageName.Text = Empty Or txtNPeople.Text = Empty Or txtNetTotal = Empty Or txtPackageID.Text = Empty Then
'Message shown to notify the user that none of the fields can be left empty.
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If
'The selected record is deleted.
    rs.Delete
'Message shown to notify user that the record has been deleted successfully.
    MsgBox "The Booking has been deleted."
    rs.Close
'The reservation data is updated in Package Entry database after deletion of a booking.
    rs.Open "select * from Package_entry where Package_name ='" & cboPackageName.Text & "'", con, adOpenDynamic, adLockOptimistic
        Dim NReserve As Integer
        NReserve = rs.Fields("NumberReserved")
        NReserve = NReserve - txtNPeople.Text
        rs.Fields("NumberReserved") = NReserve
    rs.Update
'The fields are emptied after deletion of the record.
    cboPackageName.Text = ""
    txtBookingID.Text = ""
    lblName.Caption = ""
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
    lblBookingDate.Caption = ""
    
    Call Form_Load
    
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'The fields are cleared and the form is refreshed on the click of a button.
    cboPackageName.Text = ""
    txtBookingID.Text = ""
    lblName.Caption = ""
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
    lblItinerary.Caption = ""
    lblPackageDetails.Caption = ""
    
    cboPackageType.Clear
    cboPackageName.Clear
    
    Call Form_Load
    
End Sub
'Search button click event.
Private Sub cmdSearch_Click()
'Making sure Booking ID is provided.
    If txtBookingID.Text = "" Then
        MsgBox "Please provide Booking ID and try again."
        Exit Sub
    End If
    
'The Package booking database is searched for a paticular Booking ID.
    Set rs = New Recordset
    rs.Open "select * from Package_booking", con, adOpenDynamic, adLockOptimistic
    Dim found As Boolean
    found = False
    rs.MoveFirst
    Do While Not rs.EOF And Not found
        If rs.Fields("PackageBookingID") = txtBookingID.Text Then
        found = True
        End If
        rs.MoveNext
    Loop
    
    If found = False Then
'Message shown to notify User if Booking ID not found.
        MsgBox "There is no record for this Booking ID."
        Exit Sub
    End If
    rs.Close

'If booking ID found, the information the record containing this Booking ID holds is loaded.
    rs.Open "select * from Package_booking where PackageBookingID ='" & txtBookingID.Text & "'", con, adOpenDynamic, adLockOptimistic
        cboPackageType.Text = rs.Fields("Package_type")
        txtPackageID.Text = rs.Fields("PackageID")
        lblName.Caption = rs.Fields("Customer_name")
        txtAddress.Text = rs.Fields("Customer_address")
        lblJourneyDate.Caption = rs.Fields("Journey_date")
        txtContactNumber.Text = rs.Fields("Contact")
        txtCostPerPerson.Text = rs.Fields("Cost_per_person")
        cboPackageName.Text = rs.Fields("Package_name")
        txtNPeople.Text = rs.Fields("Number_of_people")
        txtNetTotal.Text = rs.Fields("Net_total")
        lblBookingDate.Caption = rs.Fields("Date_of_booking")
        lblItinerary.Caption = rs.Fields("Itinerary")
        lblPackageDetails.Caption = rs.Fields("Package_details")
        lblAccommodation.Caption = rs.Fields("Accommodation_type")
        lblJourneyDate.Caption = rs.Fields("Journey_date")
        lblArrivalTime.Caption = rs.Fields("Arrival_time")
        lblStartingTime.Caption = rs.Fields("Starting_time")
        lblDuration.Caption = rs.Fields("Duration/days")
        lblItinerary.Caption = rs.Fields("Itinerary")
        lblPackageDetails.Caption = rs.Fields("Package_details")
    rs.Update

'Values are assigned to the variables BID and OldNumberPeople.
    BID = txtBookingID.Text
    OldNumberPeople = txtNPeople.Text
End Sub
'Update button click event.
Private Sub cmdUpdate_Click()

'Making sure a record is opened to cause alterations.
    If txtBookingID.Text = "" Then
        MsgBox "Please provide Booking ID, press search and try again."
        Exit Sub
    End If
    
'Validation check to make sure Booking ID is not changed.
     If txtBookingID.Text <> BID Then
        MsgBox "Booking ID cannot be changed."
        txtBookingID = BID
        Exit Sub
    End If
    
'Presence check to make sure all the required fields are filled by the customer.
    If cboPackageType.Text = "" Or cboPackageName.Text = "" Or txtPackageID.Text = "" Or txtCostPerPerson.Text = "" Or txtNPeople = "" Or txtAddress.Text = "" Or txtContactNumber.Text = "" Then
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
'Validation check to ensure the updated reservations doesnot exceed the reservation limit.
     rs.Open "select * from Package_entry where Package_name ='" & cboPackageName.Text & "'", con, adOpenDynamic, adLockOptimistic
        Dim NewNumberReserved As Integer
        NewNumberReserved = rs.Fields("NumberReserved")
       
        NewNumberReserved = NewNumberReserved - OldNumberPeople + txtNPeople.Text
      
        If NewNumberReserved > rs.Fields("NumberOfReservations") Then
            MsgBox "There are not enough vacancies to make this reservation."
            Exit Sub
        End If
        
        If NewNumberReserved < rs.Fields("NumberOfReservations") Then
            rs.Fields("NumberReserved") = NewNumberReserved
        rs.Update
        End If
        
    rs.Close
'Validation check to ensure none of the fields  are left empty.
    If txtBookingID.Text = Empty Or lblName.Caption = Empty _
    Or txtAddress.Text = Empty Or cboPackageType.Text = Empty _
    Or txtContactNumber.Text = Empty Or txtCostPerPerson.Text = Empty _
    Or cboPackageName.Text = Empty Or txtNPeople.Text = Empty _
    Or txtNetTotal = Empty Or txtPackageID.Text = Empty Then
'Message shown to notify user that none of the fields can be kept empty.
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If
'The fields are added to the database.
     rs.Open "select * from Package_booking", con, adOpenDynamic, adLockOptimistic
        rs.Fields("Customer_name") = lblName.Caption
        rs.Fields("Customer_address") = txtAddress.Text
        rs.Fields("Package_type") = cboPackageType.Text
        rs.Fields("Journey_Date") = lblJourneyDate.Caption
        rs.Fields("Contact") = txtContactNumber.Text
        rs.Fields("Cost_per_person") = txtCostPerPerson.Text
        rs.Fields("Package_name") = cboPackageName.Text
        rs.Fields("Number_of_people") = txtNPeople.Text
        rs.Fields("Net_total") = txtNetTotal.Text
        rs.Fields("Date_of_booking") = lblDate.Caption
        rs.Fields("Accommodation_type") = lblAccommodation.Caption
        rs.Fields("Journey_date") = lblJourneyDate.Caption
        rs.Fields("Starting_time") = lblStartingTime.Caption
        rs.Fields("Arrival_time") = lblArrivalTime.Caption
        rs.Fields("Duration/days") = lblDuration.Caption
    rs.Update
    
    MsgBox "Booking information has been Updated successfully."
'The fields on the form are emptied.
    cboPackageName.Text = ""
    txtBookingID.Text = ""
    lblName.Caption = ""
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
    lblItinerary.Caption = ""
    lblPackageDetails.Caption = ""
    lblBookingDate.Caption = ""
    
    Call Form_Load
    
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
'Loading Package types from database.
    Set rs = New Recordset
    rs.Open "select * from Package_entry", con, adOpenDynamic, adLockOptimistic
 
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
'Booking ID keyup event.
Private Sub txtBookingID_KeyUp(KeyCode As Integer, Shift As Integer)
'Shifts to Search button when "Enter" pressed.
    If KeyCode = vbKeyReturn Then
        cmdSearch.SetFocus
    End If
End Sub
'Contact number keypress event.
Private Sub txtContactNumber_KeyPress(KeyAscii As Integer)
'Character type check for numerical input.
   If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
'Cost per person keypress event.
Private Sub txtCostPerPerson_KeyPress(KeyAscii As Integer)
'Character type check for numerical input.
   If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Net total keypress event.
Private Sub txtNetTotal_KeyPress(KeyAscii As Integer)
'Character type check for numerical input.
  If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Number of people keypress event.
Private Sub txtNPeople_KeyPress(KeyAscii As Integer)
'Character type check for numerical input.
     If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Number of people Keyup event.
Private Sub txtNPeople_KeyUp(KeyCode As Integer, Shift As Integer)
'Character type check for numerical input.
    If KeyCode = vbKeyReturn Then
        cmdUpdate.SetFocus
    End If
    
End Sub
'Number of people Lostfocus event.
Private Sub txtNPeople_LostFocus()
'Calculation of Net total.
    txtNetTotal.Text = txtNPeople.Text * txtCostPerPerson.Text
End Sub

