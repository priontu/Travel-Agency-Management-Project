VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PackageEntry 
   Caption         =   "Package entry"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form5"
   ScaleHeight     =   9075
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNRes 
      Height          =   315
      Left            =   3240
      TabIndex        =   38
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtPackageDetails 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   7080
      Width           =   9375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Package information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   4575
      Begin VB.TextBox txtPackageID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox cboPackageType 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Package entry.frx":0000
         Left            =   1680
         List            =   "Package entry.frx":000A
         TabIndex        =   33
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cboAccommodation 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Package entry.frx":0025
         Left            =   1680
         List            =   "Package entry.frx":0032
         TabIndex        =   28
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtPackageName 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtDuration 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtCostPerPerson 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   1440
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpArrivalTime 
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   234749954
         CurrentDate     =   41961
      End
      Begin MSComCtl2.DTPicker dtpStartingTime 
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   234749954
         CurrentDate     =   41961
      End
      Begin MSComCtl2.DTPicker dtpJourneyDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   234749953
         CurrentDate     =   41961
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Package ID"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Accommodation type"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Arrival time"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Journey Date"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Days"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   25
         Top             =   3360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Package type"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cost per person"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Starting time"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Name of package"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   8040
      Width           =   9135
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
         Height          =   405
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1665
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
         Height          =   405
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   1665
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
         Height          =   405
         Left            =   5520
         TabIndex        =   8
         Top             =   240
         Width           =   1665
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
         Height          =   405
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   4575
      Begin VB.ComboBox cboPackageTypeUpdate 
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
         ItemData        =   "Package entry.frx":004E
         Left            =   1320
         List            =   "Package entry.frx":0050
         TabIndex        =   39
         Top             =   360
         Width           =   3015
      End
      Begin VB.ListBox lstPackageNameUpdate 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         ItemData        =   "Package entry.frx":0052
         Left            =   1320
         List            =   "Package entry.frx":0054
         TabIndex        =   5
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Package Name list"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Package type"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtItinerary 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5760
      Width           =   9375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Number of reservations available?"
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
      Left            =   240
      TabIndex        =   37
      Top             =   5040
      Width           =   2910
   End
   Begin VB.Label Label15 
      Caption         =   "Package Entry"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   36
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Package details"
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
      Left            =   240
      TabIndex        =   32
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8160
      TabIndex        =   13
      Top             =   360
      Width           =   960
   End
   Begin VB.Label Label14 
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
      Left            =   7560
      TabIndex        =   12
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   0
      Top             =   5400
      Width           =   675
   End
End
Attribute VB_Name = "PackageEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldPID As String
'Click event Package type combobox of the update section of the form.
Private Sub cboPackageTypeUpdate_Click()
'The Package name list cleared if package type is changed.
    lstPackageNameUpdate.Clear
    Set rs = New Recordset
'Opening the Package Entry database.
    rs.Open "select * from Package_entry where Package_type ='" & cboPackageTypeUpdate.Text & "'", con, adOpenDynamic, adLockOptimistic
'Loading the Package names of the particular type selected into the Package name list.
    Do While Not rs.EOF
        lstPackageNameUpdate.AddItem (rs.Fields("Package_name"))
        rs.MoveNext
    Loop
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Delete button click event
Private Sub cmdDelete_Click()
  Set rs = New Recordset
  'Validation check to make sure none of the important fields are left empty
    If cboPackageType.Text = Empty Or txtPackageName.Text = Empty Or txtPackageDetails.Text = Empty Or txtCostPerPerson.Text = Empty Or cboAccommodation = Empty Or dtpJourneyDate.Value = Date Or dtpStartingTime = Default Or dtpArrivalTime.Value = Default Or txtDuration.Text = Empty Or txtItinerary.Text = Empty Then
'Message shown if any empty fields are fields are found.
        MsgBox "All the information need to be provided."
        Exit Sub
    End If
'Deletion of selected record with the particular Package name from the Package Entry database.
  rs.Open "delete * from Package_entry where Package_name='" & lstPackageNameUpdate.Text & "' and Package_type='" & cboPackageTypeUpdate.Text & "'", con, adOpenDynamic, adLockOptimistic

    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh
    cboPackageTypeUpdate.Clear
    cboPackageTypeUpdate.Refresh
    
    cboPackageTypeUpdate.Text = ""
    cboPackageType.Text = ""
    txtPackageName.Text = ""
    txtCostPerPerson.Text = ""
    cboAccommodation.Text = ""
    dtpJourneyDate.Value = Default
    dtpStartingTime.Value = Default
    dtpArrivalTime.Value = Default
    txtDuration.Text = ""
    txtItinerary.Text = ""
    txtPackageDetails.Text = ""
    txtPackageID.Text = ""
    txtNRes.Text = ""
    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh
    MsgBox "Record has been deleted successfully."
    
    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh

   Call Form_Load
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'Emptying all the fields and refreshing the forms if the user requires it.
    cboPackageTypeUpdate.Text = ""
    cboPackageType.Text = ""
    txtPackageName.Text = ""
    txtCostPerPerson.Text = ""
    cboAccommodation.Text = ""
    dtpJourneyDate.Value = Default
    dtpStartingTime.Value = Default
    dtpArrivalTime.Value = Default
    txtDuration.Text = ""
    txtItinerary.Text = ""
    txtPackageDetails.Text = ""
    txtPackageID.Text = ""
    txtNRes.Text = ""
    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh
    
   Call Form_Load
   
End Sub
'Save button click event.
Private Sub cmdSave_Click()
'Validation check to make sure none of the important fields are left empty
    If cboPackageType.Text = Empty Or txtPackageName.Text = Empty Or txtPackageDetails.Text = Empty Or txtCostPerPerson.Text = Empty Or cboAccommodation = Empty Or dtpJourneyDate.Value = Date Or dtpStartingTime = Default Or dtpArrivalTime.Value = Default Or txtDuration.Text = Empty Or txtItinerary.Text = Empty Then
'Message shown if any empty fields are fields are found.
        MsgBox "All the information need to be provided."
        Exit Sub
    End If
'Validation check to make sure same Package ID is not used twice.
     If Not rs.EOF Then
        Dim IsOldID As Boolean
            IsOldID = False
            Do While Not rs.EOF
                If txtPackageID.Text = rs.Fields("PackageID") Then
                    IsOldID = True
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
    
    If IsOldID = True Then
        MsgBox "This Package ID is already in use. Please provide a differnt Package ID."
        txtPackageID.Text = ""
        Exit Sub
    End If
'Storing Package information into the database.
     Set rs = New Recordset
      rs.Open "select * from Package_entry", con, adOpenDynamic, adLockOptimistic
      rs.AddNew
        rs.Fields("Package_type") = cboPackageType.Text
        rs.Fields("Package_name") = txtPackageName.Text
        rs.Fields("Accommodation_type") = cboAccommodation.Text
        rs.Fields("Journey_date") = dtpJourneyDate.Value
        rs.Fields("Starting_time") = dtpStartingTime.Value
        rs.Fields("Duration/days") = txtDuration.Text
        rs.Fields("Arrival_time") = dtpArrivalTime.Value
        rs.Fields("Cost_per_person") = txtCostPerPerson.Text
        rs.Fields("Date_of_entry") = lblDate.Caption
        rs.Fields("Itinerary") = txtItinerary.Text
        rs.Fields("Package_details") = txtPackageDetails.Text
        rs.Fields("PackageID") = txtPackageID.Text
        rs.Fields("NumberOfReservations") = txtNRes.Text
        
      rs.Update
'Message shown to notify the user that the package information have been added successfully.
    MsgBox "Record has been added successfully"
'Emptying the fields after the information is added to the database.
    cboPackageTypeUpdate.Text = ""
    cboPackageType.Text = ""
    txtPackageName.Text = ""
    txtCostPerPerson.Text = ""
    cboAccommodation.Text = ""
    dtpJourneyDate.Value = Default
    dtpStartingTime.Value = Default
    dtpArrivalTime.Value = Default
    txtDuration.Text = ""
    txtItinerary.Text = ""
    txtPackageDetails.Text = ""
    txtPackageID.Text = ""
    txtNRes.Text = ""
    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh
    
    rs.MoveLast
    txtPackageID.Text = CInt(rs!PackageID) + 1
    
'Reloading the names on the Package Type combobox of the update section to make sure that the new names appear immediately.
        Set rs = New Recordset
        rs.Open "select * from Package_entry", con, adOpenDynamic, adLockOptimistic
        
        If rs.EOF = False Then
           Do While rs.EOF = False
               Dim i As Integer
               Dim IsNewType As Boolean
               IsNewType = True
               For i = 0 To cboPackageTypeUpdate.ListCount - 1
                   If cboPackageTypeUpdate.List(i) = rs.Fields("Package_type") Then
                       IsNewType = False
                   End If
               Next
               If IsNewType = True Then
                   cboPackageTypeUpdate.AddItem (rs.Fields("Package_type"))
               End If
               rs.MoveNext
           Loop
        End If
        
        
End Sub
'Update button click event.
Private Sub cmdUpdate_Click()
'Validation check to make sure none of the fields are left empty.
    If cboPackageType.Text = Empty Or txtPackageName.Text = Empty Or txtPackageDetails.Text = Empty Or txtCostPerPerson.Text = Empty Or cboAccommodation = Empty Or dtpJourneyDate.Value = Default Or dtpStartingTime = Default Or dtpArrivalTime.Value = Default Or txtDuration.Text = Empty Or txtItinerary.Text = Empty Then
        MsgBox "All the information need to be provided."
        Exit Sub
    End If
'Ensuring Package ID is not changed.
    If OldPID <> txtPackageID.Text Then
        MsgBox "PackageID cannot be changed."
        txtPackageID.Text = OldPID
        Exit Sub
    End If
    
'Updating information on the database.
        rs.Fields("Package_type") = cboPackageType.Text
        rs.Fields("Package_name") = txtPackageName.Text
        rs.Fields("Accommodation_type") = cboAccommodation.Text
        rs.Fields("Journey_date") = dtpJourneyDate.Value
        rs.Fields("Starting_time") = dtpStartingTime.Value
        rs.Fields("Duration/days") = txtDuration.Text
        rs.Fields("Arrival_time") = dtpArrivalTime.Value
        rs.Fields("Cost_per_person") = txtCostPerPerson.Text
        rs.Fields("Date_of_entry") = lblDate.Caption
        rs.Fields("Itinerary") = txtItinerary.Text
        rs.Fields("Package_details") = txtPackageDetails.Text
        rs.Fields("PackageID") = txtPackageID.Text
        rs.Fields("NumberOfReservations") = txtNRes.Text
        rs.Fields("NumberReserved") = 0
    rs.Update
'Message shown to notify user that the information is updated.
    MsgBox "Package information has been updated successfully."
'The fields are updated after the database is updated and the form is refreshed.
    cboPackageTypeUpdate.Text = ""
    cboPackageType.Text = ""
    txtPackageName.Text = ""
    txtCostPerPerson.Text = ""
    cboAccommodation.Text = ""
    dtpJourneyDate.Value = Default
    dtpStartingTime.Value = Default
    dtpArrivalTime.Value = Default
    txtDuration.Text = ""
    txtItinerary.Text = ""
    txtPackageDetails.Text = ""
    txtPackageID.Text = ""
    txtNRes.Text = ""
    lstPackageNameUpdate.Clear
    lstPackageNameUpdate.Refresh
    
    Call Form_Load
    
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Showing current date.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
'Current date assigned to Journey date datepicker.
    dtpJourneyDate.Value = Date
'Loading Package types into the Package type combobox of the update section.
    Set rs = New Recordset
    rs.Open "select * from Package_entry", con, adOpenDynamic, adLockOptimistic
    
    If rs.EOF = False Then
       Do While rs.EOF = False
           Dim i As Integer
           Dim IsNewType As Boolean
           IsNewType = True
           For i = 0 To cboPackageTypeUpdate.ListCount - 1
               If cboPackageTypeUpdate.List(i) = rs.Fields("Package_type") Then
                   IsNewType = False
               End If
           Next
           If IsNewType = True Then
               cboPackageTypeUpdate.AddItem (rs.Fields("Package_type"))
           End If
           rs.MoveNext
       Loop
    End If
       
'Updating Package ID textbox.
    rs.MoveLast
    txtPackageID.Text = CInt(rs!PackageID) + 1
      
End Sub
'Click event of the list of Package names.
Private Sub lstPackageNameUpdate_Click()
'The package information is loaded onto the form at the selection of a Package name.
    Set rs = New Recordset
    rs.Open "select * from Package_entry where Package_name ='" & lstPackageNameUpdate.Text & "'", con, adOpenDynamic, adLockOptimistic
        cboPackageType.Text = rs.Fields("Package_type")
        txtPackageName.Text = rs.Fields("Package_name")
        txtItinerary.Text = rs.Fields("Itinerary")
        txtPackageDetails.Text = rs.Fields("Package_details")
        cboAccommodation.Text = rs.Fields("Accommodation_type")
        dtpStartingTime.Value = rs.Fields("Starting_time")
        txtDuration.Text = rs.Fields("Duration/days")
        dtpArrivalTime.Value = rs.Fields("Arrival_time")
        dtpJourneyDate.Value = rs.Fields("Journey_date")
        txtCostPerPerson.Text = rs.Fields("Cost_per_person")
        txtPackageID.Text = rs.Fields("PackageID")
        txtNRes.Text = rs.Fields("NumberOfReservations")
    rs.Update

    OldPID = txtPackageID.Text
End Sub
'Cost per person keypress event.
Private Sub txtCostPerPerson_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
      If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Duration keypress event.
Private Sub txtDuration_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Number of reservations keypress event.
Private Sub txtNRes_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
