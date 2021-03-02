VERSION 5.00
Begin VB.Form VehicleUpdate 
   Caption         =   "Vehicle Update"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form9"
   ScaleHeight     =   4530
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Vehicle information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   6495
      Begin VB.TextBox txtRate 
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   1200
         Width           =   2880
      End
      Begin VB.TextBox txtNSeats 
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   480
         Width           =   2880
      End
      Begin VB.TextBox txtModel 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   1920
         Width           =   2880
      End
      Begin VB.TextBox txtMake 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2880
      End
      Begin VB.TextBox txtRegNum 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2880
      End
      Begin VB.ComboBox cboVehicleType 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   3360
         TabIndex        =   18
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of seats"
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
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Model"
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
         Left            =   3360
         TabIndex        =   16
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Make"
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
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Registration number"
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
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle type"
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
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   3600
      Width           =   6135
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
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1815
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
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1920
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
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.ListBox lstVehicleID 
      Height          =   3375
      ItemData        =   "Car Update.frx":0000
      Left            =   120
      List            =   "Car Update.frx":0002
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Date of Entry"
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
      Left            =   6960
      TabIndex        =   23
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblEntryDate 
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
      Left            =   8400
      TabIndex        =   22
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle Update"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   21
      Top             =   120
      Width           =   2910
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
      Left            =   8400
      TabIndex        =   20
      Top             =   240
      Width           =   960
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
      Left            =   7680
      TabIndex        =   19
      Top             =   240
      Width           =   600
   End
   Begin VB.Label listvehicleid 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle ID"
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
      TabIndex        =   0
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "VehicleUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cancel button click event.
Private Sub cmdCancel_Click()
'Closing the form.
    Unload Me
End Sub
'Delete button click event.
Private Sub cmdDelete_Click()

'Presence check to make sure the customer provides all the necessary information.
    If lstVehicleID.Text = "" Or cboVehicleType.Text = "" Or txtMake.Text = "" Or txtModel.Text = "" Or txtRegNum.Text = "" Or txtRate.Text = "" Or txtNSeats.Text = "" Then
        MsgBox "All the required information not provided. Please fill up all the fields and try again."
        Exit Sub
    End If
    
'Deletion of the record from the database.
    rs.Delete
    lstVehicleID.Refresh
    MsgBox "Record has been deleted"
    
'The information of the deleted record is cleared from the form.
    cboVehicleType.Text = ""
    txtRegNum = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtNSeats.Text = ""
    txtRate.Text = ""
    lblEntryDate.Caption = ""
    lstVehicleID.Clear
    
    Call Form_Load
  
End Sub
'Update button click event.
Private Sub cmdUpdate_Click()
'Presence check to make sure the customer provides all the necessary information.
    If lstVehicleID.Text = "" Or cboVehicleType.Text = "" Or txtMake.Text = "" Or txtModel.Text = "" Or txtRegNum.Text = "" Or txtRate.Text = "" Or txtNSeats.Text = "" Then
        MsgBox "All the required information not provided. Please fill up all the fields and try again."
        Exit Sub
    End If
    
'The Vehicle information in the database are updated.
        rs.Fields("Vehicle_type") = cboVehicleType.Text
        rs.Fields("RegistrationNo") = txtRegNum.Text
        rs.Fields("Make") = txtMake.Text
        rs.Fields("Model") = txtModel.Text
        rs.Fields("NSeats") = txtNSeats.Text
        rs.Fields("Rate_per_km") = txtRate.Text
    rs.Update
'Message shown to notify the User that the Update was successful.
    MsgBox "Vehicle information has been updated"
'Clearing the fields in the form to ready form for the next job.
    cboVehicleType.Text = ""
    txtRegNum = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtNSeats.Text = ""
    txtRate.Text = ""
    lblEntryDate.Caption = ""
    lstVehicleID.Clear
    
    Call Form_Load
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Showing the current date on the form.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
'Loading the Vehicle ID information into the list of Vehicle IDs from the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        lstVehicleID.AddItem (rs.Fields("VehicleID"))
        rs.MoveNext
    Loop
    
End Sub
'Vehicle ID list click event.
Private Sub lstVehicleID_Click()
'Loading information for the particular Vehicle ID selected into the form from the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where VehicleID = '" & lstVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
    
        cboVehicleType.Text = rs.Fields("Vehicle_type")
        txtRegNum.Text = rs.Fields("RegistrationNo")
        txtMake.Text = rs.Fields("Make")
        txtModel.Text = rs.Fields("Model")
        txtNSeats.Text = rs.Fields("NSeats")
        txtRate.Text = rs.Fields("Rate_per_km")
        lblEntryDate.Caption = rs.Fields("Date_of_entry")
    rs.Update
End Sub
'Number of seats Keypress event.
Private Sub txtNSeats_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is input.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
'Rate per kilometer keypress event.
Private Sub txtRate_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is input.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
