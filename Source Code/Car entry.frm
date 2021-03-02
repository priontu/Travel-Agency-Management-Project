VERSION 5.00
Begin VB.Form VehicleEntry 
   Caption         =   "Vehicle Entry"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form3"
   ScaleHeight     =   3960
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRatePerKm 
      Height          =   315
      Left            =   6600
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1680
      TabIndex        =   13
      Top             =   2760
      Width           =   6015
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
         Height          =   495
         Left            =   2160
         TabIndex        =   16
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
         Height          =   495
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtVehicleID 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtRegNum 
      Height          =   315
      Left            =   6600
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtMake 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtModel 
      Height          =   315
      Left            =   6600
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtNSeats 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox cboVehicleType 
      Height          =   315
      ItemData        =   "Car entry.frx":0000
      Left            =   1560
      List            =   "Car entry.frx":000D
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle Entry"
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
      Left            =   3600
      TabIndex        =   20
      Top             =   240
      Width           =   2295
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
      Left            =   8280
      TabIndex        =   19
      Top             =   480
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
      Left            =   7680
      TabIndex        =   18
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label7 
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
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   4800
      TabIndex        =   5
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label Label5 
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
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rate per kilometer"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Label Label1 
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
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "VehicleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cancel button click event.
Private Sub cmdCancel_Click()
'Closing the form.
    Unload Me
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'Clearing the fields on the form as required by the User.
    cboVehicleType.Text = ""
    txtVehicleID.Text = ""
    txtRegNum.Text = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtNSeats.Text = ""
    txtRatePerKm = ""
    Call Form_Load
End Sub
'Submit button click event.
Private Sub cmdsubmit_Click()
'Presence check to ensure all the required information is provided.
    If cboVehicleType.Text = Empty Or txtVehicleID.Text = Empty Or txtRegNum.Text = Empty Or txtMake.Text = Empty Or txtModel.Text = Empty Or txtNSeats.Text = Empty Or txtRatePerKm.Text = Empty Then
'Message shown to notify the User that none of the fields can be left empty.
            MsgBox "None of the fields can be left empty"
            Exit Sub
    End If
    
    Set rs = New Recordset
    
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
'Validation check to ensure all Vehicle IDs are unique and none of the Vehicle IDs are used twice.
    Dim IsNewID As Boolean
    IsNewID = True
            
        If Not rs.EOF Then
            Do While Not rs.EOF
                If txtVehicleID.Text = rs.Fields("VehicleID") Then
                    IsNewID = False
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
    
    If IsNewID = False Then
'Message shown if Vehicle ID is already in use.
        MsgBox "This Vehicle ID is already in use. Please try again."
        txtVehicleID.Text = ""
        Exit Sub
    End If
    rs.MoveFirst
    Dim IsNewID2 As Boolean
    IsNewID2 = True
    
     If Not rs.EOF Then
        
            Do While Not rs.EOF
                If txtRegNum.Text = rs.Fields("RegistrationNo") Then
                    IsNewID2 = False
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
    
    If IsNewID2 = False Then
        MsgBox "This Registration number is already in use. Please try again."
        txtRegNum.Text = ""
        Exit Sub
    End If

'Adding new Vehicle information into the database.
    rs.AddNew
        rs.Fields("Vehicle_type") = cboVehicleType.Text
        rs.Fields("VehicleID") = txtVehicleID.Text
        rs.Fields("RegistrationNo") = txtRegNum.Text
        rs.Fields("Make") = txtMake.Text
        rs.Fields("Model") = txtModel.Text
        rs.Fields("NSeats") = txtNSeats.Text
        rs.Fields("Rate_per_km") = txtRatePerKm.Text
        rs.Fields("Date_of_entry") = lblDate.Caption
    rs.Update
    
'Message shown to notify the the User that new vehicle information is added successfully.
    MsgBox "Vehicle information has been added successfully"
'Clearing the fields to ready the form for new transaction.
    cboVehicleType.Text = ""
    txtVehicleID.Text = ""
    txtRegNum.Text = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtNSeats.Text = ""
    txtRatePerKm = ""
    
    Call Form_Load
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Showing the current date on the form.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
    rs.MoveLast
    txtVehicleID.Text = CInt(rs!VehicleID) + 1

End Sub
'Number of seats Keypress event.
Private Sub txtNSeats_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
'Rate per kilometer Keypress event.
Private Sub txtRatePerKm_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
