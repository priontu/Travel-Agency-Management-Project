VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VehicleServicing 
   Caption         =   "Vehicle servicing"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3720
      TabIndex        =   12
      Top             =   3240
      Width           =   5175
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
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1560
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
         Height          =   435
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   1560
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
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.ListBox lstVehicleID 
      Height          =   4155
      ItemData        =   "Car servicing.frx":0000
      Left            =   240
      List            =   "Car servicing.frx":0002
      TabIndex        =   8
      Top             =   1680
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid flxgrdHistory 
      Height          =   1455
      Left            =   3720
      TabIndex        =   7
      Top             =   4440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedRows       =   0
   End
   Begin VB.TextBox txtServicingDetails 
      Height          =   675
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   3720
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   2400
   End
   Begin VB.ComboBox cboVehicleType 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Vehicle Servicing"
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
      Left            =   3120
      TabIndex        =   16
      Top             =   240
      Width           =   3255
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
      Left            =   7560
      TabIndex        =   11
      Top             =   840
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
      Left            =   6960
      TabIndex        =   10
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Vehi 
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
      TabIndex        =   9
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Servicing History Preview"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   4200
      Width           =   2235
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Servicing details"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label Label3 
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
      TabIndex        =   0
      Top             =   960
      Width           =   1065
   End
End
Attribute VB_Name = "VehicleServicing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vehicle type combobox click event.
Private Sub cboVehicleType_Click()
'Clearing Vehicle ID list before entering new data.
    lstVehicleID.Clear
    txtAmount.Text = ""
    txtServicingDetails.Text = ""
'Loading Vehicle IDs onto the Vehicle ID list.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where Vehicle_type = '" & cboVehicleType.Text & "'", con, adOpenDynamic, adLockOptimistic
        Do While rs.EOF = False
            lstVehicleID.AddItem (rs.Fields("VehicleID"))
            rs.MoveNext
        Loop
    rs.Close
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Closing the form.
    Unload Me
End Sub
'Refresh buton click event.
Private Sub cmdRefresh_Click()
'Clearing the fields as the User requires in case the User wants to start anew.
    cboVehicleType.Text = ""
    lstVehicleID.Clear
    txtAmount.Text = ""
    txtServicingDetails.Text = ""
    
End Sub
'Save button click event.
Private Sub cmdSave_Click()
'Presence check to make sure the User provides all the necessary information.
    If cboVehicleType.Text = "" Or lstVehicleID.Text = "" Or txtAmount.Text = "" Or txtServicingDetails.Text = "" Then
        MsgBox "Some of the required fields are left empty. Please fill them up and try again."
        Exit Sub
    End If
'Storing information into the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_servicing", con, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("Vehicle_type") = cboVehicleType.Text
            rs.Fields("VehicleID") = lstVehicleID.Text
            rs.Fields("Servicing_details") = txtServicingDetails.Text
            rs.Fields("Amount") = txtAmount.Text
            rs.Fields("Date") = Date
        rs.Update
'Clearing the fields to input new data.
    txtAmount.Text = ""
    txtServicingDetails.Text = ""
    flxgrdHistory.Rows = 2
    Call lstVehicleID_Click
End Sub
'Form load event.
Private Sub Form_Load()
'Calling funtion to connect to database.
    Call dblink
'Showing current date on the form.
lblDate.Caption = Date
'Loading the Vehicle type information onto the Vehicle type combobox from Vehicle entry database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
'Validity check to make sure each type is loaded only once.
    If rs.EOF = False Then
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To cboVehicleType.ListCount - 1
                If cboVehicleType.List(i) = rs.Fields("Vehicle_type") Then
                    IsNewType = False
                End If
            Next
            If IsNewType = True Then
                cboVehicleType.AddItem (rs.Fields("Vehicle_type"))
            End If
            rs.MoveNext
        Loop
    End If
        flxgrdHistory.TextMatrix(0, 0) = "Date"
        flxgrdHistory.TextMatrix(0, 1) = "Servicing Details"
        flxgrdHistory.TextMatrix(0, 2) = "Amount"
        flxgrdHistory.ColWidth(1) = 3000
        flxgrdHistory.RowHeight(1) = 500
End Sub
'Vehicle ID list click event.
Private Sub lstVehicleID_Click()
    Dim r As Integer
    r = 0
    txtAmount.Text = ""
    txtServicingDetails.Text = ""
    Set rs = New Recordset
    rs.Open "select * from Vehicle_servicing where VehicleID ='" & lstVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
         r = r + 1

         flxgrdHistory.TextMatrix(r, 0) = rs!Date
         flxgrdHistory.TextMatrix(r, 1) = rs!Servicing_details
         flxgrdHistory.TextMatrix(r, 2) = rs!Amount

         rs.MoveNext
         flxgrdHistory.Rows = flxgrdHistory.Rows + 1
    Wend
End Sub
'Amount keypress event.
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is input.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
