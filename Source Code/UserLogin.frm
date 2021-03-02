VERSION 5.00
Begin VB.Form UserLogin 
   Caption         =   "User login"
   ClientHeight    =   3285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboUserType 
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
      ItemData        =   "UserLogin.frx":0000
      Left            =   1680
      List            =   "UserLogin.frx":0010
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtUserID 
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4095
      Begin VB.CommandButton cmdUserLogin 
         Caption         =   "Log-in"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label15 
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
      Left            =   2880
      TabIndex        =   11
      Top             =   360
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
      Left            =   3480
      TabIndex        =   10
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "User Log-in"
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
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "User type"
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
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "UserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'User type combobox Click event.
Private Sub cboUserType_Click()
'Fields are cleared if User type is changed.
    txtUserID.Text = ""
    txtPassword.Text = ""
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
    Unload Me
End Sub
'Log-in button click event.
Private Sub cmdUserLogin_Click()
    'Check if any of the fields ar empty.
    If cboUserType.Text = Empty Or txtUserID.Text = Empty Or txtPassword.Text = Empty Then
        MsgBox "User type or User ID or Password cannot be empty"
        Exit Sub
    End If
    
    Set rs = New Recordset
    rs.Open "select * from User_login where User_type='" & cboUserType.Text & "'", con, adOpenDynamic, adLockOptimistic
    'Check if the combination of User type and User ID provided is present in the database.
    Dim found As Boolean
    found = False
    
    Do While rs.EOF = False And Not found
        If rs.Fields("User_ID") = txtUserID.Text Then
            found = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    'Access denial message shown if the User Type and USer ID are not found in the database.
    If Not found Then
    'Message shoen to notify user that the requested User ID is not found.
        MsgBox ("User ID not found. Please check your User type and User ID and try again.")
        Exit Sub
    End If
    'Check for validation of password if combination of User type and User ID are accepted.
    If txtPassword.Text = rs.Fields("Password") Then
        MsgBox "Login successful."
    Else
    'Access denail message shown if password is not accepted.
        MsgBox "Invalid password. Please try again."
    End If
    
    MenuForm.Show
End Sub
'Form load event.
Private Sub Form_Load()
'Calling the function to connect the database.
    Call dblink
    lblDate.Caption = Date
End Sub

