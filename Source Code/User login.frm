VERSION 5.00
Begin VB.Form UserLogi 
   Caption         =   "User Log-in"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdUserLogin 
         Caption         =   "Log-in"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.ComboBox cboUserType 
      Height          =   315
      ItemData        =   "User login.frx":0000
      Left            =   1680
      List            =   "User login.frx":000D
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User type"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "UserLogi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboUserType_Click()
    rs.Open "select * from User_login", con, adOpenDynamic, adLockOptimistic
    
    found = False
    Do While Not rs.EOF Or found = True
        If cboUserType.Text = rs.Fields("User_type") Then
            found = True
        End If
    Loop
    
    If found = False Then
        MsgBox "There is no User ID for this User type."
        cboUserType.Text = ""
        Exit Sub
    End If
        
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUserLogin_Click()

    If cboUserType.Text = Empty Or txtUserID.Text = Empty Or txtPassword.Text = Empty Then
        MsgBox "User type or User ID or Password cannot be empty"
        Exit Sub
    End If
    
    Set rs = New Recordset
    rs.Open "select * from User_login where User_type='" & cboUserType.Text & "'", con, adOpenDynamic, adLockOptimistic
    
    Dim found As Boolean
    found = False
    
    rs.MoveFirst
    Do While rs.EOF = False And Not found
        If rs.Fields("User_ID") = txtUserID.Text Then
            found = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    
    If Not found Then
        MsgBox ("User ID not found")
        Exit Sub
    End If
    
    If txtPassword.Text = rs.Fields("Password") Then
        MsgBox "Login successful"
    Else
        MsgBox "Wrong password"
    End If
    
    
End Sub

Private Sub Form_Load()
    Call dblink
End Sub
