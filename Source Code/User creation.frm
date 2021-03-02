VERSION 5.00
Begin VB.Form usercreation 
   Caption         =   "User Account Creation"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form2"
   ScaleHeight     =   3810
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   5160
      TabIndex        =   8
      Top             =   480
      Width           =   2055
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
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
         Height          =   450
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
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
         Height          =   450
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdCreateAccount 
         Caption         =   "Create account"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
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
      ItemData        =   "User creation.frx":0000
      Left            =   1560
      List            =   "User creation.frx":000D
      TabIndex        =   7
      Top             =   1320
      Width           =   3375
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
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2520
      Width           =   3375
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
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtConfirmPassword 
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
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3120
      Width           =   3375
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
      Left            =   5040
      TabIndex        =   13
      Top             =   120
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
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "User Creation"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Confirm password"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
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
      Left            =   120
      TabIndex        =   2
      Top             =   2520
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
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
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
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "usercreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Decalaration of global variables.
Dim UType, UID As String

'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Create Account click event.
Private Sub cmdCreateAccount_Click()
'Check to make sure all important fields are filled.
    If cboUserType.Text = Empty Or txtUserID.Text = Empty Or txtPassword.Text = Empty Or txtConfirmPassword.Text = Empty Then
        MsgBox "None of the fields can be kept empty."
        Exit Sub
    End If
'Storing information into the database.
    Set rs = New Recordset
    rs.Open "select * from User_login", con, adOpenDynamic, adLockOptimistic
 'Checking if User ID is already present so that same ID is not saved twice.
    Dim IsNewID As Boolean
    IsNewID = True
    If Not rs.EOF Then
        Do While Not rs.EOF
            If txtUserID.Text = rs.Fields("User_ID") Then
                IsNewID = False
                Exit Do
            End If
            rs.MoveNext
        Loop
    End If
    
    If IsNewID = False Then
        MsgBox "This User ID is already in use. Please try again."
        txtUserID.Text = ""
        Exit Sub
    End If
    
    
    rs.AddNew
    'Check made to see if password and confirm password match if User type and User ID are Valid.
    If txtPassword.Text = txtConfirmPassword.Text Then
        rs.Fields("Password") = txtPassword.Text
    Else
        MsgBox ("The password and the confirm password do not match. Please re-enter them and try again")
        txtPassword.Text = ""
        txtConfirmPassword.Text = ""
        Exit Sub
    End If
    
        rs.Fields("User_type") = cboUserType.Text
        rs.Fields("User_ID") = txtUserID.Text
    
   
         
  rs.Update
  
  MsgBox ("User account has been created successfully.")
  'The fields are emptied after the information is stored.
  cboUserType.Text = ""
  txtUserID.Text = ""
  txtPassword.Text = ""
  txtConfirmPassword = ""
  
End Sub
'Delete button click event.
Private Sub cmdDelete_Click()
'Check to make sure all important fields are filled.
    If cboUserType.Text = Empty Or txtUserID.Text = Empty Or txtPassword.Text = Empty Or txtConfirmPassword.Text = Empty Then
        MsgBox "None of the fields can be kept empty."
        Exit Sub
    End If
'The selected record is deleted.
    rs.Delete
    MsgBox " The record has been deleted."
'The fields on the form are cleared after deletion.
    cboUserType.Text = ""
    txtUserID.Text = ""
    txtPassword.Text = ""
    txtConfirmPassword = ""
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'The fields are cleared.
  cboUserType.Text = ""
  txtUserID.Text = ""
  txtPassword.Text = ""
  txtConfirmPassword = ""
End Sub
'Search button click event.
Private Sub cmdSearch_Click()
'Fields are checked to see if any of the fields are empty.
     If cboUserType.Text = Empty Or txtUserID.Text = Empty Then
        MsgBox "Please provide User type and User ID both."
        Exit Sub
    End If
'Information is loaded from the database.
    Set rs = New Recordset
    rs.Open "select * from User_login where User_type ='" & cboUserType.Text & "'", con, adOpenDynamic, adLockOptimistic
    
    Dim found As Boolean
    found = False
'The database is checked to see if the User ID is present.
    Do While rs.EOF = False And Not found
        If rs.Fields("User_ID") = txtUserID.Text Then
            found = True
            Exit Do
        End If
        rs.MoveNext
    Loop
'Message if database is not found.
    If Not found Then
        MsgBox ("User ID not found. Please check your User type and User ID and try again.")
        Exit Sub
    End If

    txtPassword.Text = rs.Fields("Password")
    txtConfirmPassword.Text = txtPassword.Text
    
     
        UType = cboUserType.Text
        UID = txtUserID.Text

End Sub

Private Sub cmdUpdate_Click()
'The fields are checked to see if they are all filled.
     If cboUserType.Text = Empty Or txtUserID.Text = Empty Or txtPassword.Text = Empty Or txtConfirmPassword.Text = Empty Then
        MsgBox "None of the fields can be kept empty."
        Exit Sub
    End If
'Validation to ensure User type is not changed.
        If cboUserType.Text <> UType Then
            MsgBox "User type cannot be changed."
            cboUserType.Text = UType
            Exit Sub
        End If
 'Validation to check if User ID is already present if User plans on changing the User ID.
        If txtUserID <> UID Then
             If Not rs.EOF Then
                Dim IsNewID As Boolean
                    IsNewID = True
                    Do While Not rs.EOF
                        If txtUserID.Text = rs.Fields("User_ID") Then
                            IsNewID = False
                            Exit Do
                        End If
                    rs.MoveNext
                Loop
            End If
            
            If IsNewID = False Then
                MsgBox "This User ID is already in use. Please try again."
                txtUserID.Text = ""
                Exit Sub
            End If
        End If
        
    
'User information in the database is updated.
    rs.MoveFirst
       
'Validation to make sure password and confirm password match.
    If txtPassword.Text = txtConfirmPassword.Text Then
        rs.Fields("Password") = txtPassword.Text
    Else
'Message shown if password and confirm password don't match
        MsgBox ("The password and the confirm password do not match. Please check them and try again")
        Exit Sub
    End If
         
          rs.Fields("User_ID") = txtUserID.Text
          
  rs.Update
'Message shown to notify the User that the Update was successful.
  MsgBox ("Record has been Updated successfully.")
'The fields are cleared after the Update is successful.
  cboUserType.Text = ""
  txtUserID.Text = ""
  txtPassword.Text = ""
  txtConfirmPassword = ""
  
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
    lblDate.Caption = Date
End Sub

