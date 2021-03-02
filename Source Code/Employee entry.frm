VERSION 5.00
Begin VB.Form EmployeeEntry 
   Caption         =   "Employee entry"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Elephant"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   7200
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Employee information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   3615
      Begin VB.ComboBox cboDesig 
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
         ItemData        =   "Employee entry.frx":0000
         Left            =   240
         List            =   "Employee entry.frx":000D
         TabIndex        =   23
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox txtEmployeeName 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtEmployeeID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtEmployeeAddress 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtContactNumber 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtBasicSalary 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Employee ID"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Basic salary"
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
         TabIndex        =   20
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
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
         TabIndex        =   19
         Top             =   3120
         Width           =   1050
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name of Employee"
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
         TabIndex        =   17
         Top             =   960
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   4455
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         TabIndex        =   8
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ListBox lstEmployeeID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         ItemData        =   "Employee entry.frx":002B
         Left            =   120
         List            =   "Employee entry.frx":002D
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Employee ID"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5895
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
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1860
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
         Height          =   495
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1860
      End
      Begin VB.CommandButton cmdMakeEntry 
         Caption         =   "Make entry"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Employee Entry"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Top             =   240
      Width           =   4275
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
      Left            =   5880
      TabIndex        =   10
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label8 
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
      Left            =   5160
      TabIndex        =   9
      Top             =   840
      Width           =   600
   End
End
Attribute VB_Name = "EmployeeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EID As String

'Cancel button click event
Private Sub cmdCancel_Click()
'The form is unloaded.
    Unload Me
End Sub
'Delete button click event.
Private Sub cmdDelete_Click()
    If txtEmployeeID.Text = Empty Or txtEmployeeName.Text = Empty Or txtEmployeeAddress.Text = Empty Or txtContactNumber.Text = Empty Or cboDesig.Text = Empty Or txtBasicSalary.Text = Empty Then
'Message shown to notify the user that none of the fields can be left empty.
        MsgBox "Select the record to be deleted."
        Exit Sub
    End If
'The selected record is deleted.
    rs.Delete
'The Employee ID list is uodated.
    lstEmployeeID.RemoveItem (lstEmployeeID.ListIndex)
    lstEmployeeID.Refresh
'Message shown to notify the user that the record is deleted from the database.
    MsgBox "The Record has been deleted."
'The fields are cleared after deletion of a record.
    txtEmployeeID.Text = ""
    txtEmployeeName.Text = ""
    txtEmployeeAddress.Text = ""
    txtContactNumber.Text = ""
    cboDesig.Text = ""
    txtBasicSalary.Text = ""
    lstEmployeeID.Clear
    Call Form_Load
End Sub
'Make Entry button click event.
Private Sub cmdMakeEntry_Click()

'The fields are checked to make sure none of them are empty.
    If txtEmployeeID.Text = Empty Or txtEmployeeName.Text = Empty Or txtEmployeeAddress.Text = Empty Or txtContactNumber.Text = Empty Or cboDesig.Text = Empty Or txtBasicSalary.Text = Empty Then
'Message shown to notify the user that none of the fields can be left empty.
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If
    
'Length check to make sure contact number is of 11 digits.
    If Len(txtContactNumber.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If
    
'The data of the new employee are added to the Employee Entry database.
    Set rs = New Recordset
    rs.Open "select * from Employee_entry", con, adOpenDynamic, adLockOptimistic
'Validation check to make sure same Employee is not used twice.
     If Not rs.EOF Then
        Dim IsNewID As Boolean
            IsNewID = True
            Do While Not rs.EOF
                If txtEmployeeID.Text = rs.Fields("EmployeeID") Then
                    IsNewID = False
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
    
    If IsNewID = False Then
        MsgBox "This ID is already in use. Please try again."
        txtEmployeeID.Text = ""
        Exit Sub
    End If
    
    rs.AddNew
        rs.Fields("EmployeeID") = txtEmployeeID.Text
        rs.Fields("EName") = txtEmployeeName.Text
        rs.Fields("Address") = txtEmployeeAddress.Text
        rs.Fields("Contact") = txtContactNumber.Text
        rs.Fields("Designation") = cboDesig.Text
        rs.Fields("Basic_salary") = txtBasicSalary.Text
        rs.Fields("Date_of_entry") = lblDate.Caption
    rs.Update
    
'Message shown to notify the user that the record has been added successfully.
    MsgBox "Record has been added successfully."
'The fields are cleared after the information is added to the database.
    txtEmployeeID.Text = ""
    txtEmployeeName.Text = ""
    txtEmployeeAddress.Text = ""
    txtContactNumber.Text = ""
    cboDesig.Text = ""
    txtBasicSalary.Text = ""
    
    lstEmployeeID.Clear
    lstEmployeeID.Refresh

    Call Form_Load
    
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'The fields are cleared at the click of the button.
    lstEmployeeID.Clear
    lstEmployeeID.Refresh
    txtEmployeeID.Text = ""
    txtEmployeeName.Text = ""
    txtEmployeeAddress.Text = ""
    txtContactNumber.Text = ""
    cboDesig.Text = ""
    txtBasicSalary.Text = ""
    
    Call Form_Load
End Sub

'Update button click event.
Private Sub cmdUpdate_Click()
'Validation check to make sure none of the fields are kept empty.
    If txtEmployeeID.Text = "" Or txtEmployeeName.Text = "" Or txtEmployeeAddress.Text = "" Or txtContactNumber.Text = "" Or cboDesig.Text = "" Or txtBasicSalary.Text = "" Then
        MsgBox "Check if all the required information are filled and then try again."
        Exit Sub
    End If
'Ensuring the Employee ID is not changed as it serves as the basis of identification of the employee.
    If txtEmployeeID.Text <> EID Then
        MsgBox "Employee ID cannot be changed."
        txtEmployeeID.Text = EID
        Exit Sub
    End If
    
'Updating Employee Entry database information for a particular Employee record.
        rs.Fields("EName") = txtEmployeeName.Text
        rs.Fields("Address") = txtEmployeeAddress.Text
        rs.Fields("Contact") = txtContactNumber.Text
        rs.Fields("Designation") = cboDesig.Text
        rs.Fields("Basic_salary") = txtBasicSalary.Text
    rs.Update
'Message shown to notify the user that the update was successful.
    MsgBox "The record has been updated."
    
'The fields on the form are emptied to ready the form for next transaction.
    lstEmployeeID.Clear
    lstEmployeeID.Refresh
    txtEmployeeID.Text = ""
    txtEmployeeName.Text = ""
    txtEmployeeAddress.Text = ""
    txtContactNumber.Text = ""
    cboDesig.Text = ""
    txtBasicSalary.Text = ""
    
    Call Form_Load
    
End Sub

Private Sub Form_Load()
'Calling the function to connect to the database.
    Call dblink
'Updating the date.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
'The Employee IDs present in the Employee database are loaded onto the Employee ID list.
    Set rs = New Recordset
    rs.Open "select * from Employee_entry", con, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        lstEmployeeID.AddItem (rs.Fields("EmployeeID"))
        rs.MoveNext
    Loop
            
    rs.MoveLast
    txtEmployeeID.Text = CInt(rs!EmployeeID) + 1

End Sub
'Employee ID list click event.
Private Sub lstEmployeeID_Click()

    Set rs = New Recordset
'Employee information loaded from database on selection of Employee ID from Employee list.
    rs.Open "select * from Employee_entry where EmployeeID ='" & lstEmployeeID.Text & "'", con, adOpenDynamic, adLockOptimistic
    
        txtEmployeeID.Text = rs.Fields("EmployeeID")
        txtEmployeeName.Text = rs.Fields("EName")
        txtEmployeeAddress.Text = rs.Fields("Address")
        txtContactNumber.Text = rs.Fields("Contact")
        cboDesig.Text = rs.Fields("Designation")
        txtBasicSalary.Text = rs.Fields("Basic_salary")
        
        EID = txtEmployeeID.Text
End Sub


Private Sub txtContactNumber_KeyPress(KeyAscii As Integer)
'Validating that the new vat on picnicspot entered is numeric
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        MsgBox "Please input numbers only"
    Exit Sub
    End If
End Sub
