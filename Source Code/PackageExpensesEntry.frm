VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PackageExpensesEntry 
   Caption         =   "Package expenses entry"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Overall expenses data"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   4815
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         Begin VB.CommandButton cmdCalculate 
            Caption         =   "Calculate"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
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
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1215
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
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   1215
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
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Width           =   1215
         End
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   2640
         Width           =   2250
      End
      Begin MSFlexGridLib.MSFlexGrid flxgrdExpensesDetails 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total "
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
         TabIndex        =   11
         Top             =   2640
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Package expenses information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
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
         Left            =   2160
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cboExpenses 
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
         ItemData        =   "PackageExpensesEntry.frx":0000
         Left            =   1320
         List            =   "PackageExpensesEntry.frx":000D
         TabIndex        =   8
         Top             =   1320
         Width           =   2490
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
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2490
      End
      Begin VB.TextBox txtPackageName 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   2490
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   5
         Top             =   1800
         Width           =   2490
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Expenses"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Package Expenses Entry"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   120
      Width           =   3375
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
      Left            =   4200
      TabIndex        =   20
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
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "PackageExpensesEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaration of a global variable.
Dim r As Integer
'Calcualte button click event.
Private Sub cmdCalculate_Click()
'Calculation of total expenses
    Dim total As Double
    total = 0
    For i = 1 To r
        total = total + Val(flxgrdExpensesDetails.TextMatrix(i, 2))
    Next i
    
    txtTotal.Text = total
End Sub
'Done button click event.
Private Sub cmdDone_Click()
'Presence check to make sure all the information is provided.
    If txtPackageID.Text = "" Or txtPackageName.Text = "" Or cboExpenses.Text = "" Or txtAmount.Text = "" Then
        MsgBox "None of the fields in the Package Expenses information section can be left empty. "
        Exit Sub
    End If
    
'Adding expenses information into the flexgrid.
        r = r + 1
    flxgrdExpensesDetails.TextMatrix(r, 0) = txtPackageID.Text
    flxgrdExpensesDetails.TextMatrix(r, 1) = cboExpenses.Text
    flxgrdExpensesDetails.TextMatrix(r, 2) = txtAmount.Text
    
    flxgrdExpensesDetails.Rows = flxgrdExpensesDetails.Rows + 1
'Adding expenses information into the Package expenses details database table.
    Set rs = New Recordset
    rs.Open "select * from Package_expenses_details", con, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("PackageID") = txtPackageID.Text
            rs.Fields("Expenses") = cboExpenses.Text
            rs.Fields("Amount") = Val(txtAmount.Text)
        rs.Update
    rs.MoveNext
'These fields are cleared in case they need to be changed.
    cboExpenses.Text = ""
    txtAmount.Text = ""
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'All the fields are emptied and refreshed in case the user wants to start anew.
    txtPackageID.Text = ""
    txtPackageName.Text = ""
    cboExpenses.Text = ""
    flxgrdExpensesDetails.Clear
    flxgrdExpensesDetails.Rows = 2
    Call Form_Load
    txtTotal.Text = ""
End Sub
'Save button click event.
Private Sub cmdSave_Click()
'Storing final cost and expenses information in the Package expenses master database.
    Set rs = New Recordset
    rs.Open "select * from Package_expenses_master", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("PackageID") = txtPackageID.Text
        rs.Fields("Package_name") = txtPackageName.Text
        rs.Fields("Total") = txtTotal.Text
    rs.Update
    MsgBox "The record has been saved."
End Sub
'Form load event.
Private Sub Form_Load()
'Calling the function to connect the form to the database.
    Call dblink
'Showing the current date on the form.
    lblDate.Caption = Date
'Assigning fields to the flexgrid columns.
    flxgrdExpensesDetails.TextMatrix(0, 0) = "Package ID"
    flxgrdExpensesDetails.TextMatrix(0, 1) = "Expenses"
    flxgrdExpensesDetails.TextMatrix(0, 2) = "Amount"
    r = 0
   
End Sub

'Amount keypress event.
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

'Amount textbox keyup event.
Private Sub txtAmount_KeyUp(KeyCode As Integer, Shift As Integer)
'Directing the cursor to the Done button after information is inserted in the amount textbox.
    If KeyCode = vbKeyReturn Then
        cmdDone.SetFocus
    End If
End Sub
'Total keypress event.
Private Sub txtTotal_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
