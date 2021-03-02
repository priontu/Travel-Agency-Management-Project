VERSION 5.00
Begin VB.Form FuelRecord 
   Caption         =   "Fuel record"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form10"
   ScaleHeight     =   4185
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboVehicleID 
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
      ItemData        =   "Fuel Record.frx":0000
      Left            =   1560
      List            =   "Fuel Record.frx":0002
      TabIndex        =   11
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1680
      TabIndex        =   8
      Top             =   3000
      Width           =   6015
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
         TabIndex        =   13
         Top             =   240
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
         Height          =   495
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1815
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
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox cboFuelType 
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
      ItemData        =   "Fuel Record.frx":0004
      Left            =   1560
      List            =   "Fuel Record.frx":000E
      TabIndex        =   7
      Top             =   2280
      Width           =   3000
   End
   Begin VB.TextBox txtQuantity 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6240
      TabIndex        =   6
      Top             =   1440
      Width           =   1560
   End
   Begin VB.TextBox txtCost 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6240
      TabIndex        =   5
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "litres"
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
      Left            =   7920
      TabIndex        =   15
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fuel Record"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   14
      Top             =   360
      Width           =   2160
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   12
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label5 
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
      Left            =   6240
      TabIndex        =   4
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Fuel type"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cost"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "VehicleID"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "FuelRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'All the fields on the form are cleared.
    cboVehicleID.Text = ""
    cboFuelType.Text = ""
    txtQuantity.Text = ""
    txtCost.Text = ""
    
End Sub
'Save button click event.
Private Sub cmdSave_Click()
'Validation check to make sure that none of the important fields are kept empty.
    If cboVehicleID.Text = Empty Or cboFuelType.Text = Empty Or txtQuantity.Text = Empty Or txtCost.Text = Empty Then
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If
'Storing information into the database.
    Set rs = New Recordset
    rs.Open "select * from Fuel_record", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("VehicleID") = cboVehicleID.Text
        rs.Fields("Date_of_entry") = lblDate.Caption
        rs.Fields("FuelType") = cboFuelType.Text
        rs.Fields("FuelQuantity/litres") = txtQuantity.Text
        rs.Fields("FuelCost") = txtCost.Text
    rs.Update
    MsgBox "Record has been added successfully"
'The fields are cleared after the information are stored into the database.
    cboVehicleID.Text = ""
    cboFuelType.Text = ""
    txtQuantity.Text = ""
    txtCost.Text = ""
    
    
    End Sub
'Form load click event.
Private Sub Form_Load()
'Calling function to connect to the database.
    Call dblink
'Shows current date.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
'The vehicle types available on the Vehicle Entry database are loaded.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        cboVehicleID.AddItem (rs.Fields("VehicleID"))
        rs.MoveNext
    Loop
End Sub
'Cost of fuel keypress event.
Private Sub txtCost_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
'Quantity of fuel keypress event.
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
        If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only"
        Exit Sub
    End If
End Sub
