VERSION 5.00
Begin VB.Form FuelUpdate 
   Caption         =   "Fuel record"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form10"
   ScaleHeight     =   4875
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   4335
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbfueltype 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   3000
   End
   Begin VB.TextBox txtdate 
      Height          =   400
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   3000
   End
   Begin VB.TextBox txtquatity 
      Height          =   400
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
      Width           =   3000
   End
   Begin VB.TextBox txtcost 
      Height          =   400
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   3000
   End
   Begin VB.TextBox txtcarid 
      Height          =   400
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "Date"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Fuel type"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cost"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Car ID"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FuelUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call dblink
End Sub
