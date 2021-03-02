VERSION 5.00
Begin VB.Form frmBackup 
   Caption         =   "Backup"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.DirListBox dirTargetFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   4095
   End
   Begin VB.DirListBox dirDestFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup Files"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Pressing the button will initiate backup process."
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose Directory for Backup File:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   2910
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Choose  Directory for Target File:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBackup_Click()
    Dim SourceFile, DestFile
    Dim SourceFileName, DestFileName
    'The user has to enter the name of the files
    SourceFileName = InputBox("Enter name of target file: ")
    DestFileName = InputBox("Enter name of backup file: ")
    On Error GoTo FileCopyFailed
    'The backup of the file is then made in the chosen directory
    SourceFile = dirTargetFile.Path + "\" + SourceFileName + ".accdb"
    DestFile = dirDestFile.Path + "\" + DestFileName + ".accdb"
    FileCopy SourceFile, DestFile
    MsgBox "Backup has been made"
    Call dblink
    Exit Sub
'If an error occurs, a description of that error is given to the user
FileCopyFailed:
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    'Exit the form
    Unload Me
End Sub

Private Sub Form_Load()
    'This is required as the backup cannot be made while the database file
    'is being used by Visual Basic 6
    con.Close
End Sub
