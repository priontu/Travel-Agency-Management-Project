VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VehicleBookingUpdate 
   Caption         =   "Vehicle Booking Update"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   13.5
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Customer information"
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
      Left            =   240
      TabIndex        =   23
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txtName 
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
         Left            =   1560
         TabIndex        =   44
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtContact 
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
         Left            =   1560
         TabIndex        =   28
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtBookingID 
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
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1320
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpJourneyDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   122355713
         CurrentDate     =   41962
      End
      Begin MSComCtl2.DTPicker dtpDropOffDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   122355713
         CurrentDate     =   41954
      End
      Begin VB.Label lblBookingDate 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label19"
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
         Left            =   1560
         TabIndex        =   38
         Top             =   3840
         Width           =   720
      End
      Begin VB.Label label19 
         AutoSize        =   -1  'True
         Caption         =   "Booking date"
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
         TabIndex        =   37
         Top             =   3840
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
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
         TabIndex        =   34
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Booking ID"
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
         TabIndex        =   31
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Drop off date"
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
         TabIndex        =   30
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Journey date"
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
         TabIndex        =   29
         Top             =   2880
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vehicle information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtStartingKm 
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
         Left            =   1920
         TabIndex        =   43
         Top             =   4680
         Width           =   2895
      End
      Begin VB.ComboBox cboMake 
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
         Left            =   1920
         TabIndex        =   41
         Top             =   840
         Width           =   2895
      End
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
         Left            =   1920
         TabIndex        =   36
         Top             =   1800
         Width           =   2895
      End
      Begin VB.ComboBox cboModel 
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
         Left            =   1920
         TabIndex        =   35
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox cboVehicleType 
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
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtDriverName 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtAdvancePaid 
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
         Left            =   1920
         TabIndex        =   9
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Starting Kilometer"
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
         TabIndex        =   42
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label lblRegNum 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   39
         Top             =   2280
         Width           =   2850
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Driver name"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label14 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label lblNSeats 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   15
         Top             =   2760
         Width           =   2850
      End
      Begin VB.Label Label16 
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
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   1605
      End
      Begin VB.Label lblRatePerKm 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   13
         Top             =   3240
         Width           =   2850
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Advance paid"
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
         TabIndex        =   12
         Top             =   3720
         Width           =   1140
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
      Height          =   2655
      Left            =   10320
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdcancel 
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
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmbUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdrefresh 
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
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Vehicle Booking Update"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   40
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label12"
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
      Left            =   10560
      TabIndex        =   5
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label11 
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
      Left            =   9960
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "VehicleBookingUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaring global variables.
Dim BID As String
Dim OldVID As String
'Make combobox click event.
Private Sub cboMake_Click()
'Clearing the fields that are dependant on Make of the vehicle to prevent errors while booking, if Make of the Vehicle is changed.
    cboModel.Clear
    cboVehicleID.Clear
    lblRegNum.Caption = ""
    cboModel.Text = ""
    cboVehicleID.Text = ""
'Loading the Models available for the selected Make onto the list combobox from the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where Vehicle_type = '" & cboVehicleType.Text & "' and Make ='" & cboMake.Text & "'", con, adOpenDynamic, adLockOptimistic
'Ensuring that every available Model is loaded only once.
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To cboModel.ListCount - 1
                If cboModel.List(i) = rs.Fields("Model") Then
                    IsNewType = False
                    Exit For
                End If
            Next
            If IsNewType = True Then
                cboModel.AddItem (rs.Fields("Model"))
            End If
            rs.MoveNext
        Loop
End Sub
'Model combobox click event.
Private Sub cboModel_Click()
'Clearing the fields that are dependant on the Model of the vehicle to prevent errors in booking information if Model chosen by the user is changed.
    cboVehicleID.Clear
    lblRegNum.Caption = ""
    cboVehicleID.Text = ""
'Loading the Vehicle IDs for the Vehicle types of the selected Make and Model.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where Vehicle_type = '" & cboVehicleType & "' and Make = '" & cboMake.Text & "'  and  Model = '" & cboModel.Text & "'", con, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    Do While Not rs.EOF
            cboVehicleID.AddItem (rs.Fields("VehicleID"))
            rs.MoveNext
    Loop
End Sub
'Vehicle ID comboox click event.
Private Sub cboVehicleID_Click()
'Loading necessary booking information onto the form from the database for the selected Vehicle ID.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where VehicleID ='" & cboVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
        lblRatePerKm.Caption = rs.Fields("Rate_per_km")
        lblNSeats.Caption = rs.Fields("NSeats")
        lblRegNum.Caption = rs.Fields("RegistrationNo")
End Sub
'Vehicle type combobox click event.
Private Sub cboVehicleType_Click()
'Clearing the fields that are depemdamt on the Vehicle type to prevent errors in booking information if the Vehicle type is changed.
    cboMake.Clear
    cboModel.Clear
    cboVehicleID.Clear
    lblRegNum.Caption = ""
    
    cboMake.Text = ""
    cboModel.Text = ""
    cboVehicleID.Text = ""
'The Makes that are available for the selected Vehicle type are loaded onto the Make conbobox from the database.
    Set rs = New Recordset
    rs.Open "select distinct * from Vehicle_entry where Vehicle_type ='" & cboVehicleType.Text & "'", con, adOpenDynamic, adLockOptimistic
'Validation check to ensure that the vehicles that are booked are not loaded.
     If rs.EOF = False Then
        Do While rs.EOF = False
            If rs.Fields("IsBooked?") = False Then
                cboMake.AddItem (rs.Fields("Make"))
            End If
            rs.MoveNext
        Loop
    End If
'Message shown to notify the user that there are no Makes available for the chosen Vehicle type if necessary.
    If cboMake.ListCount = 0 Then
        MsgBox "No cars of that type available"
        cboVehicleType.Text = ""
        Exit Sub
    End If
    
End Sub
'Update button click event.
Private Sub cmbUpdate_Click()
'Validation check to make sure the User provides the Booking ID before update.
    If txtBookingID.Text = Empty Then
        MsgBox "Please enter the Booking ID."
        Exit Sub
    End If
'Presence check to make sure all the necessary information is provided.
    If txtBookingID.Text = "" Or txtName.Text = "" Or txtContact.Text = "" Or txtAddress.Text = "" Or cboVehicleType.Text = "" Or lblRegNum.Caption = "" Or cboMake.Text = "" Or cboModel.Text = "" Or cboVehicleID.Text = "" Or lblRatePerKm.Caption = "" Or lblNSeats.Caption = "" Or txtAdvancePaid = "" Or txtDriverName.Text = "" Then
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If

'Length check to make sure contact number is of 11 digits.
    If Len(txtContact.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If

'Checking if Booking ID is changed.
   If txtBookingID.Text <> BID Then
'Message shown to notify the User that Booking ID cannot be changed.
        MsgBox "Booking ID cannot be changed."
        txtBookingID = BID
        Exit Sub
    End If
'Checking if Vehicle ID is changed, which means that the customer changed his choice of vehicle.
    If cboVehicleID.Text <> OldVID Then
        Set rs = New Recordset
'Updating availability information if Vehicle ID is changed.
        rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
            Do While Not rs.EOF
                If rs.Fields("VehicleID") = OldVID Then
                    rs.Fields("IsBooked?") = False
                End If
                
                If rs.Fields("VehicleID") = cboVehicleID.Text Then
                    rs.Fields("IsBooked?") = True
                End If
                
                rs.Update
                rs.MoveNext
            Loop
        rs.Close
    End If
          
'Updating database information for the selected Booking ID
    Set rs = New Recordset
    rs.Open "select * from Vehicle_booking where VehicleBookingID = '" & txtBookingID.Text & "'", con, adOpenDynamic, adLockOptimistic
   
        rs.Fields("Customer_name") = txtName.Text
        rs.Fields("Customer_address") = txtAddress.Text
        rs.Fields("Contact") = txtContact.Text
        rs.Fields("Journey_date") = dtpJourneyDate.Value
        rs.Fields("Drop_off_date") = dtpDropOffDate.Value
        rs.Fields("Vehicle_type") = cboVehicleType.Text
        rs.Fields("Registration_number") = lblRegNum.Caption
        rs.Fields("Make") = cboMake.Text
        rs.Fields("Model") = cboModel.Text
        rs.Fields("VehicleID") = cboVehicleID.Text
        rs.Fields("Rate_per_kilometer") = lblRatePerKm.Caption
        rs.Fields("Number_of_seats") = lblNSeats.Caption
        rs.Fields("Advance_paid") = txtAdvancePaid.Text
        rs.Fields("Driver_name") = txtDriverName.Text
        rs.Fields("Date_of_booking") = lblDate.Caption
        rs.Fields("Starting_Km") = txtStartingKm.Text
    rs.Update
'Message shown to notify the User that the update was successful.
    MsgBox "Booking has been updated successfully."
 'All the fields are cleared after the information is updated to ready form for next job.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    cboVehicleType.Text = ""
    lblRegNum.Caption = ""
    cboMake.Clear
    cboModel.Clear
    cboVehicleID.Clear
    lblRatePerKm.Caption = ""
    lblNSeats.Caption = ""
    txtAdvancePaid = ""
    txtStartingKm = ""
    txtDriverName.Text = ""
    
    Call Form_Load
    
End Sub
'Cancel button click event.
Private Sub cmdCancel_Click()
'Closing the form.
    Unload Me
End Sub
'Delete button click event.
Private Sub cmdDelete_Click()
'Presence check to make sure all the required information is provided.
    If txtBookingID.Text = "" Or txtName.Text = "" Or txtContact.Text = "" And txtAddress.Text = "" Or cboVehicleType.Text = "" Or lblRegNum.Caption = "" Or cboMake.Text = "" Or cboModel.Text = "" Or cboVehicleID.Text = "" Or lblRatePerKm.Caption = "" Or lblNSeats.Caption = "" Or txtAdvancePaid = "" Or txtDriverName.Text = "" Then
'Message shown to notify the User that all the information needs to be provided.
        MsgBox "None of the fields can be left empty."
        Exit Sub
    End If
'Deletion of the selected record from the database.
    rs.Delete
    MsgBox "Booking has been deleted."
    rs.Close
'Updating the availability information for the vehicle in the Vehicle Entry database.
    rs.Open " select * from Vehicle_entry where VehicleID ='" & cboVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
        rs.Fields("IsBooked?") = False
    rs.Update
'Clearing the unnecesary information on the form after deletion of the record.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    cboVehicleType.Text = ""
    lblRegNum.Caption = ""
    cboMake.Clear
    cboModel.Clear
    cboVehicleID.Clear
    lblRatePerKm.Caption = ""
    lblNSeats.Caption = ""
    txtAdvancePaid = ""
    txtStartingKm = ""
    txtDriverName.Text = ""
    rs.Close
    
    Call Form_Load
    
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'Clearing all the fields on the form as required by the User in case the User wants to start anew.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    cboVehicleType.Text = ""
    lblRegNum.Caption = ""
    cboMake.Clear
    cboModel.Clear
    cboVehicleID.Clear
    lblRatePerKm.Caption = ""
    lblNSeats.Caption = ""
    txtAdvancePaid = ""
    txtStartingKm.Text = ""
    txtDriverName.Text = ""
    
    Call Form_Load
    
End Sub
'Search button click event.
Private Sub cmdSearch_Click()
'Presence check to ensure the User provides the Booking ID to be searched.
    If txtBookingID.Text = Empty Then
'Message shown to ask for the Vehicle ID to be searched from the User.
        MsgBox "Please enter the Booking ID."
        Exit Sub
    End If
    
'The Vehicle booking database is searched for a paticular Booking ID.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_booking", con, adOpenDynamic, adLockOptimistic
    Dim found As Boolean
    found = False
    rs.MoveFirst
    Do While Not rs.EOF And Not found
        If rs.Fields("VehicleBookingID") = txtBookingID.Text Then
        found = True
        End If
        rs.MoveNext
    Loop
    
    If found = False Then
'Message shown to notify User if Booking ID not found.
        MsgBox "Booking ID being searched is not found. Please enter correct Booking ID and try again."
        Exit Sub
    End If
    rs.Close
'The Booking information for the particular Booking ID provided is loaded onto the form from the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_booking where VehicleBookingID = '" & txtBookingID.Text & "'", con, adOpenDynamic, adLockOptimistic
   
        txtName.Text = rs.Fields("Customer_name")
        txtContact.Text = rs.Fields("Contact")
        txtAddress.Text = rs.Fields("Customer_address")
        dtpJourneyDate.Value = rs.Fields("Journey_date")
        dtpDropOffDate.Value = rs.Fields("Drop_off_date")
        lblBookingDate.Caption = rs.Fields("Date_of_booking")
        cboVehicleType.Text = rs.Fields("Vehicle_type")
        cboMake.Text = rs.Fields("Make")
        cboModel.Text = rs.Fields("Model")
        cboVehicleID.Text = rs.Fields("VehicleID")
        lblRegNum.Caption = rs.Fields("Registration_number")
        lblNSeats.Caption = rs.Fields("Number_of_seats")
        lblRatePerKm.Caption = rs.Fields("Rate_per_kilometer")
        txtAdvancePaid.Text = rs.Fields("Advance_paid")
        txtDriverName.Text = rs.Fields("Driver_name")
        txtStartingKm.Text = rs.Fields("Starting_Km")
        
    rs.Update
    
    OldVID = cboVehicleID.Text
    BID = txtBookingID.Text
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Showing the current date on the form.
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
'Loading Vehicle type information onto the Vehicle type combobox from the database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry", con, adOpenDynamic, adLockOptimistic
'Check made to make sure each type is loaded only once from the database.
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
End Sub
'Advance paid keypress event.
Private Sub txtAdvancePaid_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub

Private Sub txtStartingKm_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
