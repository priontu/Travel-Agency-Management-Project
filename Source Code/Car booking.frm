VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VehicleBooking 
   Caption         =   "Vehicle Booking"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form4"
   ScaleHeight     =   9450
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rpt1 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   4935
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   9135
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
         Left            =   1680
         TabIndex        =   41
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtRegNum 
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
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
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
         Left            =   5640
         TabIndex        =   33
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ListBox lstModel 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "Car booking.frx":0000
         Left            =   3000
         List            =   "Car booking.frx":0002
         TabIndex        =   25
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ListBox lstMake 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "Car booking.frx":0004
         Left            =   120
         List            =   "Car booking.frx":0006
         TabIndex        =   24
         Top             =   1200
         Width           =   2775
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
         Left            =   5640
         TabIndex        =   18
         Top             =   3840
         Width           =   2535
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
         ItemData        =   "Car booking.frx":0008
         Left            =   120
         List            =   "Car booking.frx":000A
         TabIndex        =   17
         Top             =   600
         Width           =   2895
      End
      Begin VB.ListBox lstVehicleID 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "Car booking.frx":000C
         Left            =   5880
         List            =   "Car booking.frx":000E
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Starting kilometer"
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
         TabIndex        =   40
         Top             =   4440
         Width           =   1425
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
         Left            =   5880
         TabIndex        =   34
         Top             =   360
         Width           =   1635
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
         Left            =   4560
         TabIndex        =   32
         Top             =   3480
         Width           =   1020
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
         Left            =   1680
         TabIndex        =   29
         Top             =   3480
         Width           =   2610
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
         TabIndex        =   28
         Top             =   3480
         Width           =   1485
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
         Left            =   1680
         TabIndex        =   27
         Top             =   3960
         Width           =   2610
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
         TabIndex        =   26
         Top             =   3960
         Width           =   1305
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
         Left            =   4560
         TabIndex        =   23
         Top             =   3840
         Width           =   1215
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
         Left            =   5880
         TabIndex        =   22
         Top             =   960
         Width           =   975
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
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
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
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
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
         Left            =   3000
         TabIndex        =   19
         Top             =   960
         Width           =   690
      End
   End
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
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   9135
      Begin MSComCtl2.DTPicker dtpJourneyDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Top             =   1920
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
         Format          =   164757505
         CurrentDate     =   41962
      End
      Begin MSComCtl2.DTPicker dtpDropOffDate 
         Height          =   315
         Left            =   5880
         TabIndex        =   14
         Top             =   1920
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
         Format          =   164757505
         CurrentDate     =   41954
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
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   360
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
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
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
         Left            =   1440
         TabIndex        =   6
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
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
         TabIndex        =   30
         Top             =   1920
         Width           =   1005
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
         Left            =   4560
         TabIndex        =   13
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
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
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   615
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
         TabIndex        =   9
         Top             =   840
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   8520
      Width           =   5535
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
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton cmbConfirm 
         Caption         =   "Confirm"
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
         ToolTipText     =   "Click the button in order to confirm the booking."
         Top             =   240
         Width           =   1700
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
         Height          =   360
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1700
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
         Left            =   960
         TabIndex        =   36
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Vehicle Booking"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   39
      Top             =   120
      Width           =   3735
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
      Left            =   7800
      TabIndex        =   38
      Top             =   360
      Width           =   1035
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
      Left            =   7200
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "VehicleBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Vehicle type combobox click event.
Private Sub cboVehicleType_Click()
'The fields that depend on the type of the vehicle are cleared when the combobox information are changed to prevent database information storage override.
    lstMake.Clear
    lstModel.Clear
    lstVehicleID.Clear
    txtRegNum.Text = ""
'The Makes for the selected Vehicle type are loaded onto the Make list.
'Each particular Make name is selected only once.
    Set rs = New Recordset
    rs.Open "select distinct * from Vehicle_entry where Vehicle_type='" & cboVehicleType.Text & "'", con, adOpenDynamic, adLockOptimistic

'Each Make name is loaded onto the list only once.
'Validation check  ensure only the vehicles that are not booked are loaded onto the list.
  If rs.EOF = False Then
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To lstMake.ListCount - 1
                If lstMake.List(i) = rs.Fields("Make") Then
                    IsNewType = False
                End If
            Next
            If IsNewType = True Then
                 If rs.Fields("IsBooked?") = False Then
                    lstMake.AddItem (rs.Fields("Make"))
                 End If
                
            End If
            rs.MoveNext
        Loop
    End If
    
End Sub
'Comfirm button click event.
Private Sub cmbConfirm_Click()
'Presence check to make sure that none of the fields are left empty.
    If txtBookingID.Text = "" Or txtName.Text = "" Or txtContact.Text = "" Or txtAddress.Text = "" Or cboVehicleType.Text = "" Or txtRegNum.Text = "" Or txtAdvancePaid.Text = "" Or txtDriverName.Text = "" Or txtStartingKm.Text = "" Then
        MsgBox "Some of the required fields are not filled. Please fill up all of the fields and try again."
        Exit Sub
    End If
    
'Length check to make sure contact number is of 11 digits.
    If Len(txtContact.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If

    Set rs = New Recordset
    
    rs.Open "select * from Vehicle_booking", con, adOpenDynamic, adLockOptimistic
'Validation check to make sure the same Booking ID is not used twice.
     If Not rs.EOF Then
        Dim IsNewID As Boolean
            IsNewID = True
            Do While Not rs.EOF
                If txtBookingID.Text = rs.Fields("VehicleBookingID") Then
                    IsNewID = False
                    Exit Do
                End If
            rs.MoveNext
        Loop
    End If
'Message shown to notify the user that the Booking ID is in use if the same User ID is found in database.
    If IsNewID = False Then
        MsgBox "This Booking ID is already in use. Please try again."
        txtBookingID.Text = ""
        Exit Sub
    End If
'Storage of Vehicle Booking information into the database.
    rs.AddNew
        rs.Fields("VehicleBookingID") = txtBookingID.Text
        rs.Fields("Customer_name") = txtName.Text
        rs.Fields("Customer_address") = txtAddress.Text
        rs.Fields("Contact") = txtContact.Text
        rs.Fields("Journey_date") = dtpJourneyDate.Value
        rs.Fields("Drop_off_date") = dtpDropOffDate.Value
        rs.Fields("Vehicle_type") = cboVehicleType.Text
        rs.Fields("Registration_number") = txtRegNum.Text
        rs.Fields("Make") = lstMake.Text
        rs.Fields("Model") = lstModel.Text
        rs.Fields("VehicleID") = lstVehicleID.Text
        rs.Fields("Rate_per_kilometer") = lblRatePerKm.Caption
        rs.Fields("Number_of_seats") = lblNSeats.Caption
        rs.Fields("Advance_paid") = txtAdvancePaid.Text
        rs.Fields("Driver_name") = txtDriverName.Text
        rs.Fields("Date_of_booking") = lblDate.Caption
        rs.Fields("Starting_Km") = txtStartingKm.Text
    rs.Update
    
'Message shown notify the User that Booking was successful.
    MsgBox "Booking has been made successfully."
    rs.Close
'The Vehicle is made unavailable to further bookings.
    rs.Open "select * from Vehicle_entry where VehicleID ='" & lstVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
        rs.Fields("IsBooked?") = True
    rs.Update
'Generate token.
rpt1.ReportFileName = App.Path & "\Reports\Token.rpt"
rpt1.SelectionFormula = "{Vehicle_Booking.BookingID}='" & txtBookingID.Text & "'"
rpt1.Action = 2

'The fields are all cleared and the form is refreshed after the booking is made.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    cboVehicleType.Text = ""
    txtRegNum.Text = ""
    lstMake.Clear
    lstModel.Clear
    lstVehicleID.Clear
    lblRatePerKm.Caption = ""
    lblNSeats.Caption = ""
    txtAdvancePaid = ""
    txtDriverName.Text = ""
    txtStartingKm.Text = ""
    
    Call Form_Load
End Sub
'Cancel button click event
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Refresh button click event.
Private Sub cmdRefresh_Click()
'The fields are cleared at the click of the button as the User might require in case he/she needs to start anew.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtContact.Text = ""
    txtAddress.Text = ""
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    cboVehicleType.Text = ""
    txtRegNum.Text = ""
    lstMake.Clear
    lstModel.Clear
    lstVehicleID.Clear
    lblRatePerKm.Caption = ""
    lblNSeats.Caption = ""
    txtAdvancePaid = ""
    txtDriverName.Text = ""
    txtStartingKm.Text = ""
    
    Call Form_Load
    
End Sub
'Form load event.
Private Sub Form_Load()
'Calling function to connect to database.
    Call dblink
'Shows current date.
    lblDate.Caption = Format(Now, "dd / mm / yyyy")
'The date pickers are assigned current dates as default.
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
'Loading the Vehicle type information onto the Vehicle type combobox from Vehicle entry database.
  
    Set rs = New Recordset
    rs.Open "select * from Vehicle_booking", con, adOpenDynamic, adLockOptimistic
        rs.MoveLast
        txtBookingID.Text = CInt(rs!VehicleBookingID) + 1
    rs.Close
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
'Validity check to make sure any booked vehicles are not loaded.
                 If rs.Fields("IsBooked?") = False Then
                    cboVehicleType.AddItem (rs.Fields("Vehicle_type"))
                 End If
                
            End If
            rs.MoveNext
        Loop
    End If
    


End Sub
'Make list click event.
Private Sub lstMake_Click()
'The information dependent on the Make of the vehicle are cleared to make sure no override occurs when storing data into the database.
    lstModel.Clear
    lstVehicleID.Clear
    txtRegNum.Text = ""
'The particular Models are loaded for the selected Make.
    Set rs = New Recordset
    rs.Open "select distinct * from Vehicle_entry where Vehicle_type = '" & cboVehicleType.Text & "' and Make ='" & lstMake.Text & "'", con, adOpenDynamic, adLockOptimistic
'Each Model name is loaded onto the list only once.
'Validation check to make sure only the vehicles that are not booked are loaded onto the list.
  If rs.EOF = False Then
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To lstModel.ListCount - 1
                If lstModel.List(i) = rs.Fields("Model") Then
                    IsNewType = False
                End If
            Next
            If IsNewType = True Then
                 If rs.Fields("IsBooked?") = False Then
                    lstModel.AddItem (rs.Fields("Model"))
                 End If
                
            End If
            rs.MoveNext
        Loop
    End If

End Sub
'Model list click event.
Private Sub lstModel_Click()
'The information that are dependant on the Make and Model are refreshed if Model is changed, to prevent database storage override.
    lstVehicleID.Clear
    txtRegNum.Text = ""
'The Vehicle IDs of the Model of Vehicles of the particular Make is loaded.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where Vehicle_type = '" & cboVehicleType & "' and Make = '" & lstMake.Text & "'  and  Model = '" & lstModel.Text & "'", con, adOpenDynamic, adLockOptimistic
    rs.MoveFirst

'Each Make name is loaded onto the list only once.
'Validation check  ensure only the vehicles that are not booked are loaded onto the list.
  If rs.EOF = False Then
        Do While rs.EOF = False
            Dim i As Integer
            Dim IsNewType As Boolean
            IsNewType = True
            For i = 0 To lstVehicleID.ListCount - 1
                If lstVehicleID.List(i) = rs.Fields("VehicleID") Then
                    IsNewType = False
                End If
            Next
            If IsNewType = True Then
                 If rs.Fields("IsBooked?") = False Then
                    lstVehicleID.AddItem (rs.Fields("VehicleID"))
                 End If
                
            End If
            rs.MoveNext
        Loop
    End If
End Sub
'Vehicle ID list click event.
Private Sub lstVehicleID_Click()
'Further required information for the selected Vehicle is loaded.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_entry where VehicleID ='" & lstVehicleID.Text & "'", con, adOpenDynamic, adLockOptimistic
        lblRatePerKm.Caption = rs.Fields("Rate_per_km")
        lblNSeats.Caption = rs.Fields("NSeats")
        txtRegNum.Text = rs.Fields("RegistrationNo")
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
'Contact number keypress event.
Private Sub txtContact_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
'Starting kilometer keypress event.
Private Sub txtStartingKm_KeyPress(KeyAscii As Integer)
'Character type check to make sure only numerical data is entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub
