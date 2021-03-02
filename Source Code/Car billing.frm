VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VehicleBilling 
   Caption         =   "Vehicle Billing"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15765
   LinkTopic       =   "Form8"
   ScaleHeight     =   8550
   ScaleWidth      =   15765
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rpt1 
      Left            =   13680
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   73
      Top             =   1560
      Width           =   2175
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
      Height          =   375
      Left            =   10680
      TabIndex        =   64
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmbSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   61
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame8 
      Caption         =   "Billing Information"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   5760
      TabIndex        =   25
      Top             =   2160
      Width           =   9975
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   75
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   63
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmdGenerateBill 
         Caption         =   "Generate bill"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   62
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   4800
         TabIndex        =   44
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtLateFine2 
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
            Left            =   2160
            TabIndex        =   71
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtCostOfTravel2 
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
            Left            =   2160
            TabIndex        =   70
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtBookingCharge 
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
            Left            =   2160
            TabIndex        =   53
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtAdditionalExp 
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
            Left            =   2160
            TabIndex        =   52
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtDriverCharge 
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
            Left            =   2160
            TabIndex        =   51
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtGross 
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
            Left            =   2160
            TabIndex        =   50
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtNet 
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
            Left            =   2280
            TabIndex        =   49
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Frame Frame7 
            Caption         =   "(Less)"
            BeginProperty Font 
               Name            =   "Britannic Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   360
            TabIndex        =   45
            Top             =   3120
            Width           =   4575
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
               Height          =   330
               Left            =   2040
               TabIndex        =   67
               Text            =   " "
               Top             =   720
               Width           =   2415
            End
            Begin VB.TextBox txtDiscount 
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
               Left            =   2040
               TabIndex        =   46
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label25 
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
               Height          =   330
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Discount(If any)"
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
               TabIndex        =   47
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Gross total"
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
            TabIndex        =   60
            Top             =   2760
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Driver charge"
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
            TabIndex        =   59
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Additional Expenditures"
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
            TabIndex        =   58
            Top             =   2280
            Width           =   1965
         End
         Begin VB.Label Label10 
            Caption         =   "Booking Charge"
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
            TabIndex        =   57
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Net total"
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
            TabIndex        =   56
            Top             =   4440
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Late fine(If any)"
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
            TabIndex        =   55
            Top             =   1800
            Width           =   1380
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cost of travel"
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
            TabIndex        =   54
            Top             =   1320
            Width           =   1170
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dates"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   4575
         Begin VB.TextBox txtLateFine 
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
            Left            =   1560
            TabIndex        =   72
            Top             =   1920
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpJourneyDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   35
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   132186113
            CurrentDate     =   41962
         End
         Begin MSComCtl2.DTPicker dtpDropOffDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   132186113
            CurrentDate     =   41954
         End
         Begin MSComCtl2.DTPicker dtpBookingDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   37
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   132186113
            CurrentDate     =   41954
         End
         Begin MSComCtl2.DTPicker dtpBillingDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   38
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   132186113
            CurrentDate     =   41954
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Billing date"
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
            TabIndex        =   43
            Top             =   1440
            Width           =   1005
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
            TabIndex        =   42
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label16 
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
            TabIndex        =   41
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Booking Date"
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
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Late fine(If any)"
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
            TabIndex        =   39
            Top             =   2040
            Width           =   1260
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cost of travel"
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   4575
         Begin VB.TextBox txtCostOfTravel 
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
            Left            =   2040
            TabIndex        =   69
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtKmTravelled 
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
            Left            =   2040
            TabIndex        =   68
            Top             =   1320
            Width           =   2415
         End
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
            Height          =   330
            Left            =   2040
            TabIndex        =   28
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtEndingKm 
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
            Left            =   2040
            TabIndex        =   27
            Top             =   840
            Width           =   2415
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
            Height          =   285
            Left            =   2040
            TabIndex        =   74
            Top             =   1800
            Width           =   2370
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kilometer travelled"
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
            Width           =   1485
         End
         Begin VB.Label Label23 
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
            TabIndex        =   32
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ending kilometer"
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
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label Label26 
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
            TabIndex        =   30
            Top             =   1800
            Width           =   1485
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cost of travel"
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
            Top             =   2280
            Width           =   1170
         End
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
      Height          =   3735
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   5175
      Begin VB.TextBox txtVehicleType 
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
         Left            =   2040
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtVehicleID 
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
         Left            =   2040
         TabIndex        =   23
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtModel 
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
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtMake 
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
         Left            =   2040
         TabIndex        =   21
         Top             =   840
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
         Height          =   330
         Left            =   2040
         TabIndex        =   12
         Top             =   2760
         Width           =   2895
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label32 
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
         TabIndex        =   20
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label31 
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
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label30 
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
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label29 
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
         TabIndex        =   17
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label28 
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
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label27 
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
         TabIndex        =   15
         Top             =   3240
         Width           =   1305
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
         Left            =   2040
         TabIndex        =   14
         Top             =   3240
         Width           =   2850
      End
      Begin VB.Label Label24 
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1635
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
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
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
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
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
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   3015
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
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3015
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
         Height          =   1005
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label21 
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
         TabIndex        =   8
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label20 
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
         TabIndex        =   7
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label19 
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
         TabIndex        =   6
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label17 
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
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblDate 
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
      Left            =   12600
      TabIndex        =   66
      Top             =   240
      Width           =   1650
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
      Left            =   11880
      TabIndex        =   65
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label22 
      Caption         =   "Vehicle Billing"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   9
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "VehicleBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Search button click event.
Private Sub cmbSearch_Click()
'Validation check to make sure Booking ID is provided.
    If txtBookingID.Text = Empty Then
        MsgBox "Please enter the Booking ID and try again."
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
        MsgBox "There is no record for this Booking ID."
        Exit Sub
    End If
   rs.Close
'The Billing information is loaded from the Vehicle Booking database.
    Set rs = New Recordset
    rs.Open "select * from Vehicle_booking where VehicleBookingID='" & txtBookingID.Text & "'", con, adOpenDynamic, adLockOptimistic
        txtName.Text = rs.Fields("Customer_name")
        txtContact.Text = rs.Fields("Contact")
        txtAddress.Text = rs.Fields("Customer_address")
        txtVehicleType.Text = rs.Fields("Vehicle_type")
        txtMake.Text = rs.Fields("Make")
        txtModel.Text = rs.Fields("Model")
        txtRegNum.Text = rs.Fields("Registration_number")
        txtVehicleID.Text = rs.Fields("VehicleID")
        txtDriverName.Text = rs.Fields("Driver_name")
        lblNSeats.Caption = rs.Fields("Number_of_seats")
        txtStartingKm.Text = rs.Fields("Starting_km")
        lblRatePerKm.Caption = rs.Fields("Rate_per_kilometer")
        dtpBookingDate.Value = rs.Fields("Date_of_booking")
        dtpJourneyDate.Value = rs.Fields("Journey_date")
        dtpDropOffDate.Value = rs.Fields("Drop_off_date")
        txtAdvancePaid.Text = rs.Fields("Advance_paid")
        
        txtBookingID.Locked = True
        
'The difference between the dates is calculated.
    Dim days As Integer
    days = DateDiff("d", dtpDropOffDate.Value, dtpBillingDate.Value)
    If dtpBillingDate.Value > dtpDropOffDate.Value Then
'Calculation of late fine.
        txtLateFine.Text = days * 50
    Else
        txtLateFine.Text = 0
    End If
    
    txtLateFine2.Text = txtLateFine.Text
    txtBookingCharge.Text = 2000
End Sub

'Cancel button click event.
Private Sub cmdCancel_Click()
'Unloading the form.
    Unload Me
End Sub
'Generate Bill button click event.
Private Sub cmdGenerateBill_Click()
'Validation check to make sure none of the fields are left empty.
    If txtBookingID.Text = "" Or txtName.Text = "" Or txtContact.Text = "" Or txtAddress.Text = Empty Or txtVehicleType.Text = "" Or txtVehicleID.Text = "" Or txtMake.Text = "" Or txtModel.Text = "" Or txtRegNum.Text = "" Or txtDriverName.Text = "" Or lblNSeats.Caption = "" Or txtStartingKm.Text = "" Or txtEndingKm.Text = "" Or txtKmTravelled.Text = "" Or txtCostOfTravel.Text = "" Or txtLateFine.Text = "" Or txtBookingCharge.Text = "" Or txtDriverCharge.Text = "" Or txtAdditionalExp.Text = "" Or txtDiscount.Text = "" Or txtAdvancePaid.Text = "" Then
        MsgBox "Please make sure none of the fields are left empty and try again. "
        Exit Sub
    End If
'Length check to make sure contact number is of 11 digits.
    If Len(txtContact.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If

    If txtDiscount.Text = "" Then
        txtDiscount.Text = 0
    End If
    
    txtGross.Text = Val(txtBookingCharge.Text) + Val(txtDriverCharge.Text) + Val(txtCostOfTravel2.Text) + Val(txtLateFine2.Text) + Val(txtAdditionalExp.Text)
    txtNet.Text = Val(txtGross.Text) - Val(txtDiscount.Text) - Val(txtAdvancePaid.Text)
    
End Sub

Private Sub cmdPrint_Click()
rpt1.ReportFileName = App.Path & "\Reports\VehicleRentalBill.rpt"
rpt1.SelectionFormula = "{Vehicle_Billing.BookingID}='" & txtBookingID.Text & "'"
rpt1.Action = 2
End Sub

'Refresh button click event.
Private Sub cmdRefresh_Click()
'All the fields on the form are cleared on the click of the button as required by the User.
    txtBookingID.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtContact.Text = ""
    txtVehicleType.Text = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtVehicleID = ""
    txtRegNum.Text = ""
    txtDriverName.Text = ""
    lblNSeats.Caption = ""
    txtStartingKm.Text = ""
    txtEndingKm.Text = ""
    txtKmTravelled.Text = ""
    lblRatePerKm.Caption = ""
    txtCostOfTravel.Text = ""
    dtpBookingDate.Value = Date
    dtpJourneyDate.Value = Date
    dtpDropOffDate.Value = Date
    dtpBillingDate.Value = Date
    txtLateFine.Text = ""
    txtBookingCharge.Text = ""
    txtDriverCharge.Text = ""
    txtCostOfTravel2.Text = ""
    txtLateFine2.Text = ""
    txtAdditionalExp.Text = ""
    txtGross.Text = ""
    txtDiscount.Text = ""
    txtAdvancePaid.Text = ""
    txtNet.Text = ""
    
    txtBookingID.Locked = False
End Sub
'Save button click event.
Private Sub cmdSave_Click()
'Validation check to make sure none of the fields are left empty.
    If txtBookingID.Text = "" Or txtName.Text = "" Or txtContact.Text = "" Or txtAddress.Text = Empty Or txtVehicleType.Text = "" Or txtVehicleID.Text = "" Or txtMake.Text = "" Or txtModel.Text = "" Or txtRegNum.Text = "" Or txtDriverName.Text = "" Or lblNSeats.Caption = "" Or txtStartingKm.Text = "" Or txtEndingKm.Text = "" Or txtKmTravelled.Text = "" Or txtCostOfTravel.Text = "" Or txtLateFine.Text = "" Or txtBookingCharge.Text = "" Or txtDriverCharge.Text = "" Or txtAdditionalExp.Text = "" Or txtGross.Text = "" Or txtDiscount.Text = "" Or txtAdvancePaid.Text = "" Or txtNet.Text = "" Then
'Message shown to notify the User that none of the fields can be left empty.
        MsgBox "Please make sure none of the fields are left empty and try again. "
        Exit Sub
    End If
    
'Length check to make sure contact number is of 11 digits.
    If Len(txtContact.Text) <> 11 Then
        MsgBox "Number of digits used for contact number is invalid."
        Exit Sub
    End If
    
'Storage of Billing information in the Vehicle Billing database.
      Set rs = New Recordset
      rs.Open "select * from Vehicle_billing", con, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("BookingID") = txtBookingID.Text
            rs.Fields("Customer_name") = txtName.Text
            rs.Fields("Customer_address") = txtAddress.Text
            rs.Fields("VehicleID") = txtVehicleID.Text
            rs.Fields("Kilometer_travelled") = txtKmTravelled.Text
            rs.Fields("Rate_per_kilometer") = lblRatePerKm.Caption
            rs.Fields("Cost_of_travel") = txtCostOfTravel.Text
            rs.Fields("Journey_date") = dtpJourneyDate.Value
            rs.Fields("Drop_off_date") = dtpDropOffDate.Value
            rs.Fields("Billing_date") = dtpBillingDate.Value
            rs.Fields("Booking_charge") = txtBookingCharge.Text
            rs.Fields("Driver_charge") = txtDriverCharge.Text
            rs.Fields("Cost_of_travel") = txtCostOfTravel.Text
            rs.Fields("Late_fine") = txtLateFine.Text
            rs.Fields("Additional_expenditures") = txtAdditionalExp.Text
            rs.Fields("Gross_total") = txtGross.Text
            rs.Fields("Discount") = txtDiscount.Text
            rs.Fields("Advance_paid") = txtAdvancePaid.Text
            rs.Fields("Net_total") = txtNet.Text
        rs.Update
                    
        MsgBox "The Billing information is saved."

End Sub
'Form load click event.
Private Sub Form_Load()
'Calling the function to connect to database.
    Call dblink
    lblDate.Caption = Date
    dtpBillingDate.Value = Date

End Sub

Private Sub txtAdditionalExp_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtBookingCharge_txtBookingCharge_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub


Private Sub txtAdvancePaid_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtCostOfTravel_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtCostOfTravel2_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtDriverCharge_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
Private Sub txtEndingKm_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtKmTravelled_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

'Kilometers travelled textbox lostfocus event.
Private Sub txtKmTravelled_LostFocus()
'Calculation of total amount charged to the customer for the total distance travelled.
    txtKmTravelled.Text = Val(txtEndingKm.Text) - Val(txtStartingKm.Text)
    txtCostOfTravel.Text = Val(txtKmTravelled.Text) * Val(lblRatePerKm.Caption)
    txtCostOfTravel2.Text = txtCostOfTravel.Text
End Sub

Private Sub txtLateFine_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub

Private Sub txtLateFine2_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub


Private Sub txtNet_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If
End Sub

Private Sub txtStartingKm_KeyPress(KeyAscii As Integer)
'Character type check to make sure non-numerical data is not allowed to be entered.
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
            MsgBox "Please input numbers only."
        Exit Sub
    End If

End Sub
