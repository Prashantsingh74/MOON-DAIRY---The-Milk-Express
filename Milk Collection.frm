VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MilkCollection 
   BackColor       =   &H00FF8080&
   Caption         =   "Collection"
   ClientHeight    =   6795
   ClientLeft      =   4275
   ClientTop       =   3060
   ClientWidth     =   12495
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   12495
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      TabIndex        =   47
      Top             =   5760
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Morning"
      TabPicture(0)   =   "Milk Collection.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Evening"
      TabPicture(1)   =   "Milk Collection.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Evening Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   9855
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3120
            TabIndex        =   51
            Top             =   3360
            Width           =   735
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8400
            TabIndex        =   46
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8400
            TabIndex        =   45
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6360
            TabIndex        =   44
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5040
            TabIndex        =   43
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2160
            TabIndex        =   42
            Top             =   3360
            Width           =   855
         End
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3960
            TabIndex        =   41
            Text            =   "Fat"
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            TabIndex        =   40
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   39
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   38
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   37
            Text            =   "Farmer ID"
            Top             =   1080
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   7800
            TabIndex        =   36
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   225640449
            CurrentDate     =   44985
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "unit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   50
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6360
            TabIndex        =   34
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   33
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Fat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   32
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   31
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   29
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   28
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   25
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Morning Colletion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9735
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3360
            TabIndex        =   49
            Top             =   3360
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8400
            TabIndex        =   24
            Top             =   3840
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8400
            TabIndex        =   23
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6600
            TabIndex        =   22
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   21
            Top             =   3360
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4200
            TabIndex        =   20
            Text            =   "fat"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2280
            TabIndex        =   19
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   360
            TabIndex        =   18
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   12
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   11
            Top             =   1680
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5160
            TabIndex        =   8
            Text            =   "Farmer ID"
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7680
            TabIndex        =   6
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   225640449
            CurrentDate     =   44985
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1320
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   48
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   17
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   16
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   10
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Famer Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   9
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "MILK COLLECTION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4200
      TabIndex        =   54
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   255
      Left            =   8520
      TabIndex        =   53
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   375
      Left            =   8760
      TabIndex        =   52
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "MilkCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Integer
Private Sub Combo1_Click()
sql = "select f_name from farmer_entry where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select ph_no from farmer_entry where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select pr_id from farmer_entry where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "SELECT BALANCE FROM STOCK WHERE PR_ID='" + Text4.Text + "'"
Set r = c.Execute(sql)
Label23.Caption = r.Fields(0)

Text5.SetFocus
End Sub
Private Sub Combo2_Click()
If Text5.Text = "" Then
MsgBox "enter Quantity"
Else
sql = "select rate from fatlist where fat='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
Text7.Text = Val(Text6.Text * Text5.Text)
End If
Command2.SetFocus
End Sub

Private Sub Combo3_Click()
sql = "select f_name from farmer_entry where f_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text9.Text = r.Fields(0)
sql = "select ph_no from farmer_entry where f_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text10.Text = r.Fields(0)
sql = "select pr_id from farmer_entry where f_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text11.Text = r.Fields(0)
Text12.SetFocus
sql = "SELECT BALANCE FROM STOCK WHERE PR_ID='" + Text11.Text + "'"
Set r = c.Execute(sql)
Label23.Caption = r.Fields(0)

End Sub

Private Sub Combo4_Click()
If Text12.Text = "" Then
MsgBox "enter Quantity"
Else
sql = "select rate from fatlist where fat='" + Combo4.Text + "'"
Set r = c.Execute(sql)
Text13.Text = r.Fields(0)
Text14.Text = Val(Text13.Text * Text12.Text)
Command4.SetFocus
End If
End Sub
Private Sub Command1_Click()
Combo1.clear
'Sql = "select count(serialNo) from morning_coll"
'Set R = C.Execute(Sql)
'Text17.Text = R.Fields(0) + 1
Text1.Text = "Morning"
Text15.Text = "ltr"

sql = "select f_id from farmer_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select Fat from fatlist"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text15.Enabled = True
DTPicker1.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo1.SetFocus
Command2.Enabled = True
End Sub
Private Sub Command2_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Text5.Text = "" Then
MsgBox "All Fields required"
Else
sql = "insert into morning_coll values('" + Text1.Text + "','" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Combo1.Text + "'," + Text5.Text + ",'" + Combo2.Text + "'," + Text6.Text + "," + Text7.Text + ")"
Set r = c.Execute(sql)
sql = "UPDATE STOCK SET BALANCE=" + Label24.Caption + " WHERE PR_ID= '" + Text4.Text + "'"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Combo1.Text = ""
Combo2.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.SetFocus
End If
End Sub
Private Sub Command3_Click()
Combo1.clear
'Sql = "select count(serialNo) from evening_coll"
'Set R = C.Execute(Sql)
'Text18.Text = R.Fields(0) + 1
Text8.Text = "Evening"
Text16.Text = "ltr"
sql = "select f_id from farmer_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select Fat from fatlist"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text16.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
DTPicker2.Enabled = True
Combo3.SetFocus
Command4.Enabled = True
End Sub
Private Sub Command4_Click()
If Combo3.Text = "" Or Combo4.Text = "" Or Text12.Text = "" Then
MsgBox "All Fields Required"
Else
sql = "insert into evening_coll values('" + Text8.Text + "','" + Format(DTPicker2.Value, "dd mmm yyyy") + "','" + Combo3.Text + "'," + Text12.Text + ",'" + Combo4.Text + "'," + Text13.Text + "," + Text14.Text + ")"
Set r = c.Execute(sql)
sql = "UPDATE STOCK SET BALANCE=" + Label24.Caption + " WHERE PR_ID= '" + Text11.Text + "'"
Set r = c.Execute(sql)
MsgBox "Record Saved"
'p = Combo3.li
'If p > -1 Then
'Combo1.RemoveItem p
'End If
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo3.SetFocus
End If
End Sub
Private Sub Command5_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
conn
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
Text16.Locked = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Combo4.SetFocus
End If
End Sub

Private Sub Text12_LostFocus()
Label24.Caption = Val(Label23.Caption) + Val(Text12.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
Label24.Caption = Val(Label23.Caption) + Val(Text5.Text)

End Sub
