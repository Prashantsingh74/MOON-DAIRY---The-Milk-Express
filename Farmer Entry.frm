VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FarmerEntry 
   BackColor       =   &H00FF8080&
   Caption         =   " FARMER ENTRY"
   ClientHeight    =   8805
   ClientLeft      =   4605
   ClientTop       =   1740
   ClientWidth     =   14520
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   14520
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Farmer Entry.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   96
      Top             =   6840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "F_ID"
         Caption         =   "FARMER ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "F_NAME"
         Caption         =   "FARMER NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FATHER_NAME"
         Caption         =   "FATHER NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PH_NO"
         Caption         =   "PHONE NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AADHAR"
         Caption         =   "AADHAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ADDRESS"
         Caption         =   "ADDRESS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PR_ID"
         Caption         =   "PRODUCT ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2324.977
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9763
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ENTRY"
      TabPicture(0)   =   "Farmer Entry.frx":0015
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Label32"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "UPDATE"
      TabPicture(1)   =   "Farmer Entry.frx":0031
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "DELETE"
      TabPicture(2)   =   "Farmer Entry.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFC0C0&
         Height          =   4215
         Left            =   -63360
         TabIndex        =   91
         Top             =   840
         Width           =   1935
         Begin VB.CommandButton Command8 
            Caption         =   "Delete"
            Height          =   615
            Left            =   240
            TabIndex        =   95
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Search"
            Height          =   615
            Left            =   240
            TabIndex        =   94
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox Text34 
            Height          =   420
            Left            =   240
            TabIndex        =   93
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Search By Farmer ID"
            Height          =   615
            Left            =   360
            TabIndex        =   92
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFC0C0&
         Height          =   4215
         Left            =   11640
         TabIndex        =   86
         Top             =   840
         Width           =   1935
         Begin VB.CommandButton Command6 
            Caption         =   "Update"
            Height          =   735
            Left            =   240
            TabIndex        =   90
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   615
            Left            =   120
            TabIndex        =   89
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox Text33 
            Height          =   420
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Search By Farmer ID"
            Height          =   615
            Left            =   360
            TabIndex        =   87
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         Height          =   3975
         Left            =   -63240
         TabIndex        =   79
         Top             =   960
         Width           =   1815
         Begin VB.CommandButton Command4 
            Caption         =   "Exit"
            Height          =   615
            Left            =   120
            TabIndex        =   83
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            Height          =   615
            Left            =   120
            TabIndex        =   82
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save"
            Height          =   615
            Left            =   120
            TabIndex        =   81
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ADD New"
            Height          =   615
            Left            =   120
            TabIndex        =   80
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer Details"
         ForeColor       =   &H00000000&
         Height          =   2895
         Left            =   -74760
         TabIndex        =   50
         Top             =   480
         Width           =   11295
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   855
            Left            =   2520
            TabIndex        =   64
            Top             =   1920
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Farmer Entry.frx":0069
         End
         Begin VB.ComboBox Combo1 
            Height          =   420
            Left            =   8160
            TabIndex        =   63
            Text            =   "Product ID"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   55
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   54
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   53
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   52
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   51
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label41 
            Caption         =   "PRODUCT SUPPLY"
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
            Left            =   8160
            TabIndex        =   97
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label34 
            Height          =   375
            Left            =   10080
            TabIndex        =   85
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID"
            Height          =   300
            Left            =   6000
            TabIndex        =   62
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Index           =   1
            Left            =   480
            TabIndex        =   61
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aadhar No"
            Height          =   300
            Index           =   1
            Left            =   6000
            TabIndex        =   60
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Index           =   0
            Left            =   480
            TabIndex        =   59
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   300
            Left            =   6000
            TabIndex        =   58
            Top             =   960
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer Name"
            Height          =   300
            Left            =   480
            TabIndex        =   57
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer ID"
            Height          =   375
            Left            =   480
            TabIndex        =   56
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer Bank Details"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   39
         Top             =   3480
         Width           =   11295
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   44
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   43
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   42
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   41
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2520
            TabIndex        =   40
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
            Height          =   300
            Left            =   480
            TabIndex        =   49
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   300
            Left            =   480
            TabIndex        =   48
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            Height          =   300
            Left            =   480
            TabIndex        =   47
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "A/C Holder Name"
            Height          =   375
            Left            =   6000
            TabIndex        =   46
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "IFSC Code"
            Height          =   375
            Left            =   6000
            TabIndex        =   45
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer Details"
         Height          =   2895
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   11175
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8040
            TabIndex        =   66
            Top             =   2400
            Width           =   1575
         End
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   855
            Left            =   2280
            TabIndex        =   65
            Top             =   1920
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1508
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Farmer Entry.frx":00EB
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   31
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   30
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8040
            TabIndex        =   29
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   28
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8040
            TabIndex        =   27
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID"
            Height          =   300
            Left            =   6120
            TabIndex        =   38
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   480
            TabIndex        =   37
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Aadhar No"
            Height          =   300
            Left            =   6000
            TabIndex        =   36
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Left            =   480
            TabIndex        =   35
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   300
            Left            =   6000
            TabIndex        =   34
            Top             =   960
            Width           =   1560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer Name"
            Height          =   300
            Left            =   480
            TabIndex        =   33
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer ID"
            Height          =   300
            Left            =   480
            TabIndex        =   32
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer bank Details"
         Height          =   1815
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   11175
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   20
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8280
            TabIndex        =   19
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   18
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8280
            TabIndex        =   17
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2280
            TabIndex        =   16
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   300
            Left            =   480
            TabIndex        =   25
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C Holder Name"
            Height          =   300
            Left            =   6000
            TabIndex        =   24
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            Height          =   300
            Left            =   480
            TabIndex        =   23
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
            Height          =   375
            Left            =   480
            TabIndex        =   22
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "IFSC Code"
            Height          =   375
            Left            =   6000
            TabIndex        =   21
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer Details"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   11175
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   855
            Left            =   2640
            TabIndex        =   78
            Top             =   1800
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1508
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"Farmer Entry.frx":016D
         End
         Begin VB.TextBox Text27 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   72
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox Text26 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   71
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text25 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   70
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   69
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   68
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   67
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID"
            Height          =   300
            Left            =   6120
            TabIndex        =   14
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   360
            TabIndex        =   13
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aadhar No"
            Height          =   300
            Left            =   6120
            TabIndex        =   12
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Left            =   360
            TabIndex        =   11
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   300
            Left            =   6120
            TabIndex        =   10
            Top             =   840
            Width           =   1560
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer Name"
            Height          =   300
            Left            =   360
            TabIndex        =   9
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Farmer ID"
            Height          =   300
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Farmer Bank Details"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   1
         Top             =   3360
         Width           =   11175
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   77
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text31 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   76
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox Text30 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   75
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox Text29 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   8160
            TabIndex        =   74
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox Text28 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2640
            TabIndex        =   73
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Name"
            Height          =   300
            Left            =   360
            TabIndex        =   6
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IFSC Code"
            Height          =   300
            Left            =   6120
            TabIndex        =   5
            Top             =   840
            Width           =   1185
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   300
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C Holder Name"
            Height          =   300
            Left            =   6120
            TabIndex        =   3
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            Height          =   300
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Label Label32 
         Caption         =   "0"
         Height          =   375
         Left            =   -63120
         TabIndex        =   84
         Top             =   480
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10680
      Top             =   7560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=moon/admin;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=moon/admin;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from farmer_entry"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "FARMER ENTRY"
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
      Left            =   5520
      TabIndex        =   98
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "FarmerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub Combo1_Click()
Label34.Caption = Combo1.Text
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Command1_Click()
a = "F00"
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text9.Enabled = True
Text8.Enabled = True
Text10.Enabled = True
Combo1.Enabled = True
RichTextBox1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
sql = "select count(f_id) from farmer_Entry"
Set r = c.Execute(sql)
Text1.Text = a & r.Fields(0) + 1
Text1.Locked = True
Text2.SetFocus
sql = "select pr_id from product_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0) + ""
r.MoveNext
Loop
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or RichTextBox1.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "All Fields Reqired"
Else
sql = "insert into Farmer_Entry values('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "'," + Text4.Text + "," + Text5.Text + ",'" + RichTextBox1.Text + "','" + Label34.Caption + "')"
Set r = c.Execute(sql)
sql = "insert into Farmerbank_details values('" + Text1.Text + "', '" + Text6.Text + "','" + Text7.Text + "'," + Text8.Text + ",'" + Text9.Text + "','" + Text10.Text + "' )"
Set r = c.Execute(sql)
sql = "insert into farmer_dues values('" + Text1.Text + "','" + Label32.Caption + "')"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Label34.Caption = ""
RichTextBox1.Text = ""
Adodc1.Refresh
End If
End Sub

Private Sub Command3_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
home.Show
End Sub

Private Sub Command5_Click()
On Error GoTo ABC
If Text33.Text = "" Then
MsgBox "Farmer ID Required"
Else
sql = "SELECT * FROM FARMER_ENTRY WHERE F_ID='" + Text33.Text + "'"
Set r = c.Execute(sql)
Text11.Text = r.Fields(0)
Text12.Text = r.Fields(1)
Text13.Text = r.Fields(2)
Text14.Text = r.Fields(3)
Text15.Text = r.Fields(4)
RichTextBox2.Text = r.Fields(5)
Text21.Text = r.Fields(6)
sql = "SELECT * FROM Farmerbank_details WHERE F_ID='" + Text33.Text + "'"
Set r = c.Execute(sql)
Text16.Text = r.Fields(1) & ""
Text17.Text = r.Fields(2) & ""
Text18.Text = r.Fields(3) & ""
Text19.Text = r.Fields(4) & ""
Text20.Text = r.Fields(5) & ""
Text33.Text = ""
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text20.Enabled = True
Text21.Enabled = True
RichTextBox2.Enabled = True
Text11.Locked = True
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
End If
End Sub

Private Sub Command6_Click()
sql = "update FARMER_entry set  f_name='" + Text12.Text + "',father_name='" + Text13.Text + "',ph_no=" + Text14.Text + ",aadhar=" + Text15.Text + ",address='" + RichTextBox2.Text + "',Pr_ID='" + Text21.Text + "' WHERE F_ID='" + Text11.Text + "'"
Set r = c.Execute(sql)
sql = "update farmerbank_details SET BANK_NAME='" + Text16.Text + "',acc_no=" + Text18.Text + ",Ifsc='" + Text19.Text + "',acc_holdername='" + Text17.Text + "',branch_name='" + Text20.Text + "' WHERE F_ID='" + Text11.Text + "'"
Set r = c.Execute(sql)
MsgBox "Record Updated"
Text33.SetFocus
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
RichTextBox2.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
End Sub

Private Sub Command7_Click()
On Error GoTo ABC
If Text34.Text = "" Then
MsgBox "Farmer ID Required"
Else
sql = "SELECT * FROM FARMER_ENTRY WHERE F_ID='" + Text34.Text + "'"
Set r = c.Execute(sql)
Text22.Text = r.Fields(0)
Text23.Text = r.Fields(1)
Text24.Text = r.Fields(2)
Text25.Text = r.Fields(3)
Text26.Text = r.Fields(4)
RichTextBox3.Text = r.Fields(5)
Text27.Text = r.Fields(6)
sql = "SELECT * FROM Farmerbank_details WHERE F_ID='" + Text34.Text + "'"
Set r = c.Execute(sql)
Text28.Text = r.Fields(1) & ""
Text29.Text = r.Fields(2) & ""
Text30.Text = r.Fields(3) & ""
Text31.Text = r.Fields(4) & ""
Text32.Text = r.Fields(5) & ""
Text22.Enabled = True
Text23.Enabled = True
Text24.Enabled = True
Text25.Enabled = True
Text26.Enabled = True
Text27.Enabled = True
Text28.Enabled = True
Text29.Enabled = True
Text30.Enabled = True
Text31.Enabled = True
Text32.Enabled = True
RichTextBox3.Enabled = True
Text22.Locked = True
Text23.Locked = True
Text24.Locked = True
Text25.Locked = True
Text26.Locked = True
Text27.Locked = True
Text28.Locked = True
Text29.Locked = True
Text30.Locked = True
Text31.Locked = True
Text32.Locked = True
RichTextBox3.Enabled = True
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
End If
End Sub

Private Sub Command8_Click()
MsgBox "Record can't be deleted"
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
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False
Text24.Enabled = False
Text25.Enabled = False
Text26.Enabled = False
Text27.Enabled = False
Text28.Enabled = False
Text30.Enabled = False
Text29.Enabled = False
Text31.Enabled = False
Text32.Enabled = False
Combo1.Enabled = False
RichTextBox1.Enabled = False
RichTextBox2.Enabled = False
RichTextBox3.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub RichTextBox1_LostFocus()
RichTextBox1.Text = UCase(RichTextBox1.Text)
End Sub

Private Sub RichTextBox2_LostFocus()
RichTextBox2.Text = UCase(RichTextBox2.Text)
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub

Private Sub Text10_LostFocus()
Text10.Text = UCase(Text10.Text)
End Sub

Private Sub Text12_LostFocus()
Text12.Text = UCase(Text12.Text)
End Sub

Private Sub Text13_LostFocus()
Text13.Text = UCase(Text13.Text)
End Sub

Private Sub Text16_LostFocus()
Text16.Text = UCase(Text16.Text)
End Sub

Private Sub Text17_LostFocus()
Text17.Text = UCase(Text17.Text)
End Sub

Private Sub Text19_LostFocus()
Text19.Text = UCase(Text19.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub

Private Sub Text20_LostFocus()
Text20.Text = UCase(Text20.Text)
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub Text33_LostFocus()
Text33.Text = UCase(Text33.Text)
End Sub

Private Sub Text34_LostFocus()
Text34.Text = UCase(Text34.Text)
End Sub

Private Sub text4_keypress(KeyAscii As Integer)
Text4.MaxLength = 10
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Text5.MaxLength = 12
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
RichTextBox1.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
Text6.Text = UCase(Text6.Text)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Text8.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
Text7.Text = UCase(Text7.Text)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
Text9.Text = UCase(Text9.Text)
End Sub
