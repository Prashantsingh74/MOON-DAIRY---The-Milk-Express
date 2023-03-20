VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Product 
   Caption         =   "Product Entry"
   ClientHeight    =   9015
   ClientLeft      =   3630
   ClientTop       =   1065
   ClientWidth     =   13830
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   13830
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   360
      TabIndex        =   7
      Top             =   0
      Width           =   13095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Product Entry.frx":0000
         Height          =   3015
         Left            =   840
         TabIndex        =   50
         Top             =   5640
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5318
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   28
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Product View"
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "PR_NAME"
            Caption         =   "PRODUCT NAME"
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
            DataField       =   "WEIGHT"
            Caption         =   "WEIGHT"
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
            DataField       =   "UNIT"
            Caption         =   "UNIT"
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
            DataField       =   "MRP"
            Caption         =   "MRP"
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
               ColumnWidth     =   2025.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2550.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command10 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         TabIndex        =   46
         Top             =   8040
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8705
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         BackColor       =   16761024
         TabCaption(0)   =   "Entry"
         TabPicture(0)   =   "Product Entry.frx":0015
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Command3"
         Tab(0).Control(1)=   "Command2"
         Tab(0).Control(2)=   "Command1"
         Tab(0).Control(3)=   "Text3"
         Tab(0).Control(4)=   "Text2"
         Tab(0).Control(5)=   "Combo1"
         Tab(0).Control(6)=   "Text5"
         Tab(0).Control(7)=   "Text4"
         Tab(0).Control(8)=   "Text1"
         Tab(0).Control(9)=   "Label6"
         Tab(0).Control(10)=   "Label5"
         Tab(0).Control(11)=   "Label4"
         Tab(0).Control(12)=   "Label3"
         Tab(0).Control(13)=   "Label2"
         Tab(0).Control(14)=   "Label1"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Update"
         TabPicture(1)   =   "Product Entry.frx":0031
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command6"
         Tab(1).Control(1)=   "Command5"
         Tab(1).Control(2)=   "Command4"
         Tab(1).Control(3)=   "Text13"
         Tab(1).Control(4)=   "Combo2"
         Tab(1).Control(5)=   "Text11"
         Tab(1).Control(6)=   "Text10"
         Tab(1).Control(7)=   "Text8"
         Tab(1).Control(8)=   "Text7"
         Tab(1).Control(9)=   "Label16"
         Tab(1).Control(10)=   "Label14"
         Tab(1).Control(11)=   "Label13"
         Tab(1).Control(12)=   "Label12"
         Tab(1).Control(13)=   "Label10"
         Tab(1).Control(14)=   "Label9"
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "Delete"
         TabPicture(2)   =   "Product Entry.frx":004D
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label17"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label18"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label20"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label21"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label22"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Label24"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Text14"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Text15"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Text17"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Text18"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "Text20"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "Combo3"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "Command7"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "Command8"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "Command9"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).ControlCount=   15
         Begin VB.CommandButton Command9 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8280
            TabIndex        =   47
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6600
            TabIndex        =   45
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7680
            TabIndex        =   44
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -66720
            TabIndex        =   43
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000D&
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -68400
            TabIndex        =   42
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -67440
            TabIndex        =   41
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -68760
            TabIndex        =   40
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -70560
            TabIndex        =   6
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000FF00&
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -72720
            TabIndex        =   0
            Top             =   3840
            Width           =   1575
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   39
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7680
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   36
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   35
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   34
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2520
            TabIndex        =   33
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -67440
            TabIndex        =   27
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   25
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   24
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   23
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   22
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   21
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   2
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   1
            Top             =   1560
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -67800
            TabIndex        =   4
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -67800
            TabIndex        =   5
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -67800
            TabIndex        =   3
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -72360
            TabIndex        =   15
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Id"
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
            Left            =   6120
            TabIndex        =   38
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "M.R.P"
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
            TabIndex        =   32
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
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
            TabIndex        =   31
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            TabIndex        =   30
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name"
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
            TabIndex        =   29
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Id"
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
            TabIndex        =   28
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Id"
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
            Left            =   -69120
            TabIndex        =   26
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "M.R.P"
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
            Left            =   -74520
            TabIndex        =   20
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
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
            Left            =   -74520
            TabIndex        =   19
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            Left            =   -74520
            TabIndex        =   18
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name"
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
            Left            =   -74520
            TabIndex        =   17
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Id"
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
            Left            =   -74520
            TabIndex        =   16
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "M.R.P"
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
            Left            =   -69360
            TabIndex        =   14
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
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
            Left            =   -69360
            TabIndex        =   13
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            Left            =   -69360
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Company"
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
            Left            =   -74760
            TabIndex        =   11
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name"
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
            Left            =   -74640
            TabIndex        =   10
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Product Id"
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
            Left            =   -74640
            TabIndex        =   9
            Top             =   840
            Width           =   1455
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   10080
         Top             =   1920
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
         Connect         =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=MOON/ADMIN;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from product_entry"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT ENTRY"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   4920
         TabIndex        =   52
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "0"
         Height          =   375
         Left            =   7680
         TabIndex        =   51
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "NULL"
         Height          =   375
         Left            =   9240
         TabIndex        =   49
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "NULL"
         Height          =   375
         Left            =   8520
         TabIndex        =   48
         Top             =   1800
         Width           =   615
      End
   End
End
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Command1_Click()
conn
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Command3.Enabled = True
Text1.Locked = True
Text2.SetFocus
End Sub

Private Sub Command10_Click()
Unload Me
home.Show
End Sub

Private Sub Command2_Click()
If Text2.Text = "" Or Text5.Text = "" Then
MsgBox "All Fields Required"
Else
sql = "insert into product_entry values('" + Text1.Text + "','" + Text2.Text + "','" + Text4.Text + "','" + Combo1.Text + "','" + Text5.Text + "')"
Set r = c.Execute(sql)
sql = "INSERT INTO STOCK VALUES('" + Text1.Text + "','" + Text2.Text + "'," + Label11.Caption + "," + Label15.Caption + "," + Label7.Caption + ")"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""


End Sub

Private Sub Command4_Click()
On Error GoTo ABC
Text7.Enabled = True
Text8.Enabled = True
'Text9.Enabled = True
Combo2.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
sql = " select * from product_entry where pr_id='" + Text13.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Text8.Text = r.Fields(1)
'Text9.Text = R.Fields(2)
Text10.Text = r.Fields(2)
Combo2.Text = r.Fields(3)
Text11.Text = r.Fields(4)
Text13.Text = ""
Command4.Enabled = False
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
Text13.Text = ""
Text13.SetFocus
End Sub

Private Sub Command5_Click()
sql = "update product_entry set pr_id='" + Text7.Text + "',pr_name='" + Text8.Text + "',weight=" + Text10.Text + ",unit='" + Combo2.Text + "',MRP=" + Text11.Text + " where pr_id='" + Text7.Text + "'" 'or pr_name='"+ text13.text +"'
Set r = c.Execute(sql)
MsgBox "product updated succesfully"
Adodc1.Refresh
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
Combo2.Text = ""
Text11.Text = ""
End Sub

Private Sub Command6_Click()
Text7.Text = ""
Text8.Text = ""
'Text9.Text = ""
Text10.Text = ""
Combo2.Text = ""
Text11.Text = ""
Text13.Text = ""
Text7.Enabled = False
Text8.Enabled = False
Combo2.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Text13.SetFocus
End Sub

Private Sub Command7_Click()
On Error GoTo ABC
Text14.Enabled = True
Text15.Enabled = True
'Text16.Enabled = True
Text17.Enabled = True
Combo3.Enabled = True
Text18.Enabled = True
Text14.Locked = True
Text15.Locked = True
'Text16.Locked = True
Text17.Locked = True
Combo3.Locked = True
Text18.Locked = True
Command8.Enabled = True
Command9.Enabled = True
sql = " select * from product_entry where pr_id='" + Text20.Text + "'"
Set r = c.Execute(sql)
Text14.Text = r.Fields(0)
Text15.Text = r.Fields(1)
'Text16.Text = R.Fields(2)
Text17.Text = r.Fields(2)
Combo3.Text = r.Fields(3)
Text18.Text = r.Fields(4)
Text20.Text = ""
Command7.Enabled = False
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
Text20.Text = ""
Text20.SetFocus
End Sub

Private Sub Command8_Click()
sql = "delete from product_entry where pr_id='" + Text14.Text + "'"
Set r = c.Execute(sql)
MsgBox "product removed"
Adodc1.Refresh
Text14.Text = ""
Text15.Text = ""
'Text16.Text = ""
Text17.Text = ""
Combo3.Text = ""
Text18.Text = ""

End Sub

Private Sub Command9_Click()
Text14.Text = ""
Text15.Text = ""
Text17.Text = ""
Text18.Text = ""
Combo3.Text = ""
Text14.Enabled = False
Text15.Enabled = False
Combo3.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Command8.Enabled = False
Text20.SetFocus
End Sub

Private Sub Form_Load()
conn
Combo1.AddItem "gm"
Combo1.AddItem "Kg"
Combo1.AddItem "ml"
Combo1.AddItem "Ltr"
Combo2.AddItem "gm"
Combo2.AddItem "Kg"
Combo2.AddItem "ml"
Combo2.AddItem "Ltr"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
'Text9.Enabled = False
Combo2.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
'Text16.Enabled = False
Combo3.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
End Sub



Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.Text = UCase(Text10.Text)
Exit Sub
End If
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
End Sub


Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text11.Text = UCase(Text11.Text)
Command5.SetFocus
Exit Sub
End If
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
End Sub

Private Sub Text13_GotFocus()
Command4.Enabled = True
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
'Command4.Enabled = True
If KeyAscii = 13 Then
Command4.SetFocus
End If
End Sub

Private Sub Text13_LostFocus()
Text13.Text = UCase(Text13.Text)
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
Text3.SetFocus
Exit Sub
End If
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 8) Or (KeyAscii = 32) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter character only")
End If
'End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command7.Enabled = True
Command7.SetFocus
End If
End Sub

Private Sub Text20_LostFocus()
Text20.Text = UCase(Text20.Text)
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Text = UCase(Text3.Text)
Text4.SetFocus
Exit Sub
End If
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter character only")
End If
End Sub
Private Sub text4_keypress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
Combo1.SetFocus
Text1.Text = Text2.Text + Text4.Text
End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Not IsNumeric(Text4.Text) Then
cancle = True
Combo1.Enabled = False
Text5.Enabled = False
Text4.Text = ""
Text4.SetFocus
MsgBox "enter numberic value"
Else
cancle = False
Combo1.Enabled = True
Text5.Enabled = True
End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
If Not IsNumeric(Text5.Text) Then
cancle = True
MsgBox "enter numberic value"
Text5.Text = ""
Text5.SetFocus
Else
cancle = False
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
Command2.Enabled = True
Command2.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
Command2.Enabled = True
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.Text = UCase(Text8.Text)
Text9.SetFocus
Exit Sub
End If
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 8) Or (KeyAscii = 32) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter character only")
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.Text = UCase(Text9.Text)
Exit Sub
End If
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 8) Or (KeyAscii = 32) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter character only")
End If
End Sub
