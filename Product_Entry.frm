VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ProductEntry 
   Caption         =   "Form2"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18030
   LinkTopic       =   "Form2"
   ScaleHeight     =   10215
   ScaleWidth      =   18030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
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
      Height          =   7935
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   12015
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   10680
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Connect         =   "Provider=MSDAORA.1;User ID=MOONDAIRY/MOON;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=MOONDAIRY/MOON;Persist Security Info=False"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   480
         TabIndex        =   48
         Top             =   5400
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "PRODUCT VIEW"
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2115.213
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
      Begin VB.CommandButton exit 
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
         Left            =   10680
         TabIndex        =   46
         Top             =   7200
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4695
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8281
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BackColor       =   0
         TabCaption(0)   =   "Entry"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(3)=   "Label4"
         Tab(0).Control(4)=   "Label5"
         Tab(0).Control(5)=   "Label6"
         Tab(0).Control(6)=   "Text1"
         Tab(0).Control(7)=   "Text4"
         Tab(0).Control(8)=   "Text5"
         Tab(0).Control(9)=   "Combo1"
         Tab(0).Control(10)=   "Text2"
         Tab(0).Control(11)=   "Text3"
         Tab(0).Control(12)=   "addnew"
         Tab(0).Control(13)=   "save"
         Tab(0).Control(14)=   "clear"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Update"
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label9"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label10"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label12"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label13"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label14"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label16"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Text7"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Text8"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Text10"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Text11"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Combo2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Text13"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "search"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "update"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Command6"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "Delete"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label17"
         Tab(2).Control(1)=   "Label18"
         Tab(2).Control(2)=   "Label20"
         Tab(2).Control(3)=   "Label21"
         Tab(2).Control(4)=   "Label22"
         Tab(2).Control(5)=   "Label24"
         Tab(2).Control(6)=   "Text14"
         Tab(2).Control(7)=   "Text15"
         Tab(2).Control(8)=   "Text17"
         Tab(2).Control(9)=   "Text18"
         Tab(2).Control(10)=   "Text20"
         Tab(2).Control(11)=   "Combo3"
         Tab(2).Control(12)=   "Command7"
         Tab(2).Control(13)=   "delete"
         Tab(2).Control(14)=   "Command9"
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
            Left            =   -66720
            TabIndex        =   47
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton delete 
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
            Left            =   -68400
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
            Left            =   -67320
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
            Left            =   8280
            TabIndex        =   43
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton update 
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
            Left            =   6600
            TabIndex        =   42
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton search 
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
            Left            =   7560
            TabIndex        =   41
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton clear 
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
            Left            =   -69480
            TabIndex        =   40
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton save 
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
            Left            =   -71280
            TabIndex        =   6
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton addnew 
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
            Left            =   -73200
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
            Left            =   -72120
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
            Left            =   -67320
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
            Left            =   -72120
            TabIndex        =   36
            Top             =   3840
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
            Left            =   -72120
            TabIndex        =   35
            Top             =   2280
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
            Left            =   -72120
            TabIndex        =   34
            Top             =   1560
            Width           =   1815
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
            Left            =   -72120
            TabIndex        =   33
            Top             =   840
            Width           =   1815
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
            Left            =   7560
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
            Left            =   2640
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
            Left            =   2640
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
            Left            =   2640
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
            Left            =   2640
            TabIndex        =   22
            Top             =   1440
            Width           =   1695
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
            Left            =   2640
            TabIndex        =   21
            Top             =   720
            Width           =   1695
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
            Left            =   -68880
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
            Left            =   -74640
            TabIndex        =   32
            Top             =   3720
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
            Left            =   -74640
            TabIndex        =   31
            Top             =   3000
            Width           =   495
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
            Left            =   -74640
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
            Left            =   -74640
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
            Left            =   -74640
            TabIndex        =   28
            Top             =   840
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
            Left            =   6000
            TabIndex        =   26
            Top             =   960
            Width           =   1335
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
            Left            =   360
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
            Left            =   360
            TabIndex        =   19
            Top             =   2880
            Width           =   495
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
            Left            =   360
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
            Left            =   360
            TabIndex        =   17
            Top             =   1440
            Width           =   1815
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
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   1335
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
   End
End
Attribute VB_Name = "ProductEntry"
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

Private Sub addnew_Click()
conn
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
clear.Enabled = True
Text1.Locked = True
Text2.SetFocus
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub save_Click()
If Text2.Text = "" Or Text4.Text = "" Or Text3.Text = "" Or Text5.Text = "" Then
MsgBox "all fields required"
Else
Sql = "insert into product_entry values('" + Text1.Text + "','" + Text2.Text + "','" + Text4.Text + "','" + Combo1.Text + "','" + Text5.Text + "')"
Set R = C.Execute(Sql)
MsgBox "Record saved"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""


End Sub

Private Sub search_Click()
On Error GoTo ABC
Text7.Enabled = True
Text8.Enabled = True
'Text9.Enabled = True
Combo2.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
update.Enabled = True
Command6.Enabled = True
Sql = " select * from product_entry where pr_id='" + Text13.Text + "'"
Set R = C.Execute(Sql)
Text7.Text = R.Fields(0)
Text8.Text = R.Fields(1)
'Text9.Text = R.Fields(2)
Text10.Text = R.Fields(2)
Combo2.Text = R.Fields(3)
Text11.Text = R.Fields(4)
Text13.Text = ""
search.Enabled = False
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
Text13.Text = ""
Text13.SetFocus
End Sub

Private Sub update_Click()
Sql = "update product_entry set pr_id='" + Text7.Text + "',pr_name='" + Text8.Text + "',weight=" + Text10.Text + ",unit='" + Combo2.Text + "',MRP=" + Text11.Text + " where pr_id='" + Text7.Text + "'" 'or pr_name='"+ text13.text +"'
Set R = C.Execute(Sql)
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
update.Enabled = False
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
delete.Enabled = True
Command9.Enabled = True
Sql = " select * from product_entry where pr_id='" + Text20.Text + "'"
Set R = C.Execute(Sql)
Text14.Text = R.Fields(0)
Text15.Text = R.Fields(1)
'Text16.Text = R.Fields(2)
Text17.Text = R.Fields(2)
Combo3.Text = R.Fields(3)
Text18.Text = R.Fields(4)
Text20.Text = ""
Command7.Enabled = False
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
Text20.Text = ""
Text20.SetFocus
End Sub

Private Sub delete_Click()
Sql = "delete from product_entry where pr_id='" + Text14.Text + "'"
Set R = C.Execute(Sql)
MsgBox "product removed"
Adodc1.Refresh
Text14.Text = ""
Text15.Text = ""
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
delete.Enabled = False
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
save.Enabled = False
clear.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
'Text9.Enabled = False
Combo2.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
update.Enabled = False
Command6.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
'Text16.Enabled = False
Combo3.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
delete.Enabled = False
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
update.SetFocus
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
search.Enabled = True
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
'Command4.Enabled = True
If KeyAscii = 13 Then
saerch.SetFocus
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

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.Text = UCase(Text3.Text)
Text4.SetFocus
Exit Sub
End If
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter character only")
End If
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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

Private Sub Text4_LostFocus()
Text1.Text = Text2.Text + Text4.Text
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
save.Enabled = True
save.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
save.Enabled = True
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
