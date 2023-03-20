VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseReturn 
   BackColor       =   &H00FF8080&
   Caption         =   "Purchase Return"
   ClientHeight    =   10335
   ClientLeft      =   4605
   ClientTop       =   1740
   ClientWidth     =   14850
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
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   14850
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   12120
      TabIndex        =   37
      Top             =   2520
      Width           =   1815
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
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1575
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
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   1575
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
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton exit 
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
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   720
      TabIndex        =   14
      Top             =   4560
      Width           =   11055
      Begin VB.TextBox Text16 
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
         TabIndex        =   57
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox Text15 
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
         Left            =   9600
         TabIndex        =   55
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text14 
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
         Left            =   8400
         TabIndex        =   54
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text13 
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
         Left            =   6960
         TabIndex        =   53
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox Text12 
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
         Left            =   3600
         TabIndex        =   52
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   47
         Top             =   4440
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   45
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ListBox List7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   8400
         TabIndex        =   43
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         TabIndex        =   29
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1920
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3960
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text7 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   1920
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   3960
         TabIndex        =   23
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   5160
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   6120
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   7200
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton add 
         Caption         =   "Add"
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
         Left            =   9720
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton remove 
         Caption         =   "Remove"
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
         Left            =   9600
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text8 
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
         Left            =   7200
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text9 
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
         Left            =   8400
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   15
         Top             =   4440
         Width           =   1455
      End
      Begin VB.ListBox List8 
         Height          =   960
         Left            =   8760
         TabIndex        =   62
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   2400
         TabIndex        =   63
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Dues"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   56
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
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
         Left            =   9600
         TabIndex        =   51
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
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
         Left            =   8400
         TabIndex        =   50
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dues"
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
         Left            =   6960
         TabIndex        =   49
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Amt"
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
         Left            =   3600
         TabIndex        =   48
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Discount"
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
         Left            =   1920
         TabIndex        =   46
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label17 
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
         Left            =   3960
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
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
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         Left            =   1920
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
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
         Left            =   5160
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   7440
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   8520
         TabIndex        =   31
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Center Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   11055
      Begin VB.TextBox Text4 
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
         Left            =   3600
         TabIndex        =   44
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1815
         Left            =   7200
         TabIndex        =   1
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3201
         _Version        =   393217
         TextRTF         =   $"Purchase Return.frx":0000
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8400
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   140574721
         CurrentDate     =   44964
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Return No."
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Id"
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Left            =   1200
         TabIndex        =   10
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Left            =   7560
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Bill No."
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Reason For Return"
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
         Left            =   7440
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE RETURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   5040
      TabIndex        =   64
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      Height          =   495
      Left            =   9360
      TabIndex        =   61
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   375
      Left            =   9000
      TabIndex        =   60
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   375
      Left            =   9480
      TabIndex        =   59
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   375
      Left            =   9360
      TabIndex        =   58
      Top             =   6480
      Width           =   1455
   End
End
Attribute VB_Name = "PurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer, X As Integer, Y As Single
Dim p As Integer
Private Sub Combo1_Click()
Combo2.clear
sql = "select pr_id from purchasebill_pr where bill_no='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
RichTextBox1.Enabled = True
Combo2.Enabled = True
sql = "select discount from Purchasebill_details where bill_no='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text11.Text = r.Fields(0)
sql = "select dues from Purchasebill_details where bill_no='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text16.Text = r.Fields(0)
End Sub

Private Sub Combo2_Click()
Label23.Caption = Combo2.Text
sql = "select balance from stock where pr_id='" + Label23.Caption + "'"
Set r = c.Execute(sql)
Label24.Caption = r.Fields(0)

sql = "select pr_name from product_entry where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Combo3.Text = r.Fields(0)
sql = "select rate from purchasebill_pr where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text8.Text = r.Fields(0)
sql = "select qty from purchasebill_pr where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Label18.Caption = r.Fields(0)
sql = "select amount from purchasebill_pr where pr_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Text9.Text = r.Fields(0)
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Combo3.Enabled = True
Text9.Enabled = True
Text5.Locked = True
Text6.Locked = True
'Text7.Locked = True
Text8.Locked = True
Combo3.Locked = True
Text9.Locked = True
add.Enabled = True

Text7.SetFocus
End Sub

Private Sub add_Click()
Label25.Caption = Text7.Text
Label26.Caption = Val(Label24.Caption) - Val(Label25.Caption)

If Combo2.Text = "" Or Text7.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "Enter All Fields"
Else
List1.AddItem Combo2.Text
List2.AddItem Text5.Text
List3.AddItem Text6.Text
List4.AddItem Combo3.Text
List5.AddItem Text7.Text
List6.AddItem Text8.Text
List7.AddItem Text9.Text
List8.AddItem Val(Label26.Caption)
'Combo2.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo3.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
'Combo2.RemoveItem Combo2.ListIndex
p = Combo2.ListIndex
If p > -1 Then
Combo2.RemoveItem p
End If

End If

Dim tot As Single
For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text10.Text = tot
remove.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text10.Locked = True
Text12.Locked = True
Text13.Text = Val(Text12.Text) - Val(Text16.Text)
Text12.Text = Val(Text10.Text - Text10.Text * Text11.Text / 100) '- Val(Text2.Text)
Text13.Text = Val(Text16.Text) - Val(Text12.Text)

End Sub

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub remove_Click()
Combo2.AddItem List1.Text
'List1.RemoveItem List1.ListIndex
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
Text10.Text = ""
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text10.Text = total
Next
End Sub

Private Sub addnew_Click()
Combo1.clear
Text14.Text = ""
sql = "select count(return_no) from purchase_return"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
Text1.Enabled = True
Text1.Locked = True

sql = " select centre_id from collection_centre"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = " select centre_name from collection_centre"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = " select ph_no from collection_centre"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select bill_no from purchasebill_details"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
DTPicker1.Enabled = True
Combo1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text14.Enabled = True
Text13.Enabled = True
clear.Enabled = True
End Sub

Private Sub save_Click()
If Text11.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "All Fields Required"
Else
sql = "insert into purchase_return values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Text2.Text + "'," + Combo1.Text + ",'" + RichTextBox1.Text + "'," + Text10.Text + "," + Text14.Text + "," + Text15.Text + ")"
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "insert into purchasereturn_pr values(" + Text1.Text + ",'" + List1.List(k) + "'," + List5.List(k) + "," + List6.List(k) + "," + List7.List(k) + ")"
Set r = c.Execute(sql)
Next
sql = "update centre_dues set total_dues=" + Text15.Text + " where centre_id=" + Text2.Text + ""
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "UPDATE stock SET balance =" + List8.List(k) + " WHERE pr_id = '" + List1.List(k) + "'"
Set r = c.Execute(sql)
Next
MsgBox "Record Saved"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text16.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""

Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
RichTextBox1.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
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
add.Enabled = False
remove.Enabled = False
save.Enabled = False
clear.Enabled = False
'Command6.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
'Combo4.Enabled = False
DTPicker1.Enabled = False
RichTextBox1.Enabled = False
End If
End Sub

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
RichTextBox1.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
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
add.Enabled = False
remove.Enabled = False
save.Enabled = False
clear.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
DTPicker1.Enabled = False
RichTextBox1.Enabled = False
'Text13.Locked = True
'Text14.Locked = True
Text13.Enabled = False
Text14.Enabled = False
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
'If KeyAscii = 13 Then
'Text12.Text = Val(Text10.Text) - Val(Text11.Text)

Text12.SetFocus
'End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text16.SetFocus
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text14.SetFocus
End If
End Sub

Private Sub Text14_Change()
Text15.Text = Val(Text13.Text) - Val(Text14.Text)


End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text15.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
save.Enabled = True
If KeyAscii = 13 Then
save.SetFocus
End If

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.SetFocus
End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
add.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
If Val(Text7.Text) > Val(Label18.Caption) Then 'Or Val(Label24.Caption) < 1 Then
MsgBox "Ordered Quantity Exceeded OR Not Enough Stock"
Text7.Text = ""
Text7.SetFocus
ElseIf Val(Text7.Text) < 1 Then
MsgBox "Minimum Quantity 1"
Else
X = Val(Text7.Text)
Y = Val(Text8.Text)
Text9.Text = X * Y
End If

End Sub
