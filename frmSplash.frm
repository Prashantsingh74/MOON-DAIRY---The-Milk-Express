VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frm Splash"
   ClientHeight    =   5640
   ClientLeft      =   7185
   ClientTop       =   3045
   ClientWidth     =   10635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10545
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8640
         Top             =   3360
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   4560
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Left            =   4080
         TabIndex        =   5
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BY   PRJ-2231D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6600
         TabIndex        =   3
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Image imgLogo 
         Height          =   2745
         Left            =   240
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2655
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Milk Express"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3840
         TabIndex        =   2
         Top             =   1560
         Width           =   4125
      End
      Begin VB.Label MOONDAIRY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOON DAIRY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4680
         TabIndex        =   1
         Top             =   840
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame1_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 4
Label2.Caption = "Loading.... " & ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
home.Show
End If
End Sub
