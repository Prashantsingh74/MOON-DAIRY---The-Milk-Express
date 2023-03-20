VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Home"
   ClientHeight    =   6345
   ClientLeft      =   6855
   ClientTop       =   3720
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6345
      Left            =   0
      ScaleHeight     =   6285
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   0
      Width           =   9405
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALE BILL"
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   3840
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MILK COLLECTION"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9240
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Image Image1 
         Height          =   1725
         Left            =   3240
         Picture         =   "Home.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2160
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "The Milk Express"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME TO MOON DAIRY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Image Image2 
         Height          =   10500
         Left            =   0
         Picture         =   "Home.frx":461B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   18000
      End
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label3_Click()
MilkCollection.Show
End Sub

Private Sub Label4_Click()
Salebill.Show
End Sub
