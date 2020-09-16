VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form register 
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   Picture         =   "homepage.frx":0000
   ScaleHeight     =   10620
   ScaleWidth      =   16695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
      Caption         =   "SIGN IN"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      ForeColor       =   &H80000001&
      Height          =   10935
      Left            =   0
      Picture         =   "homepage.frx":BC79
      ScaleHeight     =   10875
      ScaleWidth      =   16635
      TabIndex        =   0
      Top             =   -315
      Width           =   16695
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   12675
         Left            =   -705
         TabIndex        =   7
         Top             =   270
         Width           =   22005
         URL             =   "D:\MINI PROJECT\PROJECT DESIGN AND CODE\(HDvd9.co)_KTM-Duke-250--Official-Video.mp4"
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   38814
         _cy             =   22357
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "READY TO RACE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Index           =   0
         Left            =   8280
         TabIndex        =   5
         Top             =   1800
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ktm"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1455
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   8655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE REGISTER OR SIGN IN"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   3120
         Width           =   8535
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   9480
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   9480
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7800
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
registerpage.Show
End Sub

Private Sub Command2_Click()
signup1.Show
End Sub


