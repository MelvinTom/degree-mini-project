VERSION 5.00
Begin VB.Form register 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   Picture         =   "register.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   16695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
      Caption         =   "SIGN IN"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
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
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      ForeColor       =   &H80000001&
      Height          =   10935
      Left            =   0
      Picture         =   "register.frx":BC79
      ScaleHeight     =   10875
      ScaleWidth      =   16635
      TabIndex        =   0
      Top             =   -3435
      Width           =   16695
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "READY TO RACE"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9600
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE REGISTER OR SIGN IN"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4440
         TabIndex        =   4
         Top             =   3480
         Width           =   8535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "welcome"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4800
         TabIndex        =   3
         Top             =   0
         Width           =   18015
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   9480
      TabIndex        =   5
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

End Sub

Private Sub Command2_Click()
signup.Show
End Sub

