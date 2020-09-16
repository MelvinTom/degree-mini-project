VERSION 5.00
Begin VB.Form stocksearch 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   17430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Index           =   0
      Left            =   -2760
      Picture         =   "stockserch.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   20115
      TabIndex        =   0
      Top             =   -120
      Width           =   20175
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   10800
         TabIndex        =   31
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "search"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   30
         Top             =   6840
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000080FF&
         Height          =   3495
         Index           =   1
         Left            =   3000
         ScaleHeight     =   3435
         ScaleWidth      =   4875
         TabIndex        =   16
         Top             =   1440
         Width           =   4935
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "stockserch.frx":36A5B
            Left            =   1800
            List            =   "stockserch.frx":36A65
            TabIndex        =   22
            Top             =   2400
            Width           =   2055
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Top             =   1920
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1800
            TabIndex        =   20
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1800
            TabIndex        =   19
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "fuel  trasmission"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "display"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "tank capacity"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   26
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            Caption         =   "body color"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   25
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackColor       =   &H000080FF&
            Caption         =   "varient"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H000080FF&
            Caption         =   "vehicle name"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   2775
         Left            =   8280
         Picture         =   "stockserch.frx":36A7F
         ScaleHeight     =   2715
         ScaleWidth      =   4875
         TabIndex        =   7
         Top             =   1440
         Width           =   4935
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   0
            Width           =   2295
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   2280
            TabIndex        =   8
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "length"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "height"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "tire size front"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            TabIndex        =   13
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "tire size rear"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            TabIndex        =   12
            Top             =   2160
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000080FF&
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   3000
         ScaleHeight     =   915
         ScaleWidth      =   4875
         TabIndex        =   1
         Top             =   4920
         Width           =   4935
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "stockserch.frx":38492
            Left            =   1800
            List            =   "stockserch.frx":3849C
            TabIndex        =   3
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "clutch type"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "no. of gears"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Label Label3 
         Caption         =   "model"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   33
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "vehicle name"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "stocksearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.recordset
Public cmd As New ADODB.Command
Public str As String
Public str1 As String
Public str2 As String
Public str3 As String
Public Sub connect()
con.Provider = "sqloledb"
str1 = "server=(local);database=project;trusted_connection=Yes"
con.Open str1
End Sub

Private Sub Command1_Click()
Call connect
str1 = "select * from stock1 where model = '" & Combo3.Text & "'"


rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text1.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Combo2.Text = rs.Fields(8)
Combo1.Text = rs.Fields(6)
Text7.Text = rs.Fields(7)
Text9.Text = rs.Fields(9)
Text10.Text = rs.Fields(10)
Text11.Text = rs.Fields(11)
Text12.Text = rs.Fields(12)


rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Form_Load()
Call connect
str1 = "select model from stock1"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo3.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

