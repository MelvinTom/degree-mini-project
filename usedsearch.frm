VERSION 5.00
Begin VB.Form usedsearch 
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19245
   LinkTopic       =   "Form1"
   Picture         =   "usedsearch.frx":0000
   ScaleHeight     =   11055
   ScaleWidth      =   19245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   10920
      TabIndex        =   30
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   10920
      TabIndex        =   29
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   10800
      TabIndex        =   28
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   10800
      TabIndex        =   27
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   405
      Left            =   10800
      TabIndex        =   26
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   10800
      TabIndex        =   25
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   10800
      TabIndex        =   24
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   10800
      TabIndex        =   23
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   10800
      TabIndex        =   22
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3000
      TabIndex        =   21
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3000
      TabIndex        =   19
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3000
      TabIndex        =   17
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   16
      Top             =   7440
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "varient (cc)"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   8520
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "body color"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   8640
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "trasmission"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "no.of gears"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "clutch type"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   4320
      Width           =   1935
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
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   5160
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
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   5760
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
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   6480
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
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "year"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "manufacturer"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "rc book no:"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "insurance exp:"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "pollution"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   5400
      Width           =   2415
   End
End
Attribute VB_Name = "usedsearch"
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
str1 = "select * from used where rc = '" & Combo1.Text & "'"

rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text1.Text = rs.Fields(1)
Text2.Text = rs.Fields(2)
Text3.Text = rs.Fields(3)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(10)
Text10.Text = rs.Fields(11)
Text11.Text = rs.Fields(12)
Text12.Text = rs.Fields(13)
Text14.Text = rs.Fields(15)
Text13.Text = rs.Fields(14)





rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub Form_Load()
Call connect
str1 = "select rc from used"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close


End Sub

