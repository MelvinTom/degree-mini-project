VERSION 5.00
Begin VB.Form sparesearch2 
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   18900
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   13680
      TabIndex        =   17
      Top             =   1200
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   14655
      Left            =   -360
      Picture         =   "sparesearch2.frx":0000
      ScaleHeight     =   14595
      ScaleWidth      =   29595
      TabIndex        =   0
      Top             =   0
      Width           =   29655
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   10560
         TabIndex        =   8
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "search"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9360
         TabIndex        =   7
         Top             =   7560
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   480
         Left            =   10560
         TabIndex        =   6
         Top             =   6120
         Width           =   2460
      End
      Begin VB.TextBox Text6 
         Height          =   480
         Left            =   10560
         TabIndex        =   5
         Top             =   5280
         Width           =   2460
      End
      Begin VB.TextBox Text5 
         Height          =   480
         Left            =   10560
         TabIndex        =   4
         Top             =   4560
         Width           =   2460
      End
      Begin VB.TextBox Text4 
         Height          =   480
         Left            =   10560
         TabIndex        =   3
         Top             =   3720
         Width           =   2460
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   10560
         TabIndex        =   2
         Top             =   1920
         Width           =   2505
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   10560
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   600
         Left            =   7560
         TabIndex        =   16
         Top             =   6120
         Width           =   2100
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COLOUR"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   600
         Left            =   7560
         TabIndex        =   15
         Top             =   5280
         Width           =   2100
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL No"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   600
         Left            =   7320
         TabIndex        =   14
         Top             =   3720
         Width           =   2820
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   600
         Left            =   7680
         TabIndex        =   13
         Top             =   4560
         Width           =   2100
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Centimeter (CC)"
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
         Height          =   720
         Left            =   6840
         TabIndex        =   12
         Top             =   2640
         Width           =   3540
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SPARE PART NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   615
         Left            =   6240
         TabIndex        =   11
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SPARE PART ID"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   600
         Left            =   6360
         TabIndex        =   10
         Top             =   1200
         Width           =   4380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "ADD spare"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   7680
         TabIndex        =   9
         Top             =   0
         Width           =   4815
      End
   End
End
Attribute VB_Name = "sparesearch2"
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
str1 = "select * from spare where name = '" & Combo1.Text & "'"

rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text1.Text = rs.Fields(0)


Text2.Text = rs.Fields(1)



'Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)

Text7.Text = rs.Fields(6)
Combo2.Text = rs.Fields(2)
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If



Call connect
str1 = "select name from spare"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close

End Sub

