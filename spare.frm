VERSION 5.00
Begin VB.Form spare1 
   Caption         =   "Form2"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   LinkTopic       =   "Form2"
   ScaleHeight     =   9765
   ScaleWidth      =   18540
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   14655
      Left            =   0
      Picture         =   "spare.frx":0000
      ScaleHeight     =   14595
      ScaleWidth      =   29595
      TabIndex        =   0
      Top             =   -120
      Width           =   29655
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   10560
         TabIndex        =   18
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   10560
         TabIndex        =   9
         Top             =   1920
         Width           =   2505
      End
      Begin VB.TextBox Text4 
         Height          =   480
         Left            =   10560
         TabIndex        =   8
         Top             =   3720
         Width           =   2460
      End
      Begin VB.TextBox Text5 
         Height          =   480
         Left            =   10560
         TabIndex        =   7
         Top             =   4560
         Width           =   2460
      End
      Begin VB.TextBox Text6 
         Height          =   480
         Left            =   10560
         TabIndex        =   6
         Top             =   5280
         Width           =   2460
      End
      Begin VB.TextBox Text7 
         Height          =   480
         Left            =   10560
         TabIndex        =   5
         Top             =   6120
         Width           =   2460
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
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
         Left            =   7800
         TabIndex        =   4
         Top             =   7560
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "VIEW"
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
         TabIndex        =   3
         Top             =   7560
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000A&
         Caption         =   "CANCEL"
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
         Left            =   11520
         MaskColor       =   &H8000000B&
         TabIndex        =   2
         Top             =   7560
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "spare.frx":83DB2
         Left            =   10560
         List            =   "spare.frx":83DCE
         TabIndex        =   1
         Top             =   2880
         Width           =   2415
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
         TabIndex        =   17
         Top             =   0
         Width           =   4815
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
         TabIndex        =   16
         Top             =   1200
         Width           =   4380
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
         TabIndex        =   15
         Top             =   1920
         Width           =   4575
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
         TabIndex        =   14
         Top             =   2640
         Width           =   3540
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL NAME"
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
         TabIndex        =   12
         Top             =   3720
         Width           =   2820
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
         TabIndex        =   11
         Top             =   5280
         Width           =   2100
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
         TabIndex        =   10
         Top             =   6120
         Width           =   2100
      End
   End
End
Attribute VB_Name = "spare1"
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
str1 = " insert into spare(id,name,cc,quantity,model,colour,price)values(' " & Text1.Text & " ' ,' " & Text2.Text & " ',' " & Combo2.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Text7.Text & " ')"
con.Execute str1
MsgBox "Successfull", vbInformation
Text1.Text = " "
Text2.Text = " "
Combo2.Text = " '"
'Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
con.Close
End Sub

Private Sub Command2_Click()
sparegrid.Show
End Sub

