VERSION 5.00
Begin VB.Form registerpage 
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   15240
      TabIndex        =   18
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   15240
      TabIndex        =   17
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   15240
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   15240
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   15240
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16320
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sign in"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14640
      TabIndex        =   11
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   12480
      TabIndex        =   7
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   11895
      Left            =   -240
      Picture         =   "registerpage.frx":0000
      ScaleHeight     =   11835
      ScaleWidth      =   20670
      TabIndex        =   0
      Top             =   -480
      Width           =   20730
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTER PLEASE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1335
         Index           =   1
         Left            =   12240
         TabIndex        =   19
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "location"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   10
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "password"
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
         Left            =   12960
         TabIndex        =   9
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "user name"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12840
         TabIndex        =   6
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "customer id"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12720
         TabIndex        =   4
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTER PLEASE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Index           =   0
         Left            =   12360
         TabIndex        =   3
         Top             =   840
         Width           =   7815
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "registerpage"
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
Public Sub connect()
con.Provider = "SQLOLEDB"
str1 = "server=(local);database=project; trusted_connection=yes"
con.Open str1
End Sub

Private Sub Command1_Click()
Call connect
str1 = "insert into reg_log(customer_id,name,user_name,password,location) values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"
con.Execute str1
MsgBox "successfull", vbInformation, vbOKCancel
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
con.Close
End Sub


Private Sub Command2_Click()
signup1.Show
End Sub

