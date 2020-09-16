VERSION 5.00
Begin VB.Form service1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20025
   LinkTopic       =   "Form1"
   Picture         =   "service.frx":0000
   ScaleHeight     =   11055
   ScaleWidth      =   20025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "view"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      TabIndex        =   19
      Top             =   9240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "back"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   18
      Top             =   9240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000008&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   16
      Top             =   9240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   15
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   1335
      Left            =   8040
      TabIndex        =   14
      Top             =   7800
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   12960
      TabIndex        =   13
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   12960
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   12960
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "service"
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
      Height          =   855
      Left            =   8040
      TabIndex        =   17
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
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
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NEXT SERVICE km"
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
      Height          =   735
      Left            =   10560
      TabIndex        =   10
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "LAST SERVICE km"
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
      Height          =   735
      Left            =   10560
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPLAINTS:"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION NO"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "MODEL"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICE"
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
      Height          =   735
      Left            =   7920
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "service1"
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
str1 = "insert into service1(name,model,regno,amount,lastkm,nextkm,complaint)values(' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Text7.Text & " ') "
con.Execute str1
MsgBox " successfull", vbInformation
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
con.Close
End Sub

Private Sub Command2_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
End Sub

Private Sub Command3_Click()
MDIFORM1.Show
End Sub

Private Sub Command4_Click()
servicegrid.Show
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
'If Not (IsNumeric(Text3.Text)) Then
'MsgBox " please enter valid registration number", vbExclamation, "service1"
'End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Not (IsNumeric(Text4.Text)) Then
MsgBox " please enter valid amount", vbExclamation, "service1"
End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
If Not (IsNumeric(Text5.Text)) Then
MsgBox " please enter last service km", vbExclamation, "service1"
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Not (IsNumeric(Text6.Text)) Then
MsgBox " please enter next service km", vbExclamation, "service1"
End If
End Sub
