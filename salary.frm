VERSION 5.00
Begin VB.Form salary1 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "view"
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
      Left            =   9840
      TabIndex        =   21
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   12120
      TabIndex        =   16
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
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
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "save"
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
      Left            =   9840
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   9
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3960
      TabIndex        =   8
      Top             =   9240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   11055
      Left            =   0
      Picture         =   "salary.frx":0000
      ScaleHeight     =   10995
      ScaleWidth      =   21030
      TabIndex        =   0
      Top             =   0
      Width           =   21090
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   4800
         TabIndex        =   26
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         Caption         =   "search"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9840
         TabIndex        =   24
         Top             =   6720
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "total salary"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   20
         Top             =   4080
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "salary.frx":1B88E
         Left            =   4800
         List            =   "salary.frx":1B89B
         TabIndex        =   19
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5880
         TabIndex        =   18
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   12120
         TabIndex        =   17
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   12120
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   12120
         TabIndex        =   14
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "salary id"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   25
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "calculate total salary:"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   3240
         TabIndex        =   22
         Top             =   3480
         Width           =   5775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "da"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   10200
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "BASIC PAY"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   9600
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "salary DETAILS"
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
         Left            =   6120
         TabIndex        =   7
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "HRA"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   10080
         TabIndex        =   4
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "employee NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   9000
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "department"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "employee id"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   1920
         Width           =   2775
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   9480
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "salary1"
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
con.Provider = "sqloledb"
str1 = "server=(local);database=project; trusted_connection=yes"
con.Open str1
End Sub

Private Sub Combo2_Click()


End Sub

Private Sub Combo1_Change()
Dim str1 As String
Call connect
str1 = "select * from emp1"
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
'MsgBox (rs.Fields(1))
Combo1.AddItem (rs.Fields(1))
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Command3_Click()
Call connect
str1 = " insert into salary(sid,eid,department,ename,da,hra,pf,total)values(' " & Text5.Text & " ',' " & Combo1.Text & " ' ,' " & Combo3.Text & " ',' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text6.Text & " ')"
con.Execute str1
con.Close
MsgBox "save successfull", vbInformation, vbOKCancel

Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Combo1.Text = " "
Combo3.Text = " "

End Sub


Private Sub Command4_Click()
MDIFORM1.Show
End Sub

Private Sub Command5_Click()
Text6.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text)
End Sub

Private Sub Command6_Click()
salarygrid.Show
End Sub

Private Sub Command8_Click()
Call connect
str1 = "select * from emp1 where id=" & Combo1.Text & ""
con.Execute str1
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(7)
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Form_Load()
Dim n As Integer
If con.State = adStateOpen Then
rs.Close
con.Close
End If
Text2.Enabled = False
Call connect
str1 = "select * from salary"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text5.Text = "1"
Else
rs.MoveLast
n = rs("sid").Value
Text5.Text = n + 1
End If
rs.Close
con.Close
Call connect
str1 = "select * from emp1"
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
'MsgBox (rs.Fields(1))
Combo1.AddItem (rs.Fields(1))
rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub Command7_Click()
Call connect
str1 = "delete from salary where eid= '" & Combo1.Text & "' "
con.Execute str1
MsgBox "successfully deleted ", vbInformation
con.Close
End Sub

