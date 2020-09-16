VERSION 5.00
Begin VB.Form used1 
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17400
   LinkTopic       =   "Form1"
   Picture         =   "used.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   17400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   37
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
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
      Height          =   855
      Left            =   13680
      TabIndex        =   35
      Top             =   9360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "previous page"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      TabIndex        =   34
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "view"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   33
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   32
      Top             =   9360
      Width           =   2775
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   3600
      TabIndex        =   31
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   3600
      TabIndex        =   30
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   3600
      TabIndex        =   29
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   16560
      TabIndex        =   25
      Top             =   8400
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   16560
      TabIndex        =   24
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   16560
      TabIndex        =   23
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   16560
      TabIndex        =   22
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   16920
      TabIndex        =   17
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   16920
      TabIndex        =   16
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   16920
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   16920
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   16920
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   36
      Top             =   1920
      Width           =   1455
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
      Left            =   1200
      TabIndex        =   28
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "insurance type"
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
      TabIndex        =   27
      Top             =   6360
      Width           =   2415
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
      TabIndex        =   26
      Top             =   5640
      Width           =   2415
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
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   14280
      TabIndex        =   21
      Top             =   8400
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
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   14280
      TabIndex        =   20
      Top             =   7680
      Width           =   2775
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
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   14280
      TabIndex        =   19
      Top             =   6960
      Width           =   2175
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
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   14280
      TabIndex        =   18
      Top             =   6240
      Width           =   2175
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14160
      TabIndex        =   14
      Top             =   4800
      Width           =   1935
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14160
      TabIndex        =   13
      Top             =   4200
      Width           =   3255
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14160
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
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
      Left            =   1680
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
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
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   14160
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
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
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   14160
      TabIndex        =   4
      Top             =   3000
      Width           =   2175
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label s 
      BackStyle       =   0  'Transparent
      Caption         =   "used bikes"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   9000
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "used bikes"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   8280
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "used1"
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
If Text16.Text = "" Then
MsgBox "id is mandatory", vbInformation
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "year is mandatory", vbInformation
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "name is mandatory", vbInformation
Exit Sub
End If

If Text3.Text = "" Then
MsgBox "manufacturer is mandatory", vbInformation
Exit Sub
End If
If Text13.Text = "" Then
MsgBox "rc book number is mandatory", vbInformation
Exit Sub
End If
If Text14.Text = "" Then
MsgBox "insurance expiry is mandatory", vbInformation
Exit Sub
End If
If Text15.Text = "" Then
MsgBox "pollution is mandatory", vbInformation
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "body colour is mandatory", vbInformation
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "variant is mandatory", vbInformation
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "transmission is mandatory", vbInformation
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "number of gearsis mandatory", vbInformation
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "clutch is mandatory", vbInformation
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "length is mandatory", vbInformation
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "height is mandatory", vbInformation
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "tyre front size is mandatory", vbInformation
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "tyre rear size is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = "insert into used(id,years,name,man,rc,ins,pol,body,variant,trans,nos,clutch,leng,height,front,rear)values(' " & Text16.Text & " ',' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text13.Text & " ',' " & Text14.Text & " ',' " & Text15.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Text7.Text & " ',' " & Text8.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ',' " & Text11.Text & " ',' " & Text12.Text & " ')"
con.Execute str1
con.Close
MsgBox " successfull", vbInformation
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
Text13.Text = " "
Text14.Text = " "
Text15.Text = " "
Text16.Text = " "

End Sub

Private Sub Command2_Click()
usedgrid.Show
End Sub

Private Sub Command3_Click()
MDIFORM1.Show
End Sub

Private Sub Command4_Click()
usedsearch.Show

End Sub

Private Sub Form_Load()
Dim n As Integer
If con.State = adStateOpen Then
rs.Close
con.Close
End If
Text16.Enabled = False
Call connect
str1 = "select * from used"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text16.Text = "1"
Else
rs.MoveLast
n = rs("id").Value
Text16.Text = n + 1
End If
rs.Close
con.Close
End Sub

'Private Sub Text1_Validate(Cancel As Boolean)
'If Not (IsNumeric(Text1.Text)) Then
'MsgBox " please enter vaild year", vbExclamation, "used1"
'End If
'End Sub

Private Sub Text10_Validate(Cancel As Boolean)
If Not (IsNumeric(Text10.Text)) Then
MsgBox " please enter breadth numerically", vbExclamation, "used1"
End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
If Not (IsNumeric(Text11.Text)) Then
MsgBox " please enter front tyre size numerically", vbExclamation, "used1"
End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Not (IsNumeric(Text12.Text)) Then
MsgBox " please enter rear tyre size numerically", vbExclamation, "used1"
End If
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
If Not (IsNumeric(Text13.Text)) Then
MsgBox " please enter valid rc book", vbExclamation, "used1"
End If
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
If Not (IsNumeric(Text16.Text)) Then
MsgBox " please enter valid id", vbExclamation, "used1"
End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If Not (IsNumeric(Text7.Text)) Then
MsgBox " please enter number of gears", vbExclamation, "used1"
End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Not (IsNumeric(Text13.Text)) Then
MsgBox " please enter length numerically", vbExclamation, "used1"
End If
End Sub
