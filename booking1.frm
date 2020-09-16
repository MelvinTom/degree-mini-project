VERSION 5.00
Begin VB.Form booking1 
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   Picture         =   "booking1.frx":0000
   ScaleHeight     =   9600
   ScaleWidth      =   19035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   5160
      TabIndex        =   34
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      Height          =   405
      Left            =   5160
      TabIndex        =   33
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "back"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10560
      TabIndex        =   31
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8280
      TabIndex        =   30
      Top             =   7080
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "booking1.frx":36A5B
      Left            =   12480
      List            =   "booking1.frx":36A68
      TabIndex        =   29
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "booking1.frx":36A8D
      Left            =   12480
      List            =   "booking1.frx":36A97
      TabIndex        =   28
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "book"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6000
      TabIndex        =   27
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   12480
      TabIndex        =   26
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   12480
      TabIndex        =   20
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   12480
      TabIndex        =   19
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   12480
      TabIndex        =   18
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   5160
      TabIndex        =   16
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "booking id"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   32
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "phone number"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "advance"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   23
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "extra accessories"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "type of payment"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   21
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "pin code"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "state"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "tyre"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "colour"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "varient"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "model"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "district"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "place"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "house name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "booking form"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "booking1"
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
If Text13.Text = "" Then
MsgBox "booking id is mandatory", vbInformation
Exit Sub
End If
If Text14.Text = "" Then
MsgBox "name is mandatory", vbInformation
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "house name is mandatory", vbInformation
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "place mandatory", vbInformation
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "district is mandatory", vbInformation
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "state is mandatory", vbInformation
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "pin number is mandatory", vbInformation
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "phone is mandatory", vbInformation
Exit Sub
End If
If Text12.Text = "" Then
MsgBox " advance is mandatory", vbInformation
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "model is mandatory", vbInformation
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "variant is mandatory", vbInformation
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "colour is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = "insert into booking(bid,name,hname,place,district,states,pin,mob,payment,extra,advance,model,variant,colour,tyre)values(' " & Text13.Text & " ',' " & Text14.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Text11.Text & " ',' " & Combo1.Text & " ',' " & Combo2.Text & " ',' " & Text12.Text & " ',' " & Text7.Text & " ',' " & Text8.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ');"
con.Execute str1
MsgBox " billing succesfull", vbInformation
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
Combo1.Text = " "
Combo2.Text = " "
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
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
Combo1.Text = " "
Combo2.Text = " "
End Sub

Private Sub Command3_Click()
MDIFORM1.Show
End Sub

Private Sub Form_Load()
Dim n As Integer
If con.State = adStateOpen Then
rs.Close
con.Close
End If
Text13.Enabled = False
Call connect
str1 = "select * from booking"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text13.Text = "1"
Else
rs.MoveLast
n = rs("bid").Value
Text13.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
If Not (IsNumeric(Text10.Text)) Then
MsgBox " please enter vaild tyre size", vbExclamation, "booking1"
End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
If Len(Text11.Text) < 10 Or Len(Text11.Text) > 10 Then
MsgBox " please enter 10 digit mobile number", vbExclamation, "booking1"
End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Not (IsNumeric(Text12.Text)) Then
MsgBox " please enter valid amount", vbExclamation, "booking1"
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Len(Text6.Text) < 6 Or Len(Text6.Text) > 6 Then
MsgBox " enter vaild pin", vbExclamation, "booking1"
End If
End Sub
