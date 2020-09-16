VERSION 5.00
Begin VB.Form sales1 
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20025
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   20025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   15840
      TabIndex        =   37
      Top             =   8880
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11895
      Left            =   0
      Picture         =   "sales.frx":0000
      ScaleHeight     =   11835
      ScaleWidth      =   22005
      TabIndex        =   0
      Top             =   -600
      Width           =   22065
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
         Height          =   615
         Left            =   13920
         TabIndex        =   23
         Top             =   9480
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
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
         Height          =   615
         Left            =   12000
         TabIndex        =   22
         Top             =   9480
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14520
         TabIndex        =   21
         Top             =   8400
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   14520
         TabIndex        =   20
         Top             =   7680
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         TabIndex        =   18
         Top             =   6120
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         TabIndex        =   17
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         TabIndex        =   15
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         TabIndex        =   14
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         TabIndex        =   13
         Top             =   3600
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14520
         TabIndex        =   11
         Top             =   7200
         Width           =   2535
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14520
         TabIndex        =   9
         Top             =   6720
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15960
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "sales.frx":454AA
         Left            =   15120
         List            =   "sales.frx":454AC
         TabIndex        =   2
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "sales.frx":454AE
         Left            =   14520
         List            =   "sales.frx":454B0
         TabIndex        =   1
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE OF PAYMENT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   11520
         TabIndex        =   35
         Top             =   8400
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   11520
         TabIndex        =   31
         Top             =   6240
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PIN CODE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   1
         Left            =   12120
         TabIndex        =   30
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   12480
         TabIndex        =   24
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   1
         Left            =   12480
         TabIndex        =   26
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "sales details"
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
         Left            =   12240
         TabIndex        =   25
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   12120
         TabIndex        =   19
         Top             =   7800
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PIN CODE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Index           =   0
         Left            =   12120
         TabIndex        =   16
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CITY"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   12480
         TabIndex        =   12
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "CC"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   12480
         TabIndex        =   10
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H8000000E&
         Height          =   495
         Index           =   0
         Left            =   12120
         TabIndex        =   8
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   0
         Left            =   11520
         TabIndex        =   7
         Top             =   6240
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "HOUSE NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   11880
         TabIndex        =   6
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   11640
         TabIndex        =   5
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "SALES DETAILS"
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
         Height          =   975
         Left            =   12240
         TabIndex        =   4
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   1
         Left            =   11640
         TabIndex        =   27
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "HOUSE NAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   11880
         TabIndex        =   28
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CITY"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   12480
         TabIndex        =   29
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   2
         Left            =   12120
         TabIndex        =   32
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "CC"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   1
         Left            =   12480
         TabIndex        =   33
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   1
         Left            =   12120
         TabIndex        =   34
         Top             =   7800
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE OF PAYMENT"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   1
         Left            =   11520
         TabIndex        =   36
         Top             =   8400
         Width           =   2535
      End
   End
End
Attribute VB_Name = "sales1"
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

Private Sub Combo1_Validate(Cancel As Boolean)
If Not (IsNumeric(Combo1.Text)) Then
MsgBox " please enter day", vbExclamation, "sales1"
End If
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
If Not (IsNumeric(Combo2.Text)) Then
MsgBox " please enter month", vbExclamation, "sales1"
End If
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
If Not (IsNumeric(Combo3.Text)) Then
MsgBox " please enter year", vbExclamation, "sales1"
End If
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox "day is mandatory", vbInformation
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox "month is mandatory", vbInformation
Exit Sub
End If
If Combo3.Text = "" Then
MsgBox "year is mandatory", vbInformation
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "name is mandatory", vbInformation
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "house name is mandatory", vbInformation
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "city is mandatory", vbInformation
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "pincode mandatory", vbInformation
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "phone number is mandatory", vbInformation
Exit Sub
End If
If Combo4.Text = "" Then
MsgBox "model is mandatory", vbInformation
Exit Sub
End If
If Combo5.Text = "" Then
MsgBox "cc is mandatory", vbInformation
Exit Sub
End If

If Text6.Text = "" Then
MsgBox "amount is mandatory", vbInformation
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "type of payment is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = "insert into sales(day,date,year,cname,hname,city,pin,mob,model,cc,amount,payment) values(' " & Combo1.Text & " ',' " & Combo2.Text & " ',' " & Combo3.Text & " ',' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Combo4.Text & " ',' " & Combo5.Text & " ',' " & Text6.Text & " ',' " & Text7.Text & " ');"
con.Execute str1
MsgBox "succesfull", vbInformation
Combo1.Text = " "
Combo2.Text = " "
Combo3.Text = " "
Combo4.Text = " "
Combo5.Text = " "
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
Combo1.Text = " "
Combo2.Text = " "
Combo3.Text = " "
Combo4.Text = " "
Combo5.Text = " "
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
salesgrid.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem ("1")
Combo1.AddItem ("2")
Combo1.AddItem ("3")
Combo1.AddItem ("4")
Combo1.AddItem ("5")
Combo1.AddItem ("6")
Combo1.AddItem ("7")
Combo1.AddItem ("8")
Combo1.AddItem ("9")
Combo1.AddItem ("10")
Combo1.AddItem ("11")
Combo1.AddItem ("12")
Combo1.AddItem ("13")
Combo1.AddItem ("14")
Combo1.AddItem ("15")
Combo1.AddItem ("16")
Combo1.AddItem ("17")
Combo1.AddItem ("18")
Combo1.AddItem ("19")
Combo1.AddItem ("20")
Combo1.AddItem ("21")
Combo1.AddItem ("22")
Combo1.AddItem ("23")
Combo1.AddItem ("24")
Combo1.AddItem ("25")
Combo1.AddItem ("26")
Combo1.AddItem ("27")
Combo1.AddItem ("28")
Combo1.AddItem ("29")
Combo1.AddItem ("30")
Combo1.AddItem ("31")
Combo2.AddItem ("01")
Combo2.AddItem ("02")
Combo2.AddItem ("03")
Combo2.AddItem ("04")
Combo2.AddItem ("05")
Combo2.AddItem ("06")
Combo2.AddItem ("07")
Combo2.AddItem ("08")
Combo2.AddItem ("09")
Combo2.AddItem ("10")
Combo2.AddItem ("11")
Combo2.AddItem ("12")
Combo3.AddItem ("2014")
Combo3.AddItem ("2015")
Combo3.AddItem ("2016")
Combo3.AddItem ("2017")
Combo3.AddItem ("2018")
Combo3.AddItem ("2019")
Combo3.AddItem ("2020")
Combo3.AddItem ("2021")
Combo3.AddItem ("2022")
Combo3.AddItem ("2023")
Combo4.AddItem ("RC")
Combo4.AddItem ("DUKE")
Combo5.AddItem ("125")
Combo5.AddItem ("200")
Combo5.AddItem ("250")
Combo5.AddItem ("390")
Combo5.AddItem ("790")
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Not (IsNumeric(Text4.Text)) Then
MsgBox " please enter vaild pin ", vbExclamation, "sales1"
End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
If Not (IsNumeric(Text5.Text)) Then
MsgBox " please enter mobile number numerically", vbExclamation, "billing1"
End If
If Len(Text5.Text) < 10 Or Len(Text5.Text) > 10 Then
MsgBox " enter phone number in 10 digits", vbExclamation, "sales1"
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Not (IsNumeric(Text6.Text)) Then
MsgBox " please enter valid amount", vbExclamation, "sales1"
End If
End Sub
