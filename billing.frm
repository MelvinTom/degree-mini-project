VERSION 5.00
Begin VB.Form billing1 
   Caption         =   "Form1"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11760
      Left            =   0
      Picture         =   "billing.frx":0000
      ScaleHeight     =   11700
      ScaleWidth      =   21630
      TabIndex        =   0
      Top             =   -600
      Width           =   21690
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
         Height          =   495
         Left            =   6120
         TabIndex        =   38
         Top             =   9120
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "clear"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   37
         Top             =   9120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "bill"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   36
         Top             =   9120
         Width           =   1695
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   8760
         TabIndex        =   35
         Top             =   7440
         Width           =   2055
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   8760
         TabIndex        =   33
         Top             =   6960
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   8760
         TabIndex        =   31
         Top             =   6480
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   8760
         TabIndex        =   29
         Top             =   6000
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   7560
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   6600
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   6120
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3240
         TabIndex        =   11
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "total amount"
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
         Height          =   495
         Left            =   6000
         TabIndex        =   34
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "accessories"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   6240
         TabIndex        =   32
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "insurance"
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
         Left            =   6240
         TabIndex        =   30
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "ex.showroom price"
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
         Left            =   5400
         TabIndex        =   28
         Top             =   6000
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "total amount"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   6600
         TabIndex        =   27
         Top             =   5400
         Width           =   5175
      End
      Begin VB.Label Label14 
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
         Left            =   1560
         TabIndex        =   22
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "variant"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "cc"
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
         Left            =   1800
         TabIndex        =   20
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "bike model"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "bike details"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Left            =   1680
         TabIndex        =   18
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   2520
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "email"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "mobile number"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label7 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "town"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "item id"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "last name"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "first name"
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "personal details"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5535
      End
   End
End
Attribute VB_Name = "billing1"
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

If Text1.Text = "" Then
MsgBox "first name is mandatory", vbInformation
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "last name is mandatory", vbInformation
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "item id is mandatory", vbInformation
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "house name mandatory", vbInformation
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "town is mandatory", vbInformation
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "pin code is mandatory", vbInformation
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "phone number is mandatory", vbInformation
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "Email is mandatory", vbInformation
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "bike model is mandatory", vbInformation
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "variant is mandatory", vbInformation
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "cc is mandatory", vbInformation
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "colour is mandatory", vbInformation
Exit Sub
End If
If Text13.Text = "" Then
MsgBox "ex showroom price is mandatory", vbInformation
Exit Sub
End If
If Text14.Text = "" Then
MsgBox "insurance is mandatory", vbInformation
Exit Sub
End If
If Text16.Text = "" Then
MsgBox "accessories is mandatory", vbInformation
Exit Sub
End If
If Text17.Text = "" Then
MsgBox "total is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = "insert into billing1( fname,lname,id,hname,town,pin,mob,email,model,cc,variant,colour,exprce,insurance,accessories,total) values(' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Text7.Text & " ',' " & Text8.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ',' " & Text11.Text & " ',' " & Text12.Text & " ',' " & Text13.Text & " ',' " & Text14.Text & " ',' " & Text16.Text & " ',' " & Text17.Text & " ')"
con.Execute str1
Text17.Text = Val(Text13.Text) + Val(Text14.Text) + Val(Text16.Text)

MsgBox " succesfully saved", vbInformation

Text17.Text = Val(Text13.Text) + Val(Text14.Text) + Val(Text16.Text)
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
Text13.Text = " "
Text14.Text = " "
Text16.Text = " "
Text17.Text = " "
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
Text3.Enabled = False
Call connect
str1 = "select * from billing1"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text3.Text = "1"
Else
rs.MoveLast
n = rs("id").Value
Text3.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Not (IsNumeric(Text6.Text)) Then
MsgBox " please enter pin number numerically", vbExclamation, "billing1"
End If
If Len(Text6.Text) < 6 Or Len(Text6.Text) > 6 Then
MsgBox " enter vaild pin", vbExclamation, "billing"
End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If Not (IsNumeric(Text7.Text)) Then
MsgBox " please enter mobile number numerically", vbExclamation, "billing1"
End If
If Len(Text7.Text) < 10 Or Len(Text7.Text) > 10 Then
MsgBox " enter phone number in 10 digits", vbExclamation, "billing"
End If



End Sub

