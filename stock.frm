VERSION 5.00
Begin VB.Form stock1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16860
   LinkTopic       =   "Form2"
   Picture         =   "stock.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   16860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   2775
      Left            =   13680
      Picture         =   "stock.frx":34303
      ScaleHeight     =   2715
      ScaleWidth      =   4635
      TabIndex        =   27
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2280
         TabIndex        =   31
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2280
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   2160
         Width           =   2295
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
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   35
         Top             =   0
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
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   720
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
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   33
         Top             =   1440
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
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   32
         Top             =   2160
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREVIOUS PAGE"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   22
      Top             =   7440
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
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
      Left            =   6480
      TabIndex        =   21
      Top             =   7440
      Width           =   2175
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
      Left            =   4080
      TabIndex        =   20
      Top             =   7440
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   2160
      ScaleHeight     =   915
      ScaleWidth      =   4035
      TabIndex        =   12
      Top             =   5760
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "stock.frx":35D16
         Left            =   1800
         List            =   "stock.frx":35D20
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "clutch type"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "no. of gears"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   3015
      Left            =   720
      ScaleHeight     =   2955
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "stock.frx":35D35
         Left            =   1800
         List            =   "stock.frx":35D3F
         TabIndex        =   26
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "fuel  trasmission"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "display"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "tank capacity"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         Caption         =   "body color"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "varient"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "model"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "vehile id"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "main details"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "transmission"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   17
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "add stock"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "stock1"
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
MsgBox "vehicle id is mandatory", vbInformation
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "model is mandatory", vbInformation
Exit Sub
End If
If Combo2.Text = "" Then
MsgBox " fuel transmission is mandatory", vbInformation
Exit Sub
End If

If Text3.Text = "" Then
MsgBox "variant is mandatory", vbInformation
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "colour is mandatory", vbInformation
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "tank capacity is mandatory", vbInformation
Exit Sub
End If
If Text6.Text = "" Then
MsgBox "display is mandatory", vbInformation
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "number of gears is mandatory", vbInformation
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox " clutch type is mandatory", vbInformation
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "length is mandatory", vbInformation
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "width is mandatory", vbInformation
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "tyre size is mandatory", vbInformation
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "tyre size is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = "insert into stock1(id,model,variant,colour,tank,display,fuel,gear,clutch,length,breadth,tiref,tirer)values(' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text3.Text & " ',' " & Text4.Text & " ',' " & Text5.Text & " ',' " & Text6.Text & " ',' " & Combo1.Text & " ',' " & Text7.Text & " ',' " & Combo2.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ',' " & Text11.Text & " ',' " & Text12.Text & " ')"
con.Execute str1
MsgBox "succesfully saved", vbInformation
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
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
Text1.Enabled = False
Call connect
str1 = "select * from stock1"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Not (IsNumeric(Text1.Text)) Then
MsgBox " please enter vaild id", vbExclamation, "stock1"
End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
If Not (IsNumeric(Text10.Text)) Then
MsgBox " please enter vaild height", vbExclamation, "stock1"
End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
If Not (IsNumeric(Text11.Text)) Then
MsgBox " please enter vaild tyre front size", vbExclamation, "stock1"
End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Not (IsNumeric(Text12.Text)) Then
MsgBox " please enter vaild tyre rear size", vbExclamation, "stock1"
End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
If Not (IsNumeric(Text5.Text)) Then
MsgBox " please enter vaild tank size", vbExclamation, "stock1"
End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If Not (IsNumeric(Text7.Text)) Then
MsgBox " please enter number of gears", vbExclamation, "stock1"
End If
If Len(Text7.Text) < 2 Or Len(Text7.Text) > 2 Then
MsgBox "invaild number", vbExclamation, "stock1"
End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Not (IsNumeric(Text9.Text)) Then
MsgBox " please enter vaild length", vbExclamation, "stock1"
End If

End Sub
