VERSION 5.00
Begin VB.Form spare1 
   Caption         =   "Form2"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9765
   ScaleWidth      =   18540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "search"
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
      Left            =   10200
      TabIndex        =   19
      Top             =   8400
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   14655
      Index           =   0
      Left            =   120
      Picture         =   "sparesearch.frx":0000
      ScaleHeight     =   14595
      ScaleWidth      =   29595
      TabIndex        =   0
      Top             =   -120
      Width           =   29655
      Begin VB.PictureBox Picture1 
         Height          =   14655
         Index           =   1
         Left            =   720
         Picture         =   "sparesearch.frx":83DB2
         ScaleHeight     =   14595
         ScaleWidth      =   29595
         TabIndex        =   20
         Top             =   120
         Width           =   29655
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "sparesearch.frx":107B64
            Left            =   10560
            List            =   "sparesearch.frx":107B80
            TabIndex        =   30
            Top             =   2880
            Width           =   2415
         End
         Begin VB.CommandButton Command7 
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
            TabIndex        =   29
            Top             =   7560
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
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
            TabIndex        =   28
            Top             =   7560
            Width           =   1575
         End
         Begin VB.TextBox Text12 
            Height          =   480
            Left            =   10560
            TabIndex        =   27
            Top             =   6120
            Width           =   2460
         End
         Begin VB.TextBox Text11 
            Height          =   480
            Left            =   10560
            TabIndex        =   26
            Top             =   5280
            Width           =   2460
         End
         Begin VB.TextBox Text10 
            Height          =   480
            Left            =   10560
            TabIndex        =   25
            Top             =   4560
            Width           =   2460
         End
         Begin VB.TextBox Text9 
            Height          =   480
            Left            =   10560
            TabIndex        =   24
            Top             =   3720
            Width           =   2460
         End
         Begin VB.TextBox Text8 
            Height          =   495
            Left            =   10560
            TabIndex        =   23
            Top             =   1920
            Width           =   2505
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   10560
            TabIndex        =   22
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000A&
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
            Height          =   735
            Left            =   11400
            MaskColor       =   &H8000000B&
            TabIndex        =   21
            Top             =   7560
            Width           =   1935
         End
         Begin VB.Label Label16 
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
            TabIndex        =   38
            Top             =   6120
            Width           =   2100
         End
         Begin VB.Label Label15 
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
            TabIndex        =   37
            Top             =   5280
            Width           =   2100
         End
         Begin VB.Label Label14 
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
            TabIndex        =   36
            Top             =   3720
            Width           =   2820
         End
         Begin VB.Label Label13 
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
            TabIndex        =   35
            Top             =   4560
            Width           =   2100
         End
         Begin VB.Label Label12 
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
            TabIndex        =   34
            Top             =   2640
            Width           =   3540
         End
         Begin VB.Label Label11 
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
            TabIndex        =   33
            Top             =   1920
            Width           =   4575
         End
         Begin VB.Label Label10 
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
            TabIndex        =   32
            Top             =   1200
            Width           =   4380
         End
         Begin VB.Label Label9 
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
            TabIndex        =   31
            Top             =   0
            Width           =   4815
         End
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
         Left            =   11400
         MaskColor       =   &H8000000B&
         TabIndex        =   2
         Top             =   7560
         Width           =   1935
      End
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
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "sparesearch.frx":107BAC
         Left            =   10560
         List            =   "sparesearch.frx":107BC8
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

Private Sub Command2_Click()
sparegrid.Show
End Sub

Private Sub Command4_Click()
sparesearch2.Show
End Sub

Private Sub Command5_Click()
MDIFORM1.Show
End Sub

Private Sub Command6_Click()
If Text3.Text = "" Then
MsgBox "spare id is mandatory", vbInformation
Exit Sub
End If
If Text8.Text = "" Then
MsgBox "spare nameis mandatory", vbInformation
Exit Sub
End If
If Combo1.Text = "" Then
MsgBox " cc is mandatory", vbInformation
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "model number is mandatory", vbInformation
Exit Sub
End If
If Text10.Text = "" Then
MsgBox "quantity is mandatory", vbInformation
Exit Sub
End If
If Text11.Text = "" Then
MsgBox "colour is mandatory", vbInformation
Exit Sub
End If
If Text12.Text = "" Then
MsgBox "price is mandatory", vbInformation
Exit Sub
End If
Call connect
str1 = " insert into spare(id,name,cc,model,quantity,colour,price)values(' " & Text3.Text & " ' ,' " & Text8.Text & " ',' " & Combo1.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ',' " & Text11.Text & " ',' " & Text12.Text & " ')"
con.Execute str1
MsgBox "Successfull", vbInformation
Text3.Text = " "
Text8.Text = " "
Combo1.Text = " "
'Text3.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
con.Close
End Sub

Private Sub Command7_Click()
sparegrid.Show
End Sub

Private Sub Form_Load()
Dim n As Integer
If con.State = adStateOpen Then
rs.Close
con.Close
End If
Text3.Enabled = False
Call connect
str1 = "select * from spare"
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

Private Sub Text10_Validate(Cancel As Boolean)
If Not (IsNumeric(Text10.Text)) Then
MsgBox " please enter vaild quantity", vbExclamation, "spare1"
End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
If Not (IsNumeric(Text12.Text)) Then
MsgBox " please enter vaild `amount", vbExclamation, "booking1"
End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Not (IsNumeric(Text3.Text)) Then
MsgBox " please enter vaild spare id ", vbExclamation, "booking1"
End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Not (IsNumeric(Text9.Text)) Then
MsgBox " please enter vaild model number", vbExclamation, "spare1"
End If
End Sub
