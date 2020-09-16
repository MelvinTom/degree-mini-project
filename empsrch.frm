VERSION 5.00
Begin VB.Form empsrch 
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   20085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   11295
      Index           =   0
      Left            =   0
      Picture         =   "empsrch.frx":0000
      ScaleHeight     =   11235
      ScaleWidth      =   21750
      TabIndex        =   4
      Top             =   0
      Width           =   21810
      Begin VB.CommandButton Command2 
         Caption         =   "delete"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17520
         TabIndex        =   49
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "search"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15360
         TabIndex        =   48
         Top             =   7200
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   16560
         TabIndex        =   43
         Top             =   6360
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   5175
         Left            =   1680
         Picture         =   "empsrch.frx":3AD13
         ScaleHeight     =   5115
         ScaleWidth      =   4395
         TabIndex        =   26
         Top             =   1080
         Width           =   4455
         Begin VB.TextBox Text15 
            Height          =   375
            Left            =   1920
            TabIndex        =   34
            Top             =   4440
            Width           =   2295
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   1920
            TabIndex        =   33
            Top             =   3720
            Width           =   2175
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   1920
            TabIndex        =   32
            Top             =   3240
            Width           =   2175
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   1920
            TabIndex        =   31
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   1920
            TabIndex        =   30
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   1920
            TabIndex        =   29
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1920
            TabIndex        =   28
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   1920
            TabIndex        =   27
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "email"
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
            Left            =   120
            TabIndex        =   42
            Top             =   4440
            Width           =   1695
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "country"
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
            Left            =   120
            TabIndex        =   41
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "state"
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
            Left            =   120
            TabIndex        =   40
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "pin code"
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
            Left            =   120
            TabIndex        =   39
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "town"
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
            Left            =   120
            TabIndex        =   38
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "city"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "village street"
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
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "house name"
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
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000080FF&
         Height          =   2655
         Left            =   4320
         ScaleHeight     =   2595
         ScaleWidth      =   9195
         TabIndex        =   17
         Top             =   6480
         Width           =   9255
         Begin VB.TextBox Text19 
            Height          =   405
            Left            =   3000
            TabIndex        =   21
            Top             =   2040
            Width           =   6015
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Left            =   3000
            TabIndex        =   20
            Top             =   1440
            Width           =   6015
         End
         Begin VB.TextBox Text17 
            Height          =   405
            Left            =   3000
            TabIndex        =   19
            Top             =   840
            Width           =   6015
         End
         Begin VB.TextBox Text16 
            Height          =   375
            Left            =   3000
            TabIndex        =   18
            Top             =   240
            Width           =   6015
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "alternative mob.no:"
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
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "mobile number"
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
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "branch details"
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
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "bank account no:"
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
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5055
         Index           =   1
         Left            =   10560
         Picture         =   "empsrch.frx":3DF66
         ScaleHeight     =   5025
         ScaleWidth      =   4785
         TabIndex        =   5
         Top             =   1200
         Width           =   4815
         Begin VB.TextBox Text22 
            Height          =   375
            Left            =   2400
            TabIndex        =   47
            Top             =   3120
            Width           =   2175
         End
         Begin VB.TextBox Text21 
            Height          =   405
            Left            =   2400
            TabIndex        =   46
            Top             =   2520
            Width           =   2175
         End
         Begin VB.TextBox Text20 
            Height          =   375
            Left            =   2400
            TabIndex        =   45
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2400
            TabIndex        =   9
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "empsrch.frx":411B9
            Left            =   2400
            List            =   "empsrch.frx":411C3
            TabIndex        =   8
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2400
            TabIndex        =   7
            Top             =   3720
            Width           =   2175
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2400
            TabIndex        =   6
            Top             =   4320
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "BASIC PAY"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   4440
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "JOB TITLE"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTMENT"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF JOIN"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF BIRTH"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE NAME"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "contact/bank  info"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   6120
         TabIndex        =   52
         Top             =   5760
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "main details"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   0
         Left            =   11640
         TabIndex        =   51
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Index           =   1
         Left            =   3120
         TabIndex        =   50
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "EMPLOYEE ID"
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
         Index           =   1
         Left            =   14640
         TabIndex        =   44
         Top             =   6360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   11040
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "empsrch"
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

Private Sub Combo3_Change()
Dim str1 As String
Call connect
str1 = "select sid from emp1"
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Combo2.AddItem (rs.Fields(1))
rs.MoveNext
Loop
emp1.Refresh
rs.Close
con.Close
End Sub

Private Sub Command1_Click()
Call connect
str1 = "select * from emp1 where id=" & Combo3.Text & ""
con.Execute str1
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text1.Text = rs.Fields(0)


Combo1.Text = rs.Fields(2)



Text20.Text = rs.Fields(3)
Text21.Text = rs.Fields(4)
Text22.Text = rs.Fields(5)
Text6.Text = rs.Fields(6)

Text7.Text = rs.Fields(7)
Text8.Text = rs.Fields(8)
Text9.Text = rs.Fields(9)
Text10.Text = rs.Fields(10)
Text11.Text = rs.Fields(11)
Text12.Text = rs.Fields(12)

Text13.Text = rs.Fields(13)
Text14.Text = rs.Fields(14)
Text15.Text = rs.Fields(15)
Text16.Text = rs.Fields(16)
Text17.Text = rs.Fields(17)
Text18.Text = rs.Fields(18)
Text19.Text = rs.Fields(19)

rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Command2_Click()
Call connect
str1 = "delete from emp1 where id= '" & Combo3.Text & "' "
con.Execute str1
MsgBox "successfully deleted", vbInformation
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
Text17.Text = " "
Text18.Text = " "
Combo1.Text = " "
Text19.Text = " "
con.Close
End Sub

Private Sub Form_Load()
Dim str1 As String
Call connect
str1 = "select * from emp1"
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
'MsgBox (rs.Fields(1))
Combo3.AddItem (rs.Fields(1))
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

