VERSION 5.00
Begin VB.Form signup1 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18945
   LinkTopic       =   "Form1"
   Picture         =   "signup.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   18945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.PictureBox signup 
      Height          =   11760
      Left            =   0
      Picture         =   "signup.frx":698E7
      ScaleHeight     =   11700
      ScaleWidth      =   20670
      TabIndex        =   0
      Top             =   -240
      Width           =   20730
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   6720
         TabIndex        =   10
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "login"
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
         Left            =   5520
         Picture         =   "signup.frx":C6359
         TabIndex        =   8
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " login"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   5280
         TabIndex        =   9
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   3840
         TabIndex        =   7
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   3960
         TabIndex        =   6
         Top             =   3000
         Width           =   2175
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "signup1"
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
Dim sts As Boolean
Dim ca As Integer
sts = False
str = "select * from reg_log"
rs.Open str, con, adOpenKeyset
rs.MoveFirst
Do Until sts = True Or rs.EOF
If rs.EOF Then
MsgBox "invalid login,username and password are not correct", vbOKOnly
rs.Close
con.Close
End If
If rs("user_name").Value = Text1.Text And rs("Password").Value = Text2.Text Then
sts = True
Else
rs.MoveNext
End If
Loop
If (sts = True) Then
If (Text1.Text = "admin") Then
MDIFORM1.Picture = LoadPicture("D:\MINI PROJECT\IMAGES\admin.jpg")
        
  MDIFORM1.BOOKING.Visible = False
  MDIFORM1.BILLING.Visible = False
  MDIFORM1.Show
 Else
 MDIFORM1.Picture = LoadPicture("D:\MINI PROJECT\projt\emp2.jpg")
        MDIFORM1.EMPLOYEE.Visible = False
         MDIFORM1.SALARY.Visible = False
          MDIFORM1.SALES.Visible = False
           MDIFORM1.SERVICE.Visible = False
           
 
            MDIFORM1.REPORTS.Visible = False
           ' Command1.Visible = False
            ' MDIFORM1.SPARE.Visible = False
             'Command1.Visible = False
             ' MDIFORM1.STOCK.Visible = False
             ' Command1.Visible = False
        'MDIFORM1.REPORTS.Visible = False
  MDIFORM1.Show
  End If
End If
If (sts = False) Then
MsgBox "invalid id", vbInformation, vbOKOnly
End If


'If sts = True Then
'Unload Me
'MDIFORM1.Show
'rs.MovePrevious
'rs.Close
'con.Close
'Else
'MsgBox "Invalid username or Password", vbInformation
'rs.MovePrevious
'rs.Close
con.Close
'End If
'End Sub

'Private Sub Form_Load()
'Dim flag As Integer
'Dim ca As Integer
'flag = 0

'Call connect
'str = "select*from reg_log"
'rs.Open str, con, adOpenKeyset
'rs.MoveFirst

'Do Until rs.EOF Or flag = 1

'If rs("user_name").Value = Text1.Text And rs("Password").Value = Text2.Text Then
         'MsgBox ("valid")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         

  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '        flag = 1
        ' If rs("cat") = 2 Then
         
  '       Label3.Caption = "Employyee"
         
   '     MDIFORM1.Caption = "Employee"
    '    MDIFORM1.Picture = LoadPicture("D:\MINI PROJECT\projt\emp.jpg")
     '   MDIFORM1.Show
      '  MDIFORM1.SHOWROOM.Visible = False
       ' MDIFORM1.REPORTS.Visible = False
        ' str1 = "insert into reg_log(user_name,Password)values(' " & Text1.Text & " ',' " & Text2.Text & " ')"
         'con.Execute str1
         
        'Unload Me
'End If


' If rs("cat") = 1 Then
            
 '           Label3.Caption = "Administator"
            
  '       MDIFORM1.Picture = LoadPicture("F:\Mini Poject\Images\admin.jpg")
   '      MDIFORM1.Caption = "Admin"
    '     MDIFORM1.Show
     '    MDIFORM1.GALLARY.Visible = False
      '   MDIFORM1.BOOKING.Visible = False
       '  MDIFORM1.BILLING.Visible = False
         
        ' str1 = "insert into reg_log(user_name,Password)values (' " & Text1.Text & " ',' " & Text2.Text & "')"
         'con.Execute str1
         
         
         'Unload Me
   '  End If
         





'Else
   'rs.MoveNext
        
'End If
       
'Loop

 'If flag = 0 Then

  '          MsgBox ("Invalid user")
   '         Text1.Text = " "
    '        Text2.Text = " "
     '       Unload Me
      '      signup1.Show
       '
            
            

 'End If
 

    
     
'rs.Close
'con.Close

End Sub

'End Sub

Private Sub Label5_Click()

End Sub

