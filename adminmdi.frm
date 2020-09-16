VERSION 5.00
Begin VB.MDIForm MDIFORM1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10755
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20040
   LinkTopic       =   "MDIForm1"
   Picture         =   "adminmdi.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu EMPLOYEE 
      Caption         =   "EMPLOYEE MANAGEMENT"
      Begin VB.Menu ADD1 
         Caption         =   "ADD AND VIEW"
      End
      Begin VB.Menu SEARCH1 
         Caption         =   "SEARCH AND DELETE"
      End
   End
   Begin VB.Menu SALES 
      Caption         =   "SALES "
   End
   Begin VB.Menu SERVICE 
      Caption         =   "SERVICE "
   End
   Begin VB.Menu USEDBIKES 
      Caption         =   "USED BIKES "
   End
   Begin VB.Menu SPARE 
      Caption         =   "SPARE PARTS "
   End
   Begin VB.Menu BILLING 
      Caption         =   "BILLING"
      Begin VB.Menu PREVIOUS 
         Caption         =   "PREVIOUS BILLS"
      End
      Begin VB.Menu NEW 
         Caption         =   "NEW BILL"
      End
   End
   Begin VB.Menu STOCK 
      Caption         =   "STOCK"
      Begin VB.Menu ADD 
         Caption         =   "ADD"
      End
      Begin VB.Menu VIEW 
         Caption         =   "VIEW"
         Index           =   1
      End
      Begin VB.Menu SEARCH 
         Caption         =   "SEARCH"
      End
   End
   Begin VB.Menu BOOKING 
      Caption         =   "BOOKING"
      Begin VB.Menu BOOK 
         Caption         =   "BOOK BIKE"
         Index           =   1
      End
      Begin VB.Menu BOOKED 
         Caption         =   "BOOKED LIST"
      End
   End
   Begin VB.Menu SALARY 
      Caption         =   "SALARY MANAGEMENT"
   End
   Begin VB.Menu REPORTS 
      Caption         =   "REPORTS"
      Begin VB.Menu SALES2 
         Caption         =   "SALES_REPORT"
      End
      Begin VB.Menu STOCKR 
         Caption         =   "STOCK_REPORT"
      End
   End
   Begin VB.Menu log 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "MDIFORM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
stock1.Show
End Sub

Private Sub ADD1_Click()
emp1.Show
End Sub

Private Sub BOOK_Click(Index As Integer)
booking1.Show
End Sub

Private Sub BOOKED_Click()
bookinggrid.Show

End Sub

Private Sub log_Click()
register.Show
End Sub

Private Sub MDIForm_Load()
'Timer1.Enabled = True
'Label1.Caption = Time
End Sub

Private Sub NEW_Click()
billing1.Show
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub PREVIOUS_Click()
billinggrid.Show
End Sub

Private Sub SALARY_Click()
salary1.Show
End Sub

Private Sub SALES_Click()
sales1.Show
End Sub

Private Sub SALES2_Click()
DataReport1.Show
End Sub

Private Sub SEARCH_Click()
stocksearch.Show
End Sub

Private Sub SEARCH1_Click()
empsrch.Show
End Sub

Private Sub service_Click()
service1.Show
End Sub

Private Sub SPARE_Click()
spare1.Show
End Sub

Private Sub STOCKR_Click()
DataReport2.Show
End Sub

Private Sub USEDBIKES_Click()
used1.Show
End Sub

Private Sub VIEW_Click(Index As Integer)
stockgrid.Show
End Sub
