VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11655
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu stock 
      Caption         =   "stock"
   End
   Begin VB.Menu customer 
      Caption         =   "customer"
   End
   Begin VB.Menu dealer 
      Caption         =   "dealer"
   End
   Begin VB.Menu bill 
      Caption         =   "bill"
   End
   Begin VB.Menu order 
      Caption         =   "order"
   End
   Begin VB.Menu about 
      Caption         =   "about"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
Form3.Show
End Sub

Private Sub bill_Click()
billl.Show
End Sub

Private Sub customer_Click()
customerr.Show
End Sub

Private Sub dealer_Click()
dealerr.Show

End Sub


Private Sub order_Click()
orderr.Show
End Sub

Private Sub stock_Click()
stockk.Show
End Sub
