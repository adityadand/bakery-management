VERSION 5.00
Begin VB.Form stockk 
   BackColor       =   &H80000008&
   Caption         =   "stock"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   9750
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "new"
      Height          =   735
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      Height          =   735
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   735
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      Height          =   735
      Left            =   6000
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   735
      Left            =   6960
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "update"
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Quantity"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Price"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Item name"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Item no"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "stockk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()
rs.AddNew

End Sub

Private Sub Command2_Click()
rs.Update


End Sub

Private Sub Command3_Click()
rs.Delete

End Sub

Private Sub Command4_Click()
rs.MovePrevious


End Sub

Private Sub Command5_Click()
rs.MoveNext
End Sub

Private Sub Command6_Click()
rs.Update

End Sub

Private Sub Form_Load()

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\bakery.mdb"
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\admin\Documents\bakery.mdb"
con.Open
rs.CursorLocation = adUseClient
rs.Open "select * from stock", con, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Text1.DataField = "item no"
Set Text2.DataSource = rs
Text2.DataField = "item name"
Set Text3.DataSource = rs
Text3.DataField = "price"
Set Text4.DataSource = rs
Text4.DataField = "quantity"


End Sub



