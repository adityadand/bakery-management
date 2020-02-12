VERSION 5.00
Begin VB.Form billl 
   BackColor       =   &H80000008&
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   3360
      TabIndex        =   15
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   13
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   10
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   9
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Bill id"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Item bought"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Customer name"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "billl"
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
rs.Open "select * from bill", con, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Text1.DataField = "bill id"
Set Text2.DataSource = rs
Text2.DataField = "customer name"
Set Text3.DataSource = rs
Text3.DataField = "item bought"
Set Text4.DataSource = rs
Text4.DataField = "amount"
Set Text5.DataSource = rs
Text5.DataField = "bdate"




End Sub

