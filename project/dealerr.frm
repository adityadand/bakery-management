VERSION 5.00
Begin VB.Form dealerr 
   BackColor       =   &H80000007&
   Caption         =   "Dealer"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
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
      Left            =   4680
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
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
      Left            =   8040
      TabIndex        =   12
      Top             =   3840
      Width           =   975
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
      Left            =   6840
      TabIndex        =   11
      Top             =   3840
      Width           =   975
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
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
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
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
      Left            =   720
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Dealer name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "dealerr"
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
rs.Open "select * from dealer", con, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Text1.DataField = "dealername"
Set Text2.DataSource = rs
Text2.DataField = "address"
Set Text3.DataSource = rs
Text3.DataField = "phone"
Set Text4.DataSource = rs
Text4.DataField = "email"


End Sub

