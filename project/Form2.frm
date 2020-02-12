VERSION 5.00
Begin VB.Form customerr 
   BackColor       =   &H80000008&
   Caption         =   "customer"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   9195
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   9
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   525
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Customer address"
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
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Customer phone no"
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
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Customer email"
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Customer name"
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
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "customerr"
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
rs.Open "select * from customer", con, adOpenDynamic, adLockOptimistic
Set Text1.DataSource = rs
Text1.DataField = "cname"
Set Text2.DataSource = rs
Text2.DataField = "cemail"
Set Text3.DataSource = rs
Text3.DataField = "cphno"
Set Text4.DataSource = rs
Text4.DataField = "cadd"


End Sub

