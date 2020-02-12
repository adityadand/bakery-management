VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H80000008&
   Caption         =   "About"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "www.linkdin.com/adityadand"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "www.github.com/adityadand"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Made By Aditya Dand"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
