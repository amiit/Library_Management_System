VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "last"
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "first"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "next"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "prev"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   8880
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "grade"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "students"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sayanraj Guha\Documents\VB Project 4th Semester\database\db.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   1140
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "students"
      Top             =   2160
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Command2_Click()
Data1.Recordset.Delete
Data1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Data1.Refresh

End Sub

Private Sub Command5_Click()
Data1.Recordset.MovePrevious

End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
Data1.Recordset.MoveFirst

End Sub

Private Sub Command8_Click()
Data1.Recordset.MoveLast
End Sub
