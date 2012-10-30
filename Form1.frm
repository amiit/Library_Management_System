VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Manager"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RETURN DETAILS"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   4
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ISSUE DETAILS"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11880
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BOOK DETAILS"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MEMBER DETAILS"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY  MANAGEMENT  SYSTEM"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Data2.Recordset.MoveFirst
Form4.Data2.Recordset.MoveLast
Form4.Data2.Recordset.MoveFirst
Form4.Combo1.Clear
Do Until Form4.Data2.Recordset.EOF = True
Form4.Combo1.AddItem (Form4.Data2.Recordset.Fields("bookcode").Value)
Form4.Data2.Recordset.MoveNext
Loop
Form4.Data2.Recordset.MoveFirst
Form4.Text1.Text = ""
Form4.Text2.Text = ""
Form4.Text3.Text = ""
Form4.Text4.Text = ""
Form4.Text5.Text = ""
Form4.Text6.Text = ""
Form4.Text7.Text = ""

Form4.Text3.Text = Format$(Now, "dd - mm - yyyy")
Form4.Text4.Text = Format$(Now + 7, "dd - mm - yyyy")

Form1.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub Command5_Click()
End
End Sub

