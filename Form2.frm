VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Details"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   22
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sayanraj Guha\Documents\VB Project 4th Semester\database\memdat.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "memdetails"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ISSUE"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   20
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   13
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataField       =   "memfees"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   10680
      TabIndex        =   10
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      DataField       =   "memname"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   10680
      TabIndex        =   9
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "memrdate"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "memaddress"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "memcode"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fees"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Name"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Date"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Member Code (ID)"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER DETAILS"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
Command3.Enabled = False
Command4.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command10_Click()
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
Form2.Hide
Form4.Show
End Sub

Private Sub Command11_Click()
Data1.Recordset.CancelUpdate
Command11.Visible = False
Command12.Visible = True
Command5.Visible = True
Command6.Visible = False
End Sub

Private Sub Command12_Click()
Data1.Refresh
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = True
Command5.Visible = True
Command6.Visible = False
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveLast
Command3.Enabled = True
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Data1.Recordset.MovePrevious
Command2.Enabled = True
Command4.Enabled = True
If (Data1.Recordset.BOF = True) Then
Data1.Recordset.MoveFirst
Command3.Enabled = False
End If
End Sub

Private Sub Command4_Click()
Data1.Recordset.MoveNext
Command1.Enabled = True
Command3.Enabled = True
If (Data1.Recordset.EOF = True) Then
Data1.Recordset.MoveLast
Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.AddNew
Command5.Visible = False
Command6.Visible = True
Command11.Visible = True
Command12.Visible = False
End Sub

Private Sub Command6_Click()
Data1.Recordset.Update
Command6.Visible = False
Command5.Visible = True
Command11.Visible = False
Command12.Visible = True
End Sub

Private Sub Command7_Click()
Data1.Recordset.Delete
Data1.Recordset.MoveNext
If (Data1.Recordset.EOF = True) Then
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command8_Click()
Dim id, i, length As Integer, prompt As String
'scan the recordset from 1st to last and go back to 1st
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
'activate the prev,next,first,last buttons
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
length = Data1.Recordset.RecordCount
prompt = InputBox("Enter Member ID :", "Search for Member")
If (Len(prompt) = 0) Then
MsgBox " No Search Item Entered / Cancelled", vbOKOnly + vbExclamation, "Search Result"
Exit Sub
End If
id = Val(prompt)
For i = 1 To length
If (id = Val(Text1.Text)) Then
Exit Sub
End If
Data1.Recordset.MoveNext
Next i
MsgBox "Search Item Not Found", vbOKOnly + vbExclamation, "Search Result"
Data1.Refresh
End Sub

Private Sub Command9_Click()
Form2.Hide
Form1.Show
End Sub


Private Sub Form_Load()
Command3.Enabled = False
Command1.Enabled = False
End Sub
