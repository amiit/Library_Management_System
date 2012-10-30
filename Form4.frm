VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue Details"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   4920
      Picture         =   "Form4.frx":77BE
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Picture         =   "Form4.frx":7D88
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2760
      Width           =   375
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 - ID"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sayanraj Guha\Documents\VB Project 4th Semester\database\issuedat.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "issuedetails"
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 - BD"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sayanraj Guha\Documents\VB Project 4th Semester\database\bookdat.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bookdetails"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 -MD"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Sayanraj Guha\Documents\VB Project 4th Semester\database\memdat.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "memdetails"
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10800
      TabIndex        =   18
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Left            =   10800
      TabIndex        =   17
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form4.frx":8FCB
      Left            =   10680
      List            =   "Form4.frx":8FCD
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   615
      Left            =   10680
      TabIndex        =   14
      Top             =   4800
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   10680
      TabIndex        =   12
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   7920
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Click"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "to Find Member"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Code"
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
      Left            =   8760
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock in hand"
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
      Left            =   8400
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
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
      Left            =   8640
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Return"
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
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Issue"
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
      Left            =   600
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
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
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE DETAILS"
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
      Left            =   5760
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim i As Integer
Data2.Recordset.MoveFirst
Data2.Recordset.MoveLast
Data2.Recordset.MoveFirst
For i = 1 To Data2.Recordset.RecordCount
If (Data2.Recordset.Fields("bookcode").Value = Combo1.Text) Then
Text6.Text = Data2.Recordset.Fields("bookname").Value
Text7.Text = Data2.Recordset.Fields("bookstk").Value
Exit Sub
End If
Data2.Recordset.MoveNext
Next i
End Sub

Private Sub Command1_Click()
Form4.Hide
Form1.Show
End Sub


Private Sub Command2_Click()
Dim i, flag As Integer
Data1.Recordset.MoveFirst
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst

Text3.Text = Format$(Now, "dd - mm - yyyy")
Text4.Text = Format$(Now + 7, "dd - mm - yyyy")

If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text6.Text <> "" And Text7.Text <> "" And Combo1.Text <> "" Then
If Text5.Text = "" Then
MsgBox " Give Quantity ", vbOKOnly + vbExclamation, "Issue Result"
Exit Sub
End If
If Val(Text5.Text) > Val(Text7.Text) Then
MsgBox " Quantity more than Stock - Access Denied ", vbOKOnly + vbExclamation, "Issue Result"
Exit Sub
Else
For i = 1 To Data1.Recordset.RecordCount
If Text2.Text = Data1.Recordset.Fields("memname").Value Then
flag = 1
Exit For
Else
flag = 0
Data1.Recordset.MoveNext
End If
Next i
If flag = 0 Then
MsgBox " Member Name doesn't exist!! ", vbOKOnly + vbExclamation, "Issue Result"
Exit Sub
Else
If Data1.Recordset.Fields("memcode").Value = Text1.Text Then
Data3.Recordset.AddNew
Data3.Recordset.Fields("imemcode").Value = Data1.Recordset.Fields("memcode").Value
Data3.Recordset.Fields("imemname").Value = Data1.Recordset.Fields("memname").Value
Data3.Recordset.Fields("ibookcode").Value = Data2.Recordset.Fields("bookcode").Value
Data3.Recordset.Fields("ibookname").Value = Data2.Recordset.Fields("bookname").Value
Data3.Recordset.Fields("idateissue").Value = Format$(Now, "dd - mm - yyyy")
Data3.Recordset.Fields("idatereturn").Value = Format$(Now + 7, "dd - mm - yyyy")
Data3.Recordset.Fields("iquanti").Value = Text5.Text
Data3.Recordset.Update
Data2.Recordset.Edit
Data2.Recordset.Fields("bookstk").Value = Data2.Recordset.Fields("bookstk").Value - Val(Text5.Text)
Data2.Recordset.Update
MsgBox " Book Issue Successful !! ", vbOKOnly + vbInformation, "Issue Result"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text7.Text = Data2.Recordset.Fields("bookstk").Value
Command4.Visible = False
Command3.Visible = True
Else
MsgBox " Member Name and Code do not match!! ", vbOKOnly + vbExclamation, "Issue Result"
Exit Sub
End If
End If

End If
Else
Beep
End If

End Sub

Private Sub Command3_Click()
Dim respS As String
Dim respI, i As Integer
Data1.Recordset.MoveFirst
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst
If Text1.Text = "" And Text2.Text = "" Then
MsgBox " No Value Entered ", vbOKOnly + vbInformation, "Find Result"
Exit Sub
ElseIf (Text1.Text <> "" And Text2.Text = "") Then
For i = 1 To Data1.Recordset.RecordCount
If (Data1.Recordset.Fields("memcode").Value = Text1.Text) Then
Text2.Text = Data1.Recordset.Fields("memname").Value
Text1.Enabled = False
Text2.Enabled = False
Command3.Visible = False
Command4.Visible = True
Label10.Caption = "to Refresh"
Exit Sub
End If
Data1.Recordset.MoveNext
Next i
MsgBox " Not Found ", vbOKOnly + vbExclamation, "Find Result"
Text1.Enabled = False
Text2.Enabled = False
Command3.Visible = False
Command4.Visible = True
Label10.Caption = "to Refresh"
Else
respS = Text2.Text
For i = 1 To Data1.Recordset.RecordCount
If (Data1.Recordset.Fields("memname").Value = respS) Then
Text1.Text = Data1.Recordset.Fields("memcode").Value
Text1.Enabled = False
Text2.Enabled = False
Command3.Visible = False
Command4.Visible = True
Label10.Caption = "to Refresh"
Exit Sub
End If
Data1.Recordset.MoveNext
Next i
MsgBox " Not Found ", vbOKOnly + vbExclamation, "Find Result"
Text1.Enabled = False
Text2.Enabled = False
Command3.Visible = False
Command4.Visible = True
Label10.Caption = "to Refresh"
End If

End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text1.Enabled = True
Text2.Enabled = True
Command4.Visible = False
Command3.Visible = True
Label10.Caption = "to Find Member"

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

