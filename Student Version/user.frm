VERSION 5.00
Begin VB.Form user 
   Caption         =   "User"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Books issued/pending"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update password"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Logout"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send issue request"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "Select a book"
      Top             =   3000
      Width           =   7335
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Semester :"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Department :"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Student Id :"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs2 As Recordset, rs1 As Recordset, bname As String

Private Sub Combo1_Click()
bname = Combo1.Text
Command1.Enabled = True
End Sub

Private Sub Command1_Click()
rs2.MoveFirst
While rs2.Fields("bookname").Value <> bname
    rs2.MoveNext
Wend
home.db.Execute ("insert into issues (stuid,bookid) values ('" & Label5.Caption & "','" & rs2.Fields("bookid") & "')")
MsgBox (Str(rs2.Fields("bookid")) + ": " + bname), vbInformation, "Request sent for -"
End Sub

Private Sub Command2_Click()
Unload Me
home.Show
End Sub

Private Sub Command4_Click()
issues.Show
End Sub

Private Sub Form_Load()
Set rs2 = home.db.OpenRecordset("select * from books")
Set rs1 = home.db.OpenRecordset("select * from issues")
Command1.Enabled = False
Label5.Caption = home.rs.Fields("id")
Label6.Caption = home.rs.Fields("name")
Label7.Caption = home.rs.Fields("dept")
Label8.Caption = home.rs.Fields("sem")
rs2.MoveFirst
While rs2.EOF = False
    Combo1.AddItem (rs2.Fields("bookname"))
    rs2.MoveNext
Wend
End Sub
