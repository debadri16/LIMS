VERSION 5.00
Begin VB.Form admin 
   Caption         =   "Admin Panel"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Books List"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Students List"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Library Information and Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "Author"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Add books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Issue requests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database, rs As Recordset, rs1 As Recordset, rs2 As Recordset

Private Sub Command1_Click()
Var = Text1.Text + " by " + Text2.Text
db.Execute ("insert into books (bookname) values ('" & Var & "')")
Text1.Text = ""
Text2.Text = ""
MsgBox ("Book list updated"), vbInformation, "Added"
End Sub

Private Sub Command2_Click()
Students.Show
End Sub

Private Sub Command3_Click()
books.Show
End Sub

Private Sub Command4_Click()
Unload Me
admin.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("E:\lims\student V\libraryinfo.mdb")
Set rs = db.OpenRecordset("select * from issues")
Set rs1 = db.OpenRecordset("select * from students")
Set rs2 = db.OpenRecordset("select * from books")
rs.MoveFirst
While rs.EOF <> True
    If rs.Fields("validate") = 0 Then
        List1.AddItem (rs.Fields("issueid"))
    End If
    rs.MoveNext
Wend
End Sub

Private Sub List1_Click()
rs.MoveFirst
While rs.Fields("issueid") <> List1.Text
    rs.MoveNext
Wend
rs1.MoveFirst
While rs1.Fields("id") <> rs.Fields("stuid")
    rs1.MoveNext
Wend
rs2.MoveFirst
While rs2.Fields("bookid") <> rs.Fields("bookid")
    rs2.MoveNext
Wend
r = MsgBox("Accept request:" & Chr(13) & Str(rs1.Fields("id")) + " - " + rs1.Fields("name") + " of " + rs1.Fields("dept") + " sem-" + Str(rs1.Fields("sem")) + " requested" & Chr(13) & Str(rs2.Fields("bookid")) + " - " + rs2.Fields("bookname"), vbYesNo, "Confirm issue request")
If r = vbYes Then
    rs.Edit
    rs.Fields("validate") = 1
    rs.Update
    MsgBox ("Confirmed"), vbInformation, "Request Accepted"
    Unload Me
    admin.Show
End If
End Sub



