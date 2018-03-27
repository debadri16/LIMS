VERSION 5.00
Begin VB.Form issues 
   Caption         =   "Books issued"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "issues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database, rs As Recordset, rs1 As Recordset, rs2 As Recordset

Private Sub Form_Load()
Set db = OpenDatabase("E:\lims\student V\libraryinfo.mdb")
Set rs = db.OpenRecordset("select * from students")
Set rs1 = db.OpenRecordset("select * from issues")
Set rs2 = db.OpenRecordset("select * from books")
rs1.MoveFirst
While True
    If rs1.Fields("stuid").Value = user.Label5.Caption Then
        rs2.MoveFirst
        While rs1.Fields("bookid") <> rs2.Fields("bookid")
            rs2.MoveNext
        Wend
        List1.AddItem (rs2.Fields("bookname"))
        If rs1.Fields("validate") = 0 Then
            List2.AddItem ("Pending")
        Else
            List2.AddItem ("10/4/2018 at 11:30")
        End If
    End If
    rs1.MoveNext
    If rs1.EOF = True Then
        Exit Sub
    End If
Wend
End Sub


