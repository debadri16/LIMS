VERSION 5.00
Begin VB.Form home 
   Caption         =   "Student Login"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Student Id :"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database, rs As Recordset, sid As Long

Public Sub Command1_Click()
rs.MoveFirst
While True
    If Text1.Text = rs.Fields("id") And Text2.Text = rs.Fields("password") Then
        home.Hide
        user.Show
        sid = Text1.Text
        Exit Sub
    End If
    rs.MoveNext
    If rs.EOF = True Then
        MsgBox ("No such entry found"), vbCritical, "Incorrect credentials"
        Exit Sub
    End If
Wend
End Sub

Private Sub Command2_Click()
regn.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("E:\lims\student V\libraryinfo.mdb")
Set rs = db.OpenRecordset("select * from students")
End Sub
