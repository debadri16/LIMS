VERSION 5.00
Begin VB.Form regn 
   Caption         =   "Student registration"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Semester :"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Department :"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "regn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
home.db.Execute ("insert into students (name,dept,sem,password) values ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')")
home.rs.MoveFirst
While home.rs.Fields("name").Value <> Text1.Text
    home.rs.MoveNext
Wend
home.Text1.Text = home.rs.Fields("id")
home.Text2.Text = home.rs.Fields("password")
Unload Me
home.Command1_Click
MsgBox ("Note your Student Id. You will need this to login"), vbInformation, "Successfully registered"
End Sub

