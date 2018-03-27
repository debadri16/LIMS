VERSION 5.00
Begin VB.Form Students 
   Caption         =   "Students"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      ForeColor       =   &H80000007&
      Height          =   450
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List2.AddItem "ID" & vbTab & "Student Name" & vbTab & "Dept" & vbTab & "Sem"
admin.rs1.MoveFirst
While admin.rs1.EOF <> True
    List1.AddItem admin.rs1.Fields("id") & vbTab & admin.rs1.Fields("name") & vbTab & admin.rs1.Fields("dept") & vbTab & admin.rs1.Fields("sem")
    admin.rs1.MoveNext
Wend
End Sub
