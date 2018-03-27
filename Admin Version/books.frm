VERSION 5.00
Begin VB.Form books 
   Caption         =   "Books"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4695
   End
End
Attribute VB_Name = "books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List2.AddItem "ID" & vbTab & "Bookname"
admin.rs2.MoveFirst
While admin.rs2.EOF <> True
    List1.AddItem admin.rs2.Fields("bookid") & vbTab & admin.rs2.Fields("bookname")
    admin.rs2.MoveNext
Wend
End Sub

