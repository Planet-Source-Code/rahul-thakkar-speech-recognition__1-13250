VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1665
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "               Press Esc To Close This Window"
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Change It and Make Some New Programs Out of It "
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "                This Program is Created by Rahul :)"
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
main.Show
Form2.Visible = False
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
X = 0
a = Label1.Count
For i = 1 To a Step 1
      Label1.Item(X).Visible = True
   If Label1.Item(X).Left <= 800 Then
   X = X + 1
   Else
   Label1.Item(X).Left = Label1.Item(X).Left - 100
   End If
Next
End Sub
