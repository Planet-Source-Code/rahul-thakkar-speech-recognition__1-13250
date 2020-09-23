VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form main 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DirectSR1 
      Height          =   75
      Left            =   360
      OleObjectBlob   =   "speech reco.frx":0000
      TabIndex        =   4
      Top             =   1410
      Width           =   30
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Ready to Hear Something From You :)"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "You Said :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "                            Speech Recognition Program"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       For Moving Form Without Title                                                                                                                                                            '
'***********************************************************************************************************************************************************'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       For Moving Form Without Title                                                                                                                                                            '
'***********************************************************************************************************************************************************'

Private Sub DirectSR1_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)
Label3.Caption = Phrase

Select Case Phrase
Case Is = "start Notepad"
         rt = Shell("c:\windows\Notepad.exe", vbNormalFocus)
Case Is = "start Calculator"
         rt = Shell("c:\windows\Calc.exe", vbNormalFocus)
Case Is = "start explorer"
         rt = Shell("c:\windows\EXPLORER.EXE ", vbNormalFocus)
Case Is = "start controlpanel"
         Shell "rundll32 shell32,Control_RunDLL", vbNormalFocus
Case Is = ""
         Label3.Caption = "Sorry,Was Not AbleTo Understand :)"
End Select
End Sub

Private Sub Form_Load()
Dim rt As Integer

DirectSR1.GrammarFromString "[grammer]" & vbNewLine & "type=cfg" & vbNewLine & "[<start>]" & "<start>=start Notepad" & vbNewLine & "<start>=start Calculator" & vbNewLine & "<start>= start explorer" & vbNewLine
DirectSR1.Activate
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' For Moving Form Without Title,It is for the label which i have used as title
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
Form2.Label1.Item(0).Left = 4680
Form2.Label1.Item(1).Left = 4680
Form2.Label1.Item(2).Left = 4680
main.Hide
Form2.Show
End Sub
