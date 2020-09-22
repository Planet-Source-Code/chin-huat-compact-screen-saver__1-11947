VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lines As clsLines

Private Sub Form_Load()
    Set Lines = New clsLines
    Lines.Init frmMain, 50, 3 ' Infinite posibilities...
    'Lines.Init frmMain, 5, 15 ' Dancing wires
    'Lines.Init frmMain, 50, 3, , , , , , , , , , , 15, 15, 15, 8, 8, , True, True, 255, 255, 255, 5, 5, 0 ' User define
End Sub

Private Sub Form_Paint()
    Lines.Repaint
End Sub

Private Sub Timer1_Timer()
    Lines.Run
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lines = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): Unload Me: End Sub
Private Sub Form_LostFocus(): Unload Me: End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): Unload Me: End Sub
