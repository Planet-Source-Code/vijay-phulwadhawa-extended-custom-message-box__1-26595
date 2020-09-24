VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   1035
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   1035
   End
   Begin VB.Frame fraSeparator 
      Height          =   75
      Left            =   60
      TabIndex        =   0
      Top             =   855
      Width           =   6135
   End
   Begin VB.Label lblMsgHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   795
      TabIndex        =   3
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   765
      TabIndex        =   2
      Top             =   315
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 08/04/2001 10:11:24 AM
'Project Name      :   msgdll
'Form/Class Name   :   frmMessage (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************


Private Sub cmdButton_Click(Index As Integer)
ButtonClicked = Index
Unload Me
End Sub

Private Sub Form_Activate()
Dim I As Byte
For I = 0 To frmMessage.cmdButton.UBound
    If frmMessage.cmdButton(I).Default = True Then
        frmMessage.cmdButton(I).SetFocus
        Exit For
    End If
Next I
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'reinitialize the timer
  Timer1.Enabled = False
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
SetUpDialogForm Me, Me.Left, Me.Top
Timer1.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And bAllowMove Then
'Release capture
    Call ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim I As Integer
For I = 0 To Me.cmdButton.UBound
    If Me.cmdButton(I).Default Then
        Me.cmdButton(I).SetFocus
        Exit For
    End If
Next

'Me.Height = Me.Height - 1
'Me.Height = Me.Height + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload all buttons
Dim I As Integer
For I = 1 To Me.cmdButton.Count - 1
    Unload Me.cmdButton(I)
Next

Timer1.Enabled = False
Set frmMessage = Nothing
End Sub

Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblMsgHead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub SetUpDialogForm(frm As Object, XPos As Long, YPos As Long)
Dim CurStyle As Long, NewStyle As Long
Dim Rectangle As RECT
Const WS_DLGFRAME = &H400000

With Rectangle
.Top = frm.Top
.Left = frm.Left
.Right = frm.Left + frm.Width
.Bottom = frm.Top + frm.Height
End With
frm.BorderStyle = 2 'sizeable
SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
'retrieve the window style
CurStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
'CurStyle = CurStyle And WS_THICKFRAME
CurStyle = CurStyle And WS_DLGFRAME

'Set the new style
NewStyle = SetWindowLong(frm.hWnd, GWL_STYLE, CurStyle)

frm.Left = XPos
frm.Top = YPos
'Debug.Print InvalidateRect(frm.hWnd, Rectangle, True)
'Debug.Print ValidateRect(frm.hWnd, Rectangle)
frm.Move XPos, YPos, frm.Width, frm.Height
frm.BorderStyle = 1
'MoveWindow frm.hWnd, XPos, YPos, frm.Width, frm.Height, False
frm.Refresh
End Sub


