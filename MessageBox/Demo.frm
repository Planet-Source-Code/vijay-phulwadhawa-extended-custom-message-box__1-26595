VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Message Box Creator By Vijay Phulwadhawa"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show MessageBox At 200,200"
      Height          =   420
      Left            =   1995
      TabIndex        =   2
      Top             =   1215
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Auto Unload MessageBox"
      Height          =   420
      Left            =   1995
      TabIndex        =   1
      Top             =   705
      Width           =   2610
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   150
      Picture         =   "Demo.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    addbutton 0, "&OK"
    addbutton 1, "&Cancel"
    addbutton 2, "&View Log"
    addbutton 3, "&Custom", True
    addbutton 4, "&Fifth Button"
    Select Case MessageBoxEx("This application is created in Visual Basic 6. " & vbCrLf & "If you have any queries regarding this program " & vbCrLf & "please feel free to mail me at vijaycg44@hotmail.com.!" & vbCrLf & "Read before the message box unloads automatically ! :-)", "This is The Head Of A Custom Message Box !", Msg_Custom, vbRed, , "arial", "Tahoma", 12, 12, False, False, False, True, Msg_Center, Msg_Center, True, , , 7, Picture1.Picture)
    Case BS_Button1
        MsgBox "You clicked on first button"
    Case BS_Button2
        MsgBox "You clicked on second button"
    Case BS_Button3
        MsgBox "You clicked on third button"
    Case BS_Button4
        MsgBox "You clicked on fourth button"
    Case 4
        MsgBox "You clicked on Fifth button"
    End Select
End Sub

Private Sub Command2_Click()
    MessageBoxEx "This is The Matter !", "This is The Head", Msg_Exclamation, vbRed, vbBlue, "arial", "Tahoma", 10, 10, False, False, False, True, Msg_Center, Msg_Center, False, 200, 200
End Sub

