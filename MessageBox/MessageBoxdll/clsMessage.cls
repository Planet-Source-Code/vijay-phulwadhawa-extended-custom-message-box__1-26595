VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 08/04/2001 10:11:07 AM
'Project Name      :   msgdll
'Form/Class Name   :   clsMessage (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************


Dim frm As frmMessage

Public Enum AlertIcon
    Msg_Critical
    Msg_Exclamation
    Msg_Information
    Msg_Question
    Msg_Custom
End Enum

Public Enum ButtonConstants
    BS_Button1
    BS_Button2
    BS_Button3
    BS_Button4
End Enum

Public Enum MsgAlignment
    Msg_Left
    Msg_Right
    Msg_Center
End Enum

Public Function MessageBoxEx(Message As String, Optional MessageHead As String, _
Optional Icon As AlertIcon = Msg_Custom, Optional MsgForeColor As Long = vbBlack, _
Optional MsgHeadForeColor As Long = vbBlack, Optional MsgFontName As String = "Arial", _
Optional MsgHeadFontName As String = "Arial", Optional MsgFontSize As Byte = 10, _
Optional MsgHeadFontSize As Byte = 10, Optional MsgFontItalic As Boolean = False, _
Optional MsgHeadFontItalic As Boolean = False, Optional MsgFontBold As Boolean = False, _
Optional MsgHeadFontBold As Boolean = True, Optional MsgAlignment As MsgAlignment = Msg_Left, _
Optional MsgHeadAlignment As MsgAlignment = Msg_Left, Optional AutoCenterMsgBox As Boolean = True, _
Optional XPos As Long, Optional YPos As Long, Optional WaitForSecond As Byte, _
Optional CustomIcon As StdPicture, Optional AllowMove As Boolean = False) As ButtonConstants

Dim CurrentForeGroundWnd As Long
CurrentForeGroundWnd = GetForegroundWindow()
bAllowMove = AllowMove
Call SetAutoUnloadTime(WaitForSecond)
With frm
If .cmdButton.UBound = 0 Then .cmdButton(0).Caption = "&OK"

.lblMsgHead.Caption = MessageHead
.lblMsgHead.ForeColor = MsgHeadForeColor
.lblMsgHead.FontName = MsgHeadFontName
.lblMsgHead.FontSize = MsgHeadFontSize
.lblMsgHead.FontItalic = MsgHeadFontItalic
.lblMsgHead.FontBold = MsgHeadFontBold
.lblMsgHead.Alignment = MsgHeadAlignment


.lblMessage.Caption = Message
.lblMessage.ForeColor = MsgForeColor
.lblMessage.FontName = MsgFontName
.lblMessage.FontSize = MsgFontSize
.lblMessage.FontItalic = MsgFontItalic
.lblMessage.FontBold = MsgFontBold
.lblMessage.Alignment = MsgAlignment

Select Case Icon
Case Msg_Critical
    Set .imgIcon.Picture = LoadResPicture("Critical", vbResIcon)
Case Msg_Exclamation
    Set .imgIcon.Picture = LoadResPicture("Exclamation", vbResIcon)
Case Msg_Information
    Set .imgIcon.Picture = LoadResPicture("Information", vbResIcon)
Case Msg_Question
    Set .imgIcon.Picture = LoadResPicture("Question", vbResIcon)
Case Msg_Custom
    Set .imgIcon.Picture = CustomIcon
End Select

Call ResizeMessageBox(XPos, YPos)

If AutoCenterMsgBox Then
    CenterForm frm
End If


.Show vbModal

End With

MessageBoxEx = ButtonClicked
Call SetForegroundWindow(CurrentForeGroundWnd)

End Function

Public Sub AddButton(ButtonNumber As Byte, Caption As String, Optional Default As Boolean = False, Optional Cancel As Boolean = False)
On Error Resume Next
    If ButtonNumber <> 0 Then Load frm.cmdButton(ButtonNumber)
    frm.cmdButton(ButtonNumber).Caption = Caption
    If ButtonNumber <> 0 Then
        frm.cmdButton(ButtonNumber).Left = frm.cmdButton(ButtonNumber - 1).Left + frm.cmdButton(ButtonNumber - 1).Width + 100
    End If
    frm.cmdButton(ButtonNumber).Visible = True
    
    If Default = True Then
       frm.cmdButton(ButtonNumber).Default = True
       frm.cmdButton(ButtonNumber).Cancel = False
    Else
       frm.cmdButton(ButtonNumber).Default = False
    End If
    
    If Cancel = True Then
       frm.cmdButton(ButtonNumber).Cancel = True
       frm.cmdButton(ButtonNumber).Default = False
    Else
       frm.cmdButton(ButtonNumber).Cancel = False
    End If

End Sub

Private Sub SetAutoUnloadTime(WaitForSec As Byte)
    If WaitForSec <= 0 Then
    frm.Timer1.Enabled = False
    Exit Sub
    End If
    frm.Timer1.Interval = WaitForSec * 1000
        
End Sub

Private Sub ResizeMessageBox(XPos As Long, YPos As Long)
Dim I As Integer
With frm
    .Left = XPos
    .Top = YPos
    
    .lblMsgHead.Left = .imgIcon.Left + .imgIcon.Width + 200
    .lblMessage.Left = .lblMsgHead.Left
    
    If .lblMessage.Width > .lblMsgHead.Width Then
        .lblMsgHead.Width = .lblMessage.Width
    Else
        .lblMessage.Width = .lblMsgHead.Width
    End If
    
        
    If .cmdButton(0).Width * (.cmdButton.UBound + 1) < .lblMessage.Width Then
        .Width = .lblMsgHead.Width + 2 * .imgIcon.Width + 200
    Else
        .Width = .cmdButton(0).Width * (.cmdButton.UBound + 1) + .imgIcon.Width + 200
    End If
    
    .Height = .lblMsgHead.Height + .lblMessage.Height + 2 * .cmdButton(0).Height + 500
    .lblMessage.Top = .lblMsgHead.Top + .lblMsgHead.Height + 50
    .fraSeparator.Top = .lblMessage.Top + .lblMessage.Height + 400
    .fraSeparator.Left = 0
    .fraSeparator.Width = frm.Width - 20
    

'Center align buttons and align tops
Dim EqualDist As Long
EqualDist = (frm.Width - frm.cmdButton(0).Width * (frm.cmdButton.UBound + 1)) / (frm.cmdButton.UBound + 2)
    
    For I = .cmdButton.LBound To .cmdButton.UBound
        .cmdButton(I).Top = .fraSeparator.Top + .fraSeparator.Height + 100 '.Top - .Height - .cmdButton(i).Height
        .cmdButton(I).Left = (I + 1) * EqualDist + I * frm.cmdButton(0).Width
    Next I

End With
End Sub

Private Sub Class_Initialize()
Set frm = New frmMessage
End Sub

Private Sub Class_Terminate()
Set frm = Nothing
End Sub

Private Sub CenterForm(FormName As Object, Optional Child As Boolean, Optional MDIForm As Object)
On Error Resume Next
If IsMissing(Child) Or Child = False Then
FormName.Move (Screen.Width - FormName.Width) / 2, (Screen.Height - FormName.Height) / 2
Else
FormName.Left = MDIForm.Width / 2 - FormName.Width / 2
FormName.Top = MDIForm.Height / 2 - FormName.Height / 2 - 1000
End If
End Sub

