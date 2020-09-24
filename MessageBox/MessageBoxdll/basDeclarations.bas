Attribute VB_Name = "basDeclarations"
Option Explicit

'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 08/04/2001 10:11:13 AM
'Project Name      :   msgdll
'Form/Class Name   :   basDeclarations (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************

'FORM MOVE CONSTANTS
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1

Public Const GWL_STYLE = &HFFF0


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public ButtonClicked As Byte
Public bAllowMove As Boolean

Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "User32" () As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
