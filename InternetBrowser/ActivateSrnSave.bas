Attribute VB_Name = "srnsave"
Option Explicit

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As _
  Long, ByVal lParam As Long) As Long

Public Const WM_SYSCOMMAND = &H112&
Public Const SC_SCREENSAVE = &HF140&

