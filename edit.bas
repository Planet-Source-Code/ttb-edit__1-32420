Attribute VB_Name = "Module1"
Public Const WM_CUT = &H300

Public Const WM_COPY = &H301

Public Const WM_PASTE = &H302

Public Const WM_UNDO = &H304

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Public Const VK_CONTROL = &H11
Global start As Long
Global toask As Boolean
Global edited As Boolean
Global fn As String
Global justopened As Boolean

