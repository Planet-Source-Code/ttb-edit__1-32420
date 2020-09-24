VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "edit-"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "edit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".mp3"
      DialogTitle     =   "text file"
      FileName        =   "*.txt"
      Filter          =   "text files|*.txt|"
      MaxFileSize     =   261
   End
   Begin VB.Menu file 
      Caption         =   "file"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu new 
         Caption         =   "new"
      End
      Begin VB.Menu open 
         Caption         =   "open"
      End
      Begin VB.Menu save 
         Caption         =   "save"
      End
      Begin VB.Menu saveas 
         Caption         =   "save as"
      End
      Begin VB.Menu print 
         Caption         =   "print"
      End
      Begin VB.Menu tetromakintarjetas 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "edit"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu undo 
         Caption         =   "undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu copy 
         Caption         =   "copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut 
         Caption         =   "cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste 
         Caption         =   "paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu selectall 
         Caption         =   "select all"
      End
      Begin VB.Menu changefont 
         Caption         =   "change font"
      End
   End
   Begin VB.Menu document 
      Caption         =   "document"
      Begin VB.Menu find 
         Caption         =   "find"
      End
      Begin VB.Menu findnext 
         Caption         =   "find next"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub changefont_Click()
CommonDialog1.Flags = &H3
CommonDialog1.ShowFont
With t.Font
.Bold = CommonDialog1.FontBold
.Italic = CommonDialog1.FontItalic
.Name = CommonDialog1.FontName
.Size = CommonDialog1.FontSize
.Strikethrough = CommonDialog1.FontStrikethru
.Underline = CommonDialog1.FontUnderline
End With
End Sub

Private Sub copy_Click()
SendMessage t.hwnd, WM_COPY, VK_CONTROL, 0
End Sub

Private Sub cut_Click()
SendMessage t.hwnd, WM_CUT, VK_CONTROL, 0
End Sub



Private Sub find_Click()
If start = 0 Then start = 1
Form2.Show
End Sub

Private Sub findnext_Click()
If Form2.Text1 = "" Then Exit Sub
If InStr(start, t, Form2.Text1) = 0 Then Exit Sub
Form1.t.SelStart = InStr(start, t, Form2.Text1) - 1
Form1.t.SelLength = Len(Form2.Text1)
start = InStr(start, t, Form2.Text1) + Len(Form2.Text1)
End Sub

Private Sub Form_Load()
toask = True
End Sub

Private Sub Form_Resize()
t.Height = Form1.Height - 700
t.Width = Form1.Width - 125
End Sub

Private Sub new_Click()
If edited = False Then
fn = ""
Form1.Caption = "edit-"
t.Text = ""
toask = True
End If
If edited = True Then
save_Click
End If

End Sub

Private Sub open_Click()
If edited = False Then
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName For Input As #2
blah = Input(FileLen(CommonDialog1.FileName), #2)
Close #2
t.Text = blah
fn = CommonDialog1.FileName
toask = False
justopened = True
edited = False
Exit Sub
End If
If edited = True Then
tosave = MsgBox("would you like to save the changed you made to " & fn & "?", vbYesNoCancel)
Select Case tosave
Case vbCancel
Exit Sub
Case vbNo
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName For Input As #2
blah = Input(FileLen(CommonDialog1.FileName), #2)
Close #2
t.Text = blah
fn = CommonDialog1.FileName
toask = False
Case vbYes
save_Click
End Select
justopened = True
End If
Form1.Caption = "edit-" & fn
End Sub



Private Sub paste_Click()
SendMessage t.hwnd, WM_PASTE, VK_CONTROL, 0
End Sub

Private Sub print_Click()
With Printer
.FontName = t.FontName
.FontBold = t.FontBold
.FontItalic = t.FontItalic
.FontSize = t.FontSize
.FontStrikethru = t.FontStrikethru
.FontUnderline = t.FontUnderline
.ForeColor = t.ForeColor
End With
Printer.Print t

End Sub

Private Sub printersetup_Click()
CommonDialog1.ShowPrinter
Printer.Copies = CommonDialog1.Copies
MsgBox Printer.Copies
End Sub

Private Sub save_Click()
If toask = False Then GoTo savefile
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
fn = CommonDialog1.FileName
GoTo savefile
savefile:
Open fn For Output As #1
Print #1, t.Text
Close #1
edited = False
toask = False
Form1.Caption = "edit-" & fn
End Sub

Private Sub saveas_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
fn = CommonDialog1.FileName
Open fn For Output As #1
Print #1, t.Text
Close #1
edited = False
toask = False
Form1.Caption = "edit-" & fn
End Sub

Private Sub selectall_Click()
t.SelStart = 0
t.SelLength = Len(t.Text)
End Sub

Private Sub t_Change()
If justopened = True Then
justopened = False
edited = False
Exit Sub
End If
edited = True
End Sub

Private Sub tetromakintarjetas_Click()
If edited = True Then
tosave = MsgBox("would you like to save the changed you made to " & fn & "?", vbYesNoCancel)
Select Case tosave
Case vbCancel
Exit Sub
Case vbNo
End
Case vbYes
save_Click
End Select
End If
If edit = False Then
End
End If
End Sub

Private Sub undo_Click()
SendMessage t.hwnd, WM_UNDO, VK_CONTROL, 0
End Sub

