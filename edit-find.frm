VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "find"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "find"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "match case"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "find:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
If InStr(start, Form1.t, Text1) = 0 Then Exit Sub
Form1.t.SelStart = InStr(start, Form1.t, Text1) - 1
Form1.t.SelLength = Len(Text1)
start = InStr(start, Form1.t, Text1) + Len(Text1)
End Sub

