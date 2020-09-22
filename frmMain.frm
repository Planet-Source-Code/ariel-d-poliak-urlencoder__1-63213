VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "URL Encoder"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOutput 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmMain.frx":0000
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CheckBox chkEntitiesOnly 
      Alignment       =   1  'Right Justify
      Caption         =   "Special Characters Only"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   2040
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":0015
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Row_Separator             As Single
Private MIN_HEIGHT                As Single
Private MIN_WIDTH                 As Single

Private Sub cmdProcess_Click()

  txtOutput.Text = URLEncode(txtInput.Text, True, CBool(chkEntitiesOnly.Value))

End Sub

Private Sub Form_Load()

  Row_Separator = cmdProcess.Top - txtInput.Height - txtInput.Top
  MIN_HEIGHT = Me.Height
  MIN_WIDTH = Me.Width

End Sub

Private Sub Form_Resize()

  If Me.Height < MIN_HEIGHT Or Me.Width < MIN_WIDTH Then
    If Me.Height < MIN_HEIGHT Then
      Me.Height = MIN_HEIGHT
    End If
    If Me.Width < MIN_WIDTH Then
      Me.Width = MIN_WIDTH
    End If
    Form_Resize
    Exit Sub
  End If
  With txtInput
    .Top = Row_Separator
    .Width = Me.ScaleWidth - 2 * .Left
    .Height = (Me.ScaleHeight - cmdProcess.Height) / 2 - Row_Separator
    cmdProcess.Top = .Top + .Height + Row_Separator
  End With 'txtInput
  chkEntitiesOnly.Top = cmdProcess.Top
  chkEntitiesOnly.Left = Me.ScaleWidth - chkEntitiesOnly.Width - 2 * txtInput.Left
  txtOutput.Top = cmdProcess.Top + cmdProcess.Height + Row_Separator
  txtOutput.Width = txtInput.Width
  txtOutput.Height = Me.ScaleHeight - txtOutput.Top - Row_Separator

End Sub

Public Function URLEncode(StringToEncode As String, Optional UsePlusRatherThanHexForSpace As Boolean = False, Optional encEntitiesOnly As Boolean = True) As String
'modified from "URL Encoder and Decoder for VB" by Igor <http://www.freevbcode.com/AuthorInfo.asp?AuthorID=1098>
'original source located at <http://www.freevbcode.com/ShowCode.asp?ID=1512>
Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case Asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      If encEntitiesOnly Then
        TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
      Else
        TempAns = TempAns & "%" & Format(Hex(Asc(Mid(StringToEncode, CurChr, 1))), "00")
      End If
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & Format(Hex(Asc(Mid(StringToEncode, CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function


