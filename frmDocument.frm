VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":0442
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub

Private Sub rtfText_Change()
    CS
End Sub

Public Sub CS()
    Dim Char As String
    Char = Left(rtfText.Text, 1)
    If rtfText.Text = vbNullString Then fMainForm.tbCrypt.Buttons(1).Enabled = False: fMainForm.tbCrypt.Buttons(2).Enabled = False: Exit Sub
    If Char = "ø" Then
        fMainForm.tbCrypt.Buttons(1).Enabled = False
        fMainForm.tbCrypt.Buttons(2).Enabled = True
        fMainForm.mnuEncrypt.Enabled = False
        fMainForm.mnuDecrypt.Enabled = True
    Else
        fMainForm.tbCrypt.Buttons(1).Enabled = True
        fMainForm.tbCrypt.Buttons(2).Enabled = False
        fMainForm.mnuEncrypt.Enabled = True
        fMainForm.mnuDecrypt.Enabled = False
    End If
End Sub
