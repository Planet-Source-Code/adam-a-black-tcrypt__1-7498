Attribute VB_Name = "Module1"
Public fMainForm As frmMain


Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash


    fMainForm.Show
End Sub

Public Sub Encrypt(Text As String, Output As RichTextBox)
    On Error GoTo Break
    Randomize
    Dim Char() As String
    Dim si As Long
    Dim Out As String
    Dim iLen As Long
    Dim Rand
    frmCrypt.Show
    frmCrypt.lblMethod.Caption = "Encrypting"
    Text = StrReverse(Text)
    iLen = Len(Text)
    Rand = Int(150 * Rnd) + 1
    ReDim Char(1 To iLen)
    Out = "ø"
    frmCrypt.PB.Max = iLen
    frmCrypt.Refresh
    For i = 1 To iLen
        Char(i) = Chr(Asc(Mid(Text, i, 1)) Xor Rand)
        Out = Out & Char(i)
        frmCrypt.PB.Value = i
    Next
    Output.Text = Out & StrReverse(Rand) & Len(Rand) * 2 + 1
    frmCrypt.PB.Value = 0
    Unload frmCrypt
Exit Sub
Break:
MsgBox Err.Description, vbOKOnly + vbExclamation, "Error " & Err.Number
frmCrypt.PB.Value = 0
Unload frmCrypt
End Sub

Public Sub Decrypt(Text As String, Output As RichTextBox)
    On Error GoTo Break
    Dim Char() As String
    Dim i As Long
    Dim Out As String
    Dim iLen As Long
    Dim Rand As Integer
    Dim RLen As Single
    RLen = Right(Text, 1) - 1: RLen = RLen / 2
    Rand = StrReverse(Mid(Text, Len(Text) - RLen, RLen))
    Text = Mid(Text, 2, Len(Text) - RLen - 2)
    Text = StrReverse(Text)
    iLen = Len(Text): ReDim Char(1 To iLen)
    frmCrypt.Show
    frmCrypt.lblMethod.Caption = "Decrypting"
    frmCrypt.PB.Max = iLen
    frmCrypt.Refresh
    For i = 1 To iLen
        Char(i) = Chr(Asc(Mid(Text, i, 1)) Xor Rand)
        Out = Out & Char(i)
        frmCrypt.PB.Value = i
    Next
    Output.Text = Out
    frmCrypt.PB.Value = 0
    Unload frmCrypt
Exit Sub
Break:
MsgBox Err.Description, vbOKOnly + vbExclamation, "Error " & Err.Number
frmCrypt.PB.Value = 0
Unload frmCrypt
End Sub


