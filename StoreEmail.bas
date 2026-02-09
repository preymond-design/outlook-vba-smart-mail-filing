Attribute VB_Name = "StoreEmail"
Public Sub StoreMail_UserForm()
    Debug.Print "Running StoreMail_UserForm"

    Dim uf As FrmStoreEmail
    Set uf = New FrmStoreEmail

    Dim cnt As Long
    cnt = uf.PrepareChoices

    If cnt = 0 Then
        MsgBox "No eligible folders found or no item selected.", vbExclamation
        Unload uf
        Exit Sub
    End If

    If cnt = 1 Then
        uf.MoveToPath uf.GetSingleChoicePath
        ' Optionnel: Confirmation's MsgBox
        ' MsgBox "Moved email(s) to " & uf.GetSingleChoicePath, vbInformation
        Unload uf
        Exit Sub
    End If

    uf.Show
End Sub



