' Get the latest version from https://github.com/blokeley/check_email_domains

' MIT licence https://github.com/blokeley/check_email_domains/blob/master/LICENSE

' Version 0.1

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim prompt As String
    Dim strMsg As String
    Dim strAddress As String
    Dim intAtIndex As Integer
    Dim colUniqueDomains As New Collection
    Dim domain As Variant
    Dim strMyDomain As String
    Dim strDomains As String
    Dim strDomain As String
    Dim strMyAddress As String

    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    strMyAddress = Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress 
    intAtIndex = Len(strMyAddress) - InStrRev(strMyAddress, "@")
    strMyDomain = Right(strMyAddress, intAtIndex)

    For Each recip In Item.Recipients

        Set pa = recip.PropertyAccessor
        strAddress = LCase(pa.GetProperty(PR_SMTP_ADDRESS)) 
        intAtIndex = Len(strAddress) - InStrRev(strAddress, "@")
        strDomain = Right(strAddress, intAtIndex)
    
        If strDomain <> strMyDomain Then
            On Error Resume Next
            colUniqueDomains.Add strDomain, strDomain
            On Error GoTo 0
        End If

    Next

    If colUniqueDomains.Count > 1 Then
    
        For Each domain In colUniqueDomains
            strDomains = strDomains & " " & domain
        Next

        prompt = "Do you wish to email recipients from the following domains? " & vbNewLine & strDomains

        If MsgBox(prompt, vbYesNo + vbExclamation + vbMsgBoxSetForeground, "Check Addresses") = vbNo Then
            Cancel = True
        End If
 
    End If
 
End Sub
