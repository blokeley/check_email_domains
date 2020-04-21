' Get the latest version from https://github.com/blokeley/check_email_domains

' MIT licence https://github.com/blokeley/check_email_domains/blob/master/LICENSE

' Version 0.1

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim prompt As String
    Dim strMsg As String
    Dim Address As String
    Dim intLen As Integer
    Dim uniquelist As New Collection
    Dim domain As Variant
    Dim list As String
    Dim strMyDomain As String
    Dim userAddress As String

    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    userAddress = Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress 
    intLen = Len(userAddress) - InStrRev(userAddress, "@")
    strMyDomain = Right(userAddress, intLen)

    Set recips = Item.Recipients
    
    For Each recip In recips

        Set pa = recip.PropertyAccessor
        Address = LCase(pa.GetProperty(PR_SMTP_ADDRESS)) 
        intLen = Len(Address) - InStrRev(Address, "@")
        str1 = Right(Address, intLen)
    
        If str1 <> strMyDomain Then
            On Error Resume Next
            uniquelist.Add str1, str1
            On Error GoTo 0
        End If

    Next

    If uniquelist.Count > 1 Then
    
        For Each domain In uniquelist
            list = list & " " & domain
        Next

        prompt = "Do you wish to email recipients from the following domains? " & vbNewLine & list

        If MsgBox(prompt, vbYesNo + vbExclamation + vbMsgBoxSetForeground, "Check Addresses") = vbNo Then
            Cancel = True
        End If
 
    End If
 
End Sub
