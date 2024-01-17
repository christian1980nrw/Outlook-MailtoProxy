Set objArgs = WScript.Arguments

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set tempFolder = objFSO.GetSpecialFolder(2)
logFilePath = tempFolder & "\\MailtoProcessingLog.txt"
Set logFile = objFSO.CreateTextFile(logFilePath, True)

If objArgs.Count > 0 Then
    Dim mailto, email, subject, body, attachments
    mailto = objArgs(0)

    logFile.WriteLine("Original mailto string: " & mailto)

    email = ParseAndDecodeMailto(mailto, "mailto:")
    subject = ParseAndDecodeMailto(mailto, "subject=")
    body = ParseAndDecodeMailto(mailto, "body=")
    attachments = GetAttachments(mailto)

    logFile.WriteLine("Email: " & email)
    logFile.WriteLine("Subject: " & subject)
    logFile.WriteLine("Body: " & body)
    logFile.WriteLine("Attachments: " & attachments)

    SendEmail email, subject, body, attachments, logFile

    logFile.Close
Else
    WScript.Echo "No arguments provided. Please provide a mailto link."
End If

Function SendEmail(email, subject, body, attachments, logFile)
    Dim signature, adError
    adError = ""
    signature = "Signature will be empty."

	On Error Resume Next
	Set objSysInfo = CreateObject("ADSystemInfo")
	strUser = objSysInfo.UserName
	Set objUser = GetObject("LDAP://" & strUser)
	If Err.Number <> 0 Then
		adError = "Error accessing Microsoft Active Directory: " & Err.Number & " - " & Err.Description
		logFile.WriteLine("LDAP Access Error: " & adError)
		Err.Clear
	Else
        strGiven = objUser.givenName
        strSurname = objUser.sn
        strAddress1 = objUser.streetaddress
        strAddress2 = objUser.l
        strPostcode = objUser.postalcode
        strDepartment = objUser.Department
        strCompany = objUser.Company
        strPhone = objUser.telephoneNumber
        strEmail = objUser.mail

        signature = "Mit freundlichen Grüßen / Best regards" & vbCrLf & vbCrLf & _
                    strGiven & " " & strSurname & vbCrLf & _
                    "- " & strDepartment & " -" & vbCrLf & _
                    "Tel. " & strPhone & vbCrLf & _
                    strEmail & vbCrLf & vbCrLf & _
                    strCompany & vbCrLf & _
					strAddress1 & ", " & strPostcode & " " & strAddress2 & vbCrLf
End If
On Error GoTo 0

Set OutApp = CreateObject("Outlook.Application")
Set oEmailItem = OutApp.CreateItem(0)

oEmailItem.To = email
oEmailItem.Subject = subject

Dim attachSuccess, attachFailMsg
attachFailMsg = ""
attachSuccess = True

Dim i, attachArray
attachArray = Split(attachments, "|")
For i = 0 To UBound(attachArray)
    If Len(attachArray(i)) > 0 Then
        logFile.WriteLine("Attempting to attach file: " & attachArray(i))
        If objFSO.FileExists(attachArray(i)) Then
            oEmailItem.Attachments.Add(attachArray(i))
            logFile.WriteLine("Attachment successful: " & attachArray(i))
        Else
            attachFailMsg = attachFailMsg & "Failed to attach file (does not exist or inaccessible): " & attachArray(i) & vbCrLf
            logFile.WriteLine("Attachment failed: " & attachArray(i))
            attachSuccess = False
        End If
    End If
Next

Dim finalBody
finalBody = body
If Not attachSuccess Then
    finalBody = finalBody & vbCrLf & vbCrLf & "Attachment Error:" & vbCrLf & attachFailMsg
End If
If adError <> "" Then
    finalBody = finalBody & vbCrLf & vbCrLf & "Active Directory Error: " & adError
End If
oEmailItem.Body = finalBody & vbCrLf & vbCrLf & signature
oEmailItem.Display

End Function

Function ParseAndDecodeMailto(mailto, key)
    Dim parsedValue, keyPos, endPos
    parsedValue = ""
    If key = "mailto:" Then
        keyPos = InStr(mailto, "?")
        If keyPos = 0 Then keyPos = Len(mailto) + 1
        parsedValue = Mid(mailto, Len("mailto:") + 1, keyPos - Len("mailto:") - 1)
    Else
        keyPos = InStr(mailto, key)
        If keyPos > 0 Then
            keyPos = keyPos + Len(key)
            endPos = InStr(keyPos, mailto, "&")
            If endPos = 0 Then endPos = Len(mailto) + 1
            parsedValue = Mid(mailto, keyPos, endPos - keyPos)
        End If
    End If
    parsedValue = SimpleURLDecode(parsedValue)
    ParseAndDecodeMailto = parsedValue
End Function

Function SimpleURLDecode(str)
    Dim decodedStr
    decodedStr = str
    decodedStr = Replace(decodedStr, "%20", " ")
    decodedStr = Replace(decodedStr, "%40", "@")
    decodedStr = Replace(decodedStr, "%3A", ":")
    decodedStr = Replace(decodedStr, "%2C", ",")
    decodedStr = Replace(decodedStr, "%3B", ";")
    decodedStr = Replace(decodedStr, "%0D%0A", "")
    decodedStr = Replace(decodedStr, "%0D", vbCrLf)
    decodedStr = Replace(decodedStr, "%0A", "")

    SimpleURLDecode = decodedStr
End Function


Function GetAttachments(mailtoParams)
    Dim startPos, nextAttachmentStart, attachmentString, attachmentPath
    attachmentString = ""
    startPos = InStr(mailtoParams, "Attach=")

    While startPos > 0
        startPos = startPos + Len("Attach=")
        nextAttachmentStart = InStr(startPos, mailtoParams, "&Attach=")
        If nextAttachmentStart = 0 Then
            nextAttachmentStart = Len(mailtoParams) + 1
        End If

        attachmentPath = Mid(mailtoParams, startPos, nextAttachmentStart - startPos)
        attachmentPath = SimpleURLDecode(attachmentPath)

        attachmentPath = Replace(attachmentPath, "//", "\\")
        attachmentPath = Replace(attachmentPath, "/", "\")

        If InStr(attachmentPath, " ") > 0 Then
            attachmentPath = """" & attachmentPath & """"
        End If

        If Len(attachmentString) > 0 Then
            attachmentString = attachmentString & "|"
        End If
        attachmentString = attachmentString & attachmentPath

        startPos = InStr(nextAttachmentStart, mailtoParams, "&Attach=")
        If startPos > 0 Then
            startPos = startPos + 1
        End If
    Wend

    GetAttachments = attachmentString
End Function
