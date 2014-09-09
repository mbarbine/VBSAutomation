
Sub SaveAllConstrained(objitem As MailItem)
    
    Dim objAttachments As Outlook.Attachments
    Dim strName, strLocation As String
    Dim dblCount, dblLoop As Double

    strLocation = ""

    
On Error GoTo ExitSub
If objitem.Class = olMail Then
Set objAttachments = objitem.Attachments
dblCount = objAttachments.Count
If dblCount <= 0 Then
GoTo 100
End If
For dblLoop = 1 To dblCount
strName = MailItem.Subject(dblLoop).FileName
strName = strLocation & strName & ".PDF"
objAttachments.Item(dblLoop).SaveAsFile strName
Next dblLoop
objitem.Delete
End If
100
ExitSub:
Set objAttachments = Nothing
Set objOutlook = Nothing
End Sub

