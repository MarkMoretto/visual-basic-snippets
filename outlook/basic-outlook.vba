

Public Const incr_1 As Long = 1

Public Function SqlFilter(keyword As String)
'' Create query filter for Outlook objects
    SqlFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%" & keyword & "%'"
End Function

Public Function PrintEmail()
    Dim olApp As Outlook.Application
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = Outlook.Application
    End If

    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItems As Outlook.Items

    '### Keyword to match in email subject
    Dim sql_filter As String
    sql_filter = SqlFilter("COVID") ' Looking for COVID emails in this case

    '### Outlook application objects
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    Set olItems = olFolder.Items.Restrict(sql_filter)
    
    If olItems.count > 0 Then
        For Each itm In olItems
            ' Print name of sender, sent date, and subject to console.
            Debug.Print itm.SenderName & ", " & itm.ReceivedTime & ", " & itm.Subject
        Next
    Else
        Call MsgBox("No items found!", vbOKOnly, "Outlook Mail Check")
    End If
    
    ' olApp.Quit ' Note: This will close Outlook if it's open.
    Set olApp = Nothing
    Set fdict = Nothing
End Function
