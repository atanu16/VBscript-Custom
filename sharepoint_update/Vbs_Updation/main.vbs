Option Explicit

'=== Input Variables ===
Dim botName, platform, username, machine, runDate, startTime, endTime, status, region
botName   = "Portugal"
platform  = "BluePrism"
username  = "BTNMO.AA"
machine   = "MININT-R37IK"
runDate   = "05/02/2025"
startTime = "2:00 PM"
endTime   = "4:00 PM"
status    = "InProgress" ' OR "Completed"
region    = "Portugal"

'=== SharePoint Setup ===
Dim listName, sharepointURL, soapAction, xmlhttp, soapEnvelope, existingItemID
listName = "Temp"
sharepointURL = "https://Yoursite.com/sites/Dev16/_vti_bin/Lists.asmx"
soapAction = "http://schemas.microsoft.com/sharepoint/soap/"

'=== Prompt for credentials if needed ===
Dim spUser, spPass
spUser = InputBox("Enter your SharePoint username (e.g., user@domain.com):", "SharePoint Login")
spPass = InputBox("Enter your SharePoint password:", "SharePoint Login")

'=== Log File Setup ===
Dim logFile, fso, log
Set fso = CreateObject("Scripting.FileSystemObject")
logFile = "sharepoint_log.txt"
Set log = fso.OpenTextFile(logFile, 8, True) ' Append mode

Call LogWrite("==== SharePoint Sync Started ====")

'=== Step 1: Check existing item ===
Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")
soapEnvelope = _
"<?xml version='1.0' encoding='utf-8'?>" & _
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " & _
"xmlns:xsd='http://www.w3.org/2001/XMLSchema' " & _
"xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
"<soap:Body>" & _
"<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
"<listName>" & listName & "</listName>" & _
"<query><Query><Where><Eq>" & _
"<FieldRef Name='Title'/><Value Type='Text'>" & botName & "</Value>" & _
"</Eq></Where></Query></query>" & _
"<rowLimit>1</rowLimit>" & _
"</GetListItems>" & _
"</soap:Body></soap:Envelope>"

SendRequest xmlhttp, sharepointURL, soapAction & "GetListItems", soapEnvelope, spUser, spPass

Dim xmlDoc, nodes, node, statusValue
Set xmlDoc = CreateObject("MSXML2.DOMDocument")
xmlDoc.async = False
xmlDoc.LoadXML(xmlhttp.responseText)

Set nodes = xmlDoc.getElementsByTagName("z:row")

If nodes.length > 0 Then
    Set node = nodes.Item(0)
    existingItemID = node.getAttribute("ows_ID")
    statusValue = node.getAttribute("ows_Status")

    If LCase(statusValue) = "inprogress" Then
        Call LogWrite("Matching bot found with InProgress status. Updating item ID " & existingItemID)
        UpdateOrInsertItem "Update", existingItemID
    Else
        Call LogWrite("Matching bot found, but status is '" & statusValue & "'. Creating new item.")
        UpdateOrInsertItem "New", ""
    End If
Else
    Call LogWrite("No matching bot found. Creating new item.")
    UpdateOrInsertItem "New", ""
End If

log.Close
Set log = Nothing
Set fso = Nothing
Set xmlhttp = Nothing
Set xmlDoc = Nothing

MsgBox "✅ Script completed. Check 'sharepoint_log.txt' for details."

'==================== Subroutines ====================

Sub UpdateOrInsertItem(cmd, itemID)
    Dim req, env
    Set req = CreateObject("MSXML2.XMLHTTP.6.0")
    env = "<?xml version='1.0' encoding='utf-8'?>" & _
    "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " & _
    "xmlns:xsd='http://www.w3.org/2001/XMLSchema' " & _
    "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
    "<soap:Body><UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
    "<listName>" & listName & "</listName>" & _
    "<updates><Batch OnError='Continue'>" & _
    "<Method ID='1' Cmd='" & cmd & "'>" & _
    (IIf(cmd = "Update", "<Field Name='ID'>" & itemID & "</Field>", "")) & _
    "<Field Name='Title'>" & botName & "</Field>" & _
    "<Field Name='Platform'>" & platform & "</Field>" & _
    "<Field Name='Username'>" & username & "</Field>" & _
    "<Field Name='Machine'>" & machine & "</Field>" & _
    "<Field Name='Date'>" & runDate & "</Field>" & _
    "<Field Name='StartTime'>" & startTime & "</Field>" & _
    "<Field Name='EndTime'>" & endTime & "</Field>" & _
    "<Field Name='Status'>" & status & "</Field>" & _
    "<Field Name='Region'>" & region & "</Field>" & _
    "</Method></Batch></updates>" & _
    "</UpdateListItems></soap:Body></soap:Envelope>"

    SendRequest req, sharepointURL, soapAction & "UpdateListItems", env, spUser, spPass
    Call LogWrite("✔ " & cmd & " action completed with HTTP status: " & req.status)
End Sub

Sub SendRequest(obj, url, action, body, user, pass)
    obj.Open "POST", url, False
    obj.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    obj.setRequestHeader "SOAPAction", action
    obj.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & pass)
    obj.Send body
End Sub

Function Base64Encode(text)
    Dim xml, node
    Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = xml.createElement("base64")
    node.dataType = "bin.base64"
    node.nodeTypedValue = Stream_StringToBinary(text)
    Base64Encode = Replace(node.text, vbLf, "")
End Function

Function Stream_StringToBinary(text)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1 ' Binary
    Stream_StringToBinary = stream.Read
    stream.Close
End Function

Sub LogWrite(msg)
    log.WriteLine Now & " - " & msg
End Sub
