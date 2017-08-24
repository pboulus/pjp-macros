Sub Extract_Emails_To_Excel()

#Extracts email addresses from all emails in a folder to an Excel spreadsheet

Dim olApp As Outlook.Application
Dim olExp As Outlook.Explorer
Dim olFolder As Outlook.MAPIFolder
Dim obj As Object
Dim stremBody As String
Dim stremSubject As String
Dim i As Long
Dim x As Long
Dim count As Long
Dim regEx As Object
Set regEx = CreateObject("VBScript.RegExp")
regEx.Global = True
Dim xlApp As Object 'Excel.Application
Dim xlwkbk As Object 'Excel.Workbook
Dim xlwksht As Object 'Excel.Worksheet
Dim xlRng As Object 'Excel.Range
Dim badAddresses As Variant
Set olApp = Outlook.Application
Set olExp = olApp.ActiveExplorer
Dim olMatches As Object
Set olFolder = olExp.CurrentFolder
'Open Excel
Set xlApp = GetExcelApp
xlApp.Visible = True
If xlApp Is Nothing Then GoTo ExitProc
Set xlwkbk = xlApp.Workbooks.Add
Set xlwksht = xlwkbk.Sheets(1)
Set xlRng = xlwksht.Range("A1")
xlRng.Value = "Email addresses"
'Set count of email objects
count = olFolder.Items.count
'counter for excel sheet
i = 0
'counter for emails
x = 1
For Each obj In olFolder.Items
xlApp.StatusBar = x & " of " & count & " emails completed"
stremBody = obj.Body
stremSubject = obj.Subject
regEx.Pattern = "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b"
regEx.IgnoreCase = True
regEx.MultiLine = True
Set olMatches = regEx.Execute(stremBody)
For Each match In olMatches
xlwksht.Cells(i + 2, 1).Value = match
i = i + 1
Next match
x = x + 1
Next obj
xlApp.ScreenUpdating = True
MsgBox ("All Email addresses are done being extracted")
ExitProc:
Set xlRng = Nothing
Set xlwksht = Nothing
Set xlwkbk = Nothing
Set xlApp = Nothing
Set emItm = Nothing
Set olFolder = Nothing
Set olNS = Nothing
Set olApp = Nothing
End Sub
Function GetExcelApp() As Object
' always create new instance
On Error Resume Next
Set GetExcelApp = CreateObject("Excel.Application")
On Error GoTo 0
End Function
