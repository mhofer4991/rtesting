Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    Set Msg = item
    ' ******************
    ' do something here
    ' ******************
    MsgBox item.Subject

    'run command
    ' CreateObject("WScript.Shell").Exec("CMD /S /C dir /s /b directoryPath").StdOut.ReadAll
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec("CMD /S /C echo hallo")
    Set oOutput = oExec.StdOut
    
    Dim result As String
    result = oOutput.ReadAll
    
    MsgBox result
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile("C:\Users\Markus\Desktop\outlooktest\mail.txt", True, True)
    Fileout.Write item.Body
    Fileout.Close
    
    Shell ("cmd.exe /S /K " & "mkdir C:\Users\Markus\Desktop\outlooktest\neu")
    
    Dim objOutlook As Object
    Dim objMail As Object
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    With objMail
       .To = "mhofer4991@gmail.com"
       .Subject = "Test"
       .BodyFormat = olFormatHTML
       .HTMLBody = "<html><head><style>h1 { color: red; }</style><body><h1>Hello</h1><p>Test</p></body></html>"
       .Send        'Sendet die Email automatisch
    End With
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub
