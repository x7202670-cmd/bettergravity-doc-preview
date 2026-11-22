Option Explicit

Dim word, doc, options, oldLinksAtOpen, oldLinksAtPrint, oldFieldsAtPrint
Dim source, destination

On Error Resume Next
Set word = CreateObject("Word.Application")
If Err.Number = 0 Then
    If word.Visible Or word.Documents.Count <> 0 Or word.ProtectedViewWindows.Count <> 0 Then WScript.Quit 2
    word.AutomationSecurity = 3 ' Disable macros for this automation instance.
    word.DisplayAlerts = 0
    word.Visible = False
    Set options = word.Options
End If
If Err.Number <> 0 Then WScript.Quit 2

WScript.StdOut.WriteLine "ready"
Do Until WScript.StdIn.AtEndOfStream
    source = WScript.StdIn.ReadLine
    If WScript.StdIn.AtEndOfStream Then Exit Do
    destination = WScript.StdIn.ReadLine
    WScript.StdOut.WriteLine ExportDocument(source, destination)
Loop

If Not word.Visible And word.Documents.Count = 0 And word.ProtectedViewWindows.Count = 0 Then word.Quit 0
Set options = Nothing
Set word = Nothing

Function ExportDocument(source, destination)
    Dim capturedOptions
    On Error Resume Next
    Err.Clear
    ExportDocument = "error"
    If word.Visible Or word.Documents.Count <> 0 Or word.ProtectedViewWindows.Count <> 0 Then Exit Function
    oldLinksAtOpen = options.UpdateLinksAtOpen
    oldLinksAtPrint = options.UpdateLinksAtPrint
    oldFieldsAtPrint = options.UpdateFieldsAtPrint
    If Err.Number = 0 Then
        capturedOptions = True
        options.UpdateLinksAtOpen = False
        options.UpdateLinksAtPrint = False
        options.UpdateFieldsAtPrint = False
    End If
    If Err.Number = 0 Then
        ' Open a private copy, read-only, with no recent-file entry.
        Set doc = word.Documents.Open(source, False, True, False, " ")
    End If
    If Err.Number = 0 Then doc.ExportAsFixedFormat destination, 17

    If Err.Number = 0 Then ExportDocument = "ok"
    If IsObject(doc) Then doc.Close 0
    Set doc = Nothing
    If capturedOptions Then
        options.UpdateLinksAtOpen = oldLinksAtOpen
        options.UpdateLinksAtPrint = oldLinksAtPrint
        options.UpdateFieldsAtPrint = oldFieldsAtPrint
    End If
End Function
