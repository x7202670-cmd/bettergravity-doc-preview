Option Explicit

Dim powerpoint, owned, source, destination

On Error Resume Next
Set powerpoint = GetObject(, "PowerPoint.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set powerpoint = CreateObject("PowerPoint.Application")
    owned = (Err.Number = 0 And powerpoint.Presentations.Count = 0 And powerpoint.Windows.Count = 0)
End If
If Err.Number <> 0 Then WScript.Quit 2

WScript.StdOut.WriteLine "ready"
Do Until WScript.StdIn.AtEndOfStream
    source = WScript.StdIn.ReadLine
    If WScript.StdIn.AtEndOfStream Then Exit Do
    destination = WScript.StdIn.ReadLine
    WScript.StdOut.WriteLine ExportPresentation(source, destination)
Loop

If owned And powerpoint.Presentations.Count = 0 And powerpoint.Windows.Count = 0 Then powerpoint.Quit
Set powerpoint = Nothing

Function ExportPresentation(source, destination)
    Dim presentation, oldSecurity
    On Error Resume Next
    Err.Clear
    ExportPresentation = "error"

    oldSecurity = powerpoint.AutomationSecurity
    If Err.Number <> 0 Then Exit Function
    powerpoint.AutomationSecurity = 3 ' Force-disable macros during the read-only open.
    If Err.Number = 0 Then Set presentation = powerpoint.Presentations.Open(source, -1, 0, 0)
    Dim openError
    openError = Err.Number
    Err.Clear
    powerpoint.AutomationSecurity = oldSecurity
    If openError = 0 And Err.Number = 0 Then
        ' Print-quality PDF, one page per slide, including hidden slides.
        presentation.ExportAsFixedFormat destination, 2, 2, 0, 1, 1, -1, Nothing, 1
        If Err.Number = 0 Then ExportPresentation = "ok"
    End If
    If IsObject(presentation) Then presentation.Close
    Set presentation = Nothing
End Function
