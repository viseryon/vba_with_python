Attribute VB_Name = "Module1"

Public Function ShellRun(sCmd As String) As String
    'https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba

    'Run a shell command, returning the output As a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)

    ' return whole output
    Dim path As String
    path = Application.ActiveWorkbook.Path
    ShellRun = oExec.StdOut.ReadAll
    MsgBox path,,ShellRun 

    ' Get output line by line
    ' While Not oOutput.AtEndOfStream
    '     sLine = oOutput.ReadLine
    '     If sLine <> "" Then s = s & sLine & vbCrLf
    '     Wend

    '     ShellRun = s


End Function


Public Sub ScriptRun()

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    Dim oExec As Object
    Dim oOutput As Object
    Dim sCmd As string

    Dim path As String
    path = Application.ActiveWorkbook.Path
    sCmd = "pythonw " + path + "\example.py"
    Set oExec = oShell.Exec(sCmd)

    Dim ShellRun As string
    ShellRun = oExec.StdOut.ReadAll
    MsgBox path,,ShellRun

    Worksheets("Arkusz1").Range("C11").value = Now

End Sub