Attribute VB_Name = "Module1"

Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output As a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    ShellRun = oExec.StdOut.ReadAll
    MsgBox ShellRun,,"title" 


End Function


Public Sub ScriptRun()

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    Dim oExec As Object
    Dim oOutput As Object
    Dim sCmd As string

    sCmd = "pythonw C:\Users\alans\REPOS\vba_with_python\capture_std\example.py"
    Set oExec = oShell.Exec(sCmd)

    Dim ShellRun As string
    ShellRun = oExec.StdOut.ReadAll
    MsgBox ShellRun,,"title" 

end Sub