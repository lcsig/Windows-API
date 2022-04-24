Imports System.IO

Public Class SubInACL
    Public Shared Sub GrantFullAccess(ByVal fullPath As String, ByVal domainUser As String)

        domainUser = "BUILTIN\Administrators"

        Dim p As New Process
        p.StartInfo.FileName = Path.Combine(Environment.SystemDirectory, "Subinacl.exe")
        p.StartInfo.UseShellExecute = False
        p.StartInfo.CreateNoWindow = True
        p.StartInfo.Arguments = _
            "/subdirectories """ + fullPath + """ /grant=" + domainUser + "=F"
        p.Start()
        p.WaitForExit()
        If p.ExitCode <> 0 Then
            Throw New Exception("Subinacl error: " + CStr(p.ExitCode))
        End If
    End Sub
End Class
