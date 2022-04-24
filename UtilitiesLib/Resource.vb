Imports System.IO
Imports System.Reflection

Public Class Resource

    Public Shared Function AsString(ByVal resourceName As String) As String
        Dim br As StreamReader = Nothing
        Try
            Dim assem As Assembly = Assembly.GetEntryAssembly()
            Dim source As Stream = assem.GetManifestResourceStream(resourceName)
            br = New StreamReader(source)
            Return br.ReadToEnd()
        Finally
            If br IsNot Nothing Then br.Close()
        End Try
    End Function

    Public Shared Sub StringToFile(ByVal str As String, ByVal filename As String)
        Dim bw As StreamWriter = Nothing
        Try
            bw = New StreamWriter(filename, False)
            bw.Write(str)
            bw.Flush()
        Finally
            If bw IsNot Nothing Then bw.Close()
        End Try
    End Sub

End Class
