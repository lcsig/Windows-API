Imports System.Diagnostics

Public Class Logger
    Private Shared _log As EventLog

    Public Shared EventSource As String = "LogApp"
    Public Shared EventLogName As String = "LogApp"

    Public Shared Sub Delete()
        If Logger.Exists() Then
            If _log IsNot Nothing Then
                _log.Clear()
                _log.Dispose()
                _log = Nothing
            End If
            EventLog.Delete(EventLogName)
        End If
    End Sub

    Public Shared Function Exists() As Boolean
        For Each l As EventLog In EventLog.GetEventLogs()
            If l.Log.ToLower() = EventLogName.ToLower() Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared ReadOnly Property Log() As EventLog
        Get
            If _log Is Nothing Then
                If EventLog.SourceExists(EventSource) Then
                    'verify the source is associated with correct log, if not delete and recreate
                    If EventLogName <> EventLog.LogNameFromSourceName(EventSource, KernelApi.GetComputerName()) Then
                        EventLog.DeleteEventSource(EventSource)
                    End If
                End If
                _log = New EventLog(EventSource, KernelApi.GetComputerName(), EventLogName)
            End If
            Return _log
        End Get
    End Property

    Public Shared Sub WriteEntry(ByVal message As String)
        Logger.WriteEntry(message, EventLogEntryType.Information)
    End Sub

    Public Shared Sub WriteEntry(ByVal message As String, ByVal type As EventLogEntryType)
        message = My.Application.Info.AssemblyName + ": " + message
        Try
            Log.WriteEntry(message, type)
        Catch ex As Exception
            Log.Clear()
        End Try
    End Sub

    Public Shared Sub WriteEntryWithDebugInfo(ByVal sessionID As Integer, ByVal title As String, ByVal msg As String, ByVal entryType As EventLogEntryType)

        Dim entry As String = title + " occurred on Session ID: " + CStr(sessionID)
        entry += vbCrLf + msg
        entry += vbCrLf + vbCrLf + "Current Session State: "

        Try
            If sessionID >= 0 And sessionID < 1000 Then
                Dim state As String = WtsApi.GetSessionInfo(sessionID, WtsApi.WTS_INFO_CLASS.WTSConnectState)
                entry += state + vbCrLf
            End If

            For Each s As WtsApi.WtsSession In WtsApi.GetSessions()
                entry += vbCrLf + CStr(s.SessionId) + "; " + s.ConnectionState.ToString() + "; " + s.StationName + "; " + s.UserName + "; " + s.UserDomain + "; " + s.UserSid
            Next
        Catch ex As Exception
            entry += vbCrLf + "Error while adding debug info: " + ex.Message
        End Try
        Logger.WriteEntry(entry, entryType)
    End Sub
End Class
