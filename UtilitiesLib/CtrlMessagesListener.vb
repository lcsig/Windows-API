Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class CtrlMessagesListener

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SetConsoleCtrlHandler( _
        ByVal handler As MessageDelegate, _
        ByVal add As Boolean) As Integer
    End Function

    Private Const DESKTOP_READOBJECTS As Long = &H1

    ' A delegate type to be used as the handler routine 
    ' for SetConsoleCtrlHandler.
    Delegate Function MessageDelegate(ByVal type As ControlTypes) As Boolean

    ' An enumerated type for the control messages 
    ' sent to the handler routine.
    Public Enum ControlTypes
        CTRL_C_EVENT = 0
        CTRL_BREAK_EVENT
        CTRL_CLOSE_EVENT
        CTRL_LOGOFF_EVENT = 5
        CTRL_SHUTDOWN_EVENT
    End Enum

    Public Event ControlMessage(ByVal type As ControlTypes)

    Dim msgDelegate As MessageDelegate

    Public Sub New()
        msgDelegate = New MessageDelegate(AddressOf Me.MessageProc)
        Dim ret As Integer
        ret = SetConsoleCtrlHandler(msgDelegate, True)

        If ret = 0 Then
            Console.WriteLine("Unable to set Console Control Handler: " + New Win32Exception(Marshal.GetLastWin32Error()).Message, EventLogEntryType.Error)
        End If
    End Sub

    Private Function MessageProc(ByVal type As ControlTypes) As Boolean
        RaiseEvent ControlMessage(type)
    End Function

End Class
