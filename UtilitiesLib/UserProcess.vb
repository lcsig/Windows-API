Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Security.Principal
imports System.Text

Public Class UserProcess
    Inherits Process

    Private Const ERROR_ENVVAR_NOT_FOUND As Integer = 203

    <DllImport("userenv.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function CreateEnvironmentBlock( _
        ByRef lpEnvironment As IntPtr, _
        ByVal hToken As IntPtr, _
        ByVal bInherit As Boolean) As Integer
    End Function

    <DllImport("userenv.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function DestroyEnvironmentBlock( _
        ByVal lpEnvironment As IntPtr) As Boolean
    End Function

    Declare Auto Function CreatePipe Lib "kernel32.dll" ( _
            ByRef phReadPipe As Int32, _
            ByRef phWritePipe As Int32, _
            ByRef lpPipeAttributes As KernelApi.SECURITY_ATTRIBUTES, _
            ByVal nSize As Int32) _
    As Int32


    ' <summary> 
    ' Starts the process with the security token of the calling thread. 
    ' If the security token has a token type of TokenImpersonation, 
    ' the token will be duplicated to a primary token before calling 
    ' CreateProcessAsUser. 
    ' </summary> 
    ' <param name="process">The process to start.</param> 
    Public Sub StartAsUser(ByVal winlogon As Boolean)
        StartAsUser(WindowsIdentity.GetCurrent().Token, winlogon)
    End Sub

    Public Function StartAsUser(ByVal userToken As IntPtr, ByVal winlogon As Boolean, Optional ByVal hideWindow As Boolean = False) As Integer

        StartInfo.UseShellExecute = False

        Dim si As New KernelApi.STARTUPINFO()
        Dim pi As New KernelApi.PROCESS_INFORMATION()
        Dim primaryToken As IntPtr = IntPtr.Zero
        Dim commandLine As String = GetCommandLine().ToString()
        Dim workingDirectory As String = GetWorkingDirectory()
        Dim dwCreationFlag As Integer = KernelApi.CREATE_NEW_CONSOLE
        'si.wShowWindow = KernelApi.SW.SW_SHOWNORMAL
        Try
            primaryToken = KernelApi.CreatePrimaryToken(userToken)

            If hideWindow Then
                si.dwFlags = KernelApi.STARTF.STARTF_USESHOWWINDOW Or KernelApi.STARTF.STARTF_USESTDHANDLES
                si.wShowWindow = KernelApi.SW.SW_HIDE
            Else
                si.dwFlags = KernelApi.STARTF.STARTF_USESTDHANDLES
            End If
            'si.cb = Marshal.SizeOf(GetType(KernelApi.STARTUPINFO))
            si.cb = Len(si)

            If winlogon Then
                si.lpDesktop = "winsta0\winlogon"
            Else
                si.lpDesktop = "winsta0\default"
            End If

            'Logger.WriteEntry("Starting process: " + commandLine, EventLogEntryType.Warning)
            If Not KernelApi.CreateProcessAsUserW( _
                                primaryToken, _
                                Nothing, _
                                commandLine, _
                                Nothing, _
                                Nothing, _
                                False, _
                                dwCreationFlag, _
                                Nothing, _
                                workingDirectory, _
                                si, _
                                pi) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            Return Convert.ToInt32(pi.dwProcessId >> 32)
        Finally
            If pi.hThread <> -1 Then
                KernelApi.CloseHandle(pi.hThread)
            End If
            KernelApi.CloseHandle(primaryToken)
            'If Environment <> IntPtr.Zero Then
            '    UserProcess.DestroyEnvironmentBlock(Environment)
            'End If
            'KernelApi.CloseHandle(hReadPipe)
            'KernelApi.CloseHandle(hWritePipe)
        End Try
    End Function

    'Public Function StartAsUser(ByVal userToken As IntPtr, ByVal winlogon As Boolean, Optional ByVal hideWindow As Boolean = False) As Integer

    '    StartInfo.UseShellExecute = False

    'Dim result As Integer
    'Dim si As New KernelApi.STARTUPINFO()
    'Dim pi As New KernelApi.PROCESS_INFORMATION()
    'Dim primaryToken As IntPtr = IntPtr.Zero
    'Dim environment As IntPtr = IntPtr.Zero
    'Dim commandLine As String = GetCommandLine().ToString()
    'Dim workingDirectory As String = GetWorkingDirectory()
    'Dim dwCreationFlag As Integer = KernelApi.CREATE_NEW_CONSOLE

    ''Dim hReadPipe, hWritePipe As Int32

    '    Try
    '        primaryToken = KernelApi.CreatePrimaryToken(userToken)

    '        If hideWindow Then
    '            si.dwFlags = KernelApi.STARTF.STARTF_USESHOWWINDOW Or KernelApi.STARTF.STARTF_USESTDHANDLES
    '            si.wShowWindow = KernelApi.SW.SW_HIDE
    '        Else
    '            si.dwFlags = KernelApi.STARTF.STARTF_USESTDHANDLES
    '        End If
    '        si.cb = Marshal.SizeOf(GetType(KernelApi.STARTUPINFO))
    '        si.cb = Len(si)

    '        If winlogon Then
    '            si.lpDesktop = "winsta0\winlogon"
    '        Else
    '            si.lpDesktop = "winsta0\default"
    '        End If

    ''Logger.WriteEntry("Starting process: " + commandLine, EventLogEntryType.Warning)
    '        result = KernelApi.CreateProcessAsUserW( _
    '                            primaryToken, _
    '                            Nothing, _
    '                            commandLine, _
    '                            Nothing, _
    '                            Nothing, _
    '                            False, _
    '                            dwCreationFlag, _
    '                            0, _
    '                            workingDirectory, _
    '                            si, _
    '                            pi)
    '        If result = 0 Then
    '            Throw New Win32Exception(Marshal.GetLastWin32Error())
    '        End If
    '        Return pi.dwProcessId
    '    Finally
    '        If pi.hThread <> -1 Then
    '            KernelApi.CloseHandle(pi.hThread)
    '        End If
    '        KernelApi.CloseHandle(primaryToken)
    '        If environment <> IntPtr.Zero Then
    '            UserProcess.DestroyEnvironmentBlock(environment)
    '        End If
    ''KernelApi.CloseHandle(hReadPipe)
    ''KernelApi.CloseHandle(hWritePipe)
    '    End Try

    ''CloseHandle(New HandleRef(Me, stdoutWriteHandle))
    ''Dim stdinHandle As IntPtr = IntPtr.Zero
    ''Dim stdoutReadHandle As IntPtr = IntPtr.Zero
    ''Dim stdoutWriteHandle As IntPtr = IntPtr.Zero
    ''Dim stderrHandle As IntPtr = IntPtr.Zero
    ''stdinHandle = GetStdHandle(STD_INPUT_HANDLE)
    ''stderrHandle = GetStdHandle(STD_ERROR_HANDLE)
    ''MyCreatePipe(stdoutReadHandle, stdoutWriteHandle, False)
    ''assign handles to startup info 
    ''si.dwFlags = STARTF_USESTDHANDLES
    ''si.hStdInput = stdinHandle
    ''si.hStdOutput = stdoutWriteHandle
    ''si.hStdError = stderrHandle
    ' '' get reader for standard output from the child 
    ''Dim encode As Encoding = Encoding.GetEncoding(GetConsoleOutputCP())
    ''Dim safeHandle As New Microsoft.Win32.SafeHandles.SafeFileHandle(stdoutReadHandle, True)
    ''Dim standardOutput As New StreamReader(New FileStream(SafeHandle, FileAccess.Read, &H1000, True), encode)

    ' '' set this on the object accordingly.
    ''GetType(Process).InvokeMember("standardOutput", BindingFlags.SetField Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, Me, New Object() {standardOutput})

    ' '' scream if a process wasn't started instead of returning false. 
    ''If processInformation.hProcess = IntPtr.Zero Then
    ''    Throw New Exception("failed to create process")
    ''End If

    ' '' configure the properties of the Process object correctly
    ''GetType(Process).InvokeMember("SetProcessHandle", BindingFlags.InvokeMethod Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, Me, New Object() {processInformation.hProcess})
    ''GetType(Process).InvokeMember("SetProcessId", BindingFlags.InvokeMethod Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, Me, New Object() {processInformation.dwProcessId})
    'End Function

    ' <summary> 
    ' Gets the appropriate commandLine from the process. 
    ' </summary> 
    ' <param name="process"></param> 
    ' <returns></returns> 
    Private Function GetCommandLine() As StringBuilder
        Dim cmdLine As New StringBuilder()
        Dim fileName As String = StartInfo.FileName.Trim()
        Dim args As String = StartInfo.Arguments

        Dim fullQuotes As Boolean = fileName.StartsWith("""") AndAlso fileName.EndsWith("""")
        If Not fullQuotes Then
            cmdLine.Append("""")
        End If
        cmdLine.Append(fileName)
        If Not fullQuotes Then
            cmdLine.Append("""")
        End If
        If args <> Nothing AndAlso args.Length > 0 Then
            cmdLine.Append(" ")
            cmdLine.Append(args)
        End If
        Return cmdLine
    End Function

    ' <summary> 
    ' Gets the working directory or returns null if an empty string was found. 
    ' </summary> 
    ' <returns></returns> 
    Private Function GetWorkingDirectory() As String
        If StartInfo.WorkingDirectory <> "" Then
            Return StartInfo.WorkingDirectory
        Else
            Return Utilities.GetActualSystemPath("pmhk.exe", True)
        End If
    End Function

    ' <summary> 
    ' A clone of Process.CreatePipe. This is only implemented because reflection with 
    ' out parameters are a pain. 
    ' Note: This is only finished for w2k and higher machines. 
    ' </summary> 
    ' <param name="parentHandle"></param> 
    ' <param name="childHandle"></param> 
    ' <param name="parentInputs">Specifies whether the parent will be performing the writes.</param> 
    Public Shared Sub MyCreatePipe(ByRef parentHandle As IntPtr, ByRef childHandle As IntPtr, ByVal parentInputs As Boolean)
        Dim pipename As String = "\\.\pipe\Global\" + Guid.NewGuid().ToString()

        Dim attributes As New KernelApi.SECURITY_ATTRIBUTES()
        attributes.bInheritHandle = False
        attributes.nLength = Marshal.SizeOf(GetType(KernelApi.SECURITY_ATTRIBUTES))

        parentHandle = KernelApi.CreateNamedPipe(pipename, CUInt(&H40000003), 0, CUInt(&HFF), CUInt(&H1000), CUInt(&H1000), 0, attributes)
        If parentHandle <> KernelApi.INVALID_HANDLE_VALUE Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        attributes = New KernelApi.SECURITY_ATTRIBUTES()
        attributes.nLength = Marshal.SizeOf(GetType(KernelApi.SECURITY_ATTRIBUTES))

        attributes.bInheritHandle = True
        Dim num1 As Integer = &H40000000
        If parentInputs Then
            num1 = -2147483648
        End If

        childHandle = KernelApi.CreateFile(pipename, num1, 3, attributes, 3, &H40000080, KernelApi.NullHandleRef)
        If childHandle <> KernelApi.INVALID_HANDLE_VALUE Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

End Class
