Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Text
Imports System.Security.Principal
Imports System.ServiceProcess

Imports Microsoft.Win32

Public Class ServiceApi

    Public Enum StartupType As Integer
        Boot = 0
        System = 1
        Automatic = 2
        Manual = 3
        Disabled = 4
    End Enum

    Public Structure SERVICE_STATUS
        Dim dwServiceType As Int32
        Dim dwCurrentState As Int32
        Dim dwControlsAccepted As Int32
        Dim dwWin32ExitCode As Int32
        Dim dwServiceSpecificExitCode As Int32
        Dim dwCheckPoint As Int32
        Dim dwWaitHint As Int32
    End Structure

    Public Enum SERVICE_CONTROL As Integer
        [STOP] = &H1
        PAUSE = &H2
        [CONTINUE] = &H3
        INTERROGATE = &H4
        SHUTDOWN = &H5
        PARAMCHANGE = &H6
        NETBINDADD = &H7
        NETBINDREMOVE = &H8
        NETBINDENABLE = &H9
        NETBINDDISABLE = &HA
        DEVICEEVENT = &HB
        HARDWAREPROFILECHANGE = &HC
        POWEREVENT = &HD
        SESSIONCHANGE = &HE
    End Enum

    Public Enum SERVICE_STATE As Integer
        SERVICE_STOPPED = &H1
        SERVICE_START_PENDING = &H2
        SERVICE_STOP_PENDING = &H3
        SERVICE_RUNNING = &H4
        SERVICE_CONTINUE_PENDING = &H5
        SERVICE_PAUSE_PENDING = &H6
        SERVICE_PAUSED = &H7
    End Enum

    Public Enum SERVICE_ACCEPT As Integer
        [STOP] = &H1
        PAUSE_CONTINUE = &H2
        SHUTDOWN = &H4
        PARAMCHANGE = &H8
        NETBINDCHANGE = &H10
        HARDWAREPROFILECHANGE = &H20
        POWEREVENT = &H40
        SESSIONCHANGE = &H80
    End Enum

    Public Enum SERVICE_ACCESS
        STANDARD_RIGHTS_REQUIRED = &HF0000
        SERVICE_QUERY_CONFIG = &H1
        SERVICE_CHANGE_CONFIG = &H2
        SERVICE_QUERY_STATUS = &H4
        SERVICE_ENUMERATE_DEPENDENTS = &H8
        SERVICE_START = &H10
        SERVICE_STOP = &H20
        SERVICE_PAUSE_CONTINUE = &H40
        SERVICE_INTERROGATE = &H80
        SERVICE_USER_DEFINED_CONTROL = &H100
        SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
    End Enum

    Public Enum SCM_ACCESS
        STANDARD_RIGHTS_REQUIRED = &HF0000
        SC_MANAGER_CONNECT = &H1
        SC_MANAGER_CREATE_SERVICE = &H2
        SC_MANAGER_ENUMERATE_SERVICE = &H4
        SC_MANAGER_LOCK = &H8
        SC_MANAGER_QUERY_LOCK_STATUS = &H10
        SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
        SC_MANAGER_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG
    End Enum

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function QueryServiceStatus( _
        ByVal hService As IntPtr, _
        ByRef lpServiceStatus As SERVICE_STATUS) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function ControlService( _
        ByVal hService As IntPtr, _
        ByVal dwControlCode As SERVICE_CONTROL, _
        ByRef lpServiceStatus As SERVICE_STATUS) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function OpenSCManager( _
        ByVal machineName As String, _
        ByVal databaseName As String, _
        ByVal dwDesiredAccess As SCM_ACCESS) As IntPtr
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function OpenService( _
        ByVal hSCManager As IntPtr, _
        ByVal lpServiceName As String, _
        ByVal dwDesiredAccess As SERVICE_ACCESS) As IntPtr
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function DeleteService( _
        ByVal hService As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function CloseServiceHandle( _
        ByVal hSCObject As IntPtr) As Boolean
    End Function

    Public Shared Sub DeleteService(ByVal serviceName As String)
        Dim manager As IntPtr = IntPtr.Zero
        Dim service As IntPtr = IntPtr.Zero
        Dim status As New SERVICE_STATUS
        Try
            StopService(serviceName)

            manager = OpenSCManager(Nothing, Nothing, SCM_ACCESS.SC_MANAGER_ALL_ACCESS)
            If manager = IntPtr.Zero Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            service = OpenService(manager, serviceName, SERVICE_ACCESS.SERVICE_ALL_ACCESS)
            If service = IntPtr.Zero Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            If Not DeleteService(service) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

        Finally
            CloseServiceHandle(service)
            CloseServiceHandle(manager)
        End Try
    End Sub

    Public Shared Function GetService(ByVal serviceName As String) As ServiceController
        serviceName = serviceName.ToLower()
        For Each svc As ServiceController In ServiceController.GetServices()
            If svc.ServiceName.ToLower() = serviceName Then
                Return svc
            End If
            svc.Close()
        Next
        Return Nothing
    End Function

    Public Shared Function GetServiceByDisplayName(ByVal displayName As String) As ServiceController
        displayName = displayName.ToLower()
        For Each svc As ServiceController In ServiceController.GetServices()
            If svc.DisplayName.ToLower() = displayName Then
                Return svc
            End If
            svc.Close()
        Next
        Return Nothing
    End Function

    Public Shared Function ServiceExists(ByVal serviceName As String) As Boolean
        Dim svc As ServiceController = GetService(serviceName)
        If svc Is Nothing Then
            Return False
        Else
            svc.Close()
            Return True
        End If
    End Function

    Public Shared Sub StopService(ByVal serviceName As String)
        Dim svc As ServiceController = GetService(serviceName)
        If svc Is Nothing Then
            Throw New ArgumentOutOfRangeException("Service not found: " + serviceName)
        End If
        Try
            If svc.Status = ServiceControllerStatus.Running Then
                svc.Stop()
                svc.WaitForStatus(ServiceControllerStatus.Stopped)
            End If
        Finally
            svc.Close()
        End Try
    End Sub

    Public Shared Sub SetStartupType(ByVal svc As ServiceController, ByVal startup As StartupType)
        Dim keyName As String
        keyName = "SYSTEM\CurrentControlSet\Services\" + svc.ServiceName
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(keyName, True)
            key.SetValue("Start", CInt(startup))
        End Using
    End Sub
End Class

