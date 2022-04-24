Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Text
Imports System.Security.Principal

Public Class KernelApi

#Region "Constants and Enums"

    Public Const MEM_COMMIT As UInteger = &H1000
    Public Const MEM_RELEASE As UInteger = &H8000
    Public Const PAGE_READWRITE As UInteger = &H4

    Const ERROR_NONE_MAPPED As Integer = 1332

    Public Enum LogonProvider As Integer
        'Use the standard logon provider for the system. 
        'The default security provider is negotiate, unless you pass NULL for the domain name and the user name 
        'is not in UPN format. In this case, the default provider is NTLM. 
        'NOTE: Windows 2000/NT:   The default security provider is NTLM.
        LOGON32_PROVIDER_DEFAULT = 0
    End Enum

    Public Enum LogonType As Integer
        'This logon type is intended for users who will be interactively using the computer, such as a user being logged on 
        'by a terminal server, remote shell, or similar process.
        'This logon type has the additional expense of caching logon information for disconnected operations; 
        'therefore, it is inappropriate for some client/server applications,
        'such as a mail server.
        LOGON32_LOGON_INTERACTIVE = 2

        'This logon type is intended for high performance servers to authenticate plaintext passwords.
        'The LogonUser function does not cache credentials for this logon type.
        LOGON32_LOGON_NETWORK = 3

        'This logon type is intended for batch servers, where processes may be executing on behalf of a user without 
        'their direct intervention. This type is also for higher performance servers that process many plaintext
        'authentication attempts at a time, such as mail or Web servers. 
        'The LogonUser function does not cache credentials for this logon type.
        LOGON32_LOGON_BATCH = 4

        'Indicates a service-type logon. The account provided must have the service privilege enabled. 
        LOGON32_LOGON_SERVICE = 5

        'This logon type is for GINA DLLs that log on users who will be interactively using the computer. 
        'This logon type can generate a unique audit record that shows when the workstation was unlocked. 
        LOGON32_LOGON_UNLOCK = 7

        'This logon type preserves the name and password in the authentication package, which allows the server to make 
        'connections to other network servers while impersonating the client. A server can accept plaintext credentials 
        'from a client, call LogonUser, verify that the user can access the system across the network, and still 
        'communicate with other servers.
        'NOTE: Windows NT:  This value is not supported. 
        LOGON32_LOGON_NETWORK_CLEARTEXT = 8

        'This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
        'The new logon session has the same local identifier but uses different credentials for other network connections. 
        'NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
        'NOTE: Windows NT:  This value is not supported. 
        LOGON32_LOGON_NEW_CREDENTIALS = 9
    End Enum

    Public Enum ImpersonationLevel
        SECURITY_ANONYMOUS = 0
        SECURITY_IDENTIFICATION = 1
        SECURITY_IMPERSONATION = 2
        SECURITY_DELEGATION = 3
    End Enum

    Public Enum TokenType
        TOKEN_PRIMARY = 1
        TOKEN_IMPERSONATION = 2
    End Enum

    Public Const STD_INPUT_HANDLE As Integer = -10
    Public Const STD_ERROR_HANDLE As Integer = -12

    Public Enum StandardRights As UInteger
        STANDARD_RIGHTS_REQUIRED = &HF0000
        STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
        STANDARD_RIGHTS_READ = (READ_CONTROL)
        STANDARD_RIGHTS_WRITE = (READ_CONTROL)
        STANDARD_RIGHTS_ALL = &H1F0000
    End Enum

    Public Enum TokenAccess As UInteger
        'ACCESS_SYSTEM_SECURITY = &H1000000
        TOKEN_ASSIGN_PRIMARY = &H1
        TOKEN_DUPLICATE = &H2
        TOKEN_IMPERSONATE = &H4
        TOKEN_QUERY = &H8
        TOKEN_QUERY_SOURCE = &H10
        TOKEN_ADJUST_PRIVILEGES = &H20
        TOKEN_ADJUST_GROUPS = &H40
        TOKEN_ADJUST_DEFAULT = &H80
        TOKEN_ADJUST_SESSIONID = &H100
        TOKEN_READ = (StandardAccess.STANDARD_RIGHTS_READ Or TOKEN_QUERY)
        TOKEN_WRITE = (StandardAccess.STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
        TOKEN_EXECUTE = (StandardAccess.STANDARD_RIGHTS_EXECUTE Or TOKEN_IMPERSONATE)
        TOKEN_ALL_ACCESS = _
            (TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID _
            Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_DUPLICATE _
            Or TOKEN_ASSIGN_PRIMARY Or TOKEN_ADJUST_SESSIONID Or TOKEN_IMPERSONATE _
            Or StandardAccess.STANDARD_RIGHTS_REQUIRED Or TOKEN_READ Or TOKEN_WRITE Or TOKEN_EXECUTE)
    End Enum

    Public Const GENERIC_ALL_ACCESS As Integer = &H10000000

    Public Const CREATE_NEW_CONSOLE As Integer = &H10
    Public Const NORMAL_PRIORITY_CLASS As Integer = &H20
    Public Const CREATE_UNICODE_ENVIRONMENT As Integer = &H40

    Public Enum STARTF As Integer
        STARTF_USESHOWWINDOW = &H1
        STARTF_USESIZE = &H2
        STARTF_USEPOSITION = &H4
        STARTF_USECOUNTCHARS = &H8
        STARTF_USEFILLATTRIBUTE = &H10
        STARTF_RUNFULLSCREEN = &H20
        STARTF_FORCEONFEEDBACK = &H40
        STARTF_FORCEOFFFEEDBACK = &H80
        STARTF_USESTDHANDLES = &H100
        STARTF_USEHOTKEY = &H200
    End Enum

    Public Enum SW As Short
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_SHOWDEFAULT = 10
    End Enum

    ' Microsoft.Win23.NativeMethods 
    Public Shared INVALID_HANDLE_VALUE As IntPtr = CType(-1, IntPtr)
    Public Shared NullHandleRef As HandleRef = New HandleRef(Nothing, IntPtr.Zero)

    Public Enum StandardAccess
        DELETE = &H10000       'Required to delete the object. 
        READ_CONTROL = &H20000 'Required to read information in the security descriptor for the object, not including the information in the SACL. To read or write the SACL, you must request the ACCESS_SYSTEM_SECURITY access right. For more information, see SACL Access Right. 
        SYNCHRONIZE = &H100000 'Not supported for desktop objects. 
        WRITE_DAC = &H40000    'Required to modify the DACL in the security descriptor for the object. 
        WRITE_OWNER = &H80000
        STANDARD_RIGHTS_ALL = (DELETE Or READ_CONTROL Or WRITE_DAC Or WRITE_OWNER Or SYNCHRONIZE)
        STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
        STANDARD_RIGHTS_READ = (READ_CONTROL)
        STANDARD_RIGHTS_REQUIRED = (DELETE Or READ_CONTROL Or WRITE_DAC Or WRITE_OWNER)
        STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    End Enum

    Public Const MAXIMUM_ALLOWED As Integer = &H2000000

    Public Const TOKEN_INTEGRITY_LOW_SID As String = "S-1-16-4096"
    Public Const TOKEN_INTEGRITY_MEDIUM_SID As String = "S-1-16-8192"
    Public Const TOKEN_INTEGRITY_HIGH_SID As String = "S-1-16-12288"
    Public Const TOKEN_INTEGRITY_SYSTEM_SID As String = "S-1-16-16384"

    Public Const READ_CONTROL As Integer = &H20000
    Public Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
    Public Const STANDARD_RIGHTS_READ As Integer = READ_CONTROL
    Public Const STANDARD_RIGHTS_WRITE As Integer = READ_CONTROL
    Public Const STANDARD_RIGHTS_EXECUTE As Integer = READ_CONTROL

    Public Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
    Public Const SPECIFIC_RIGHTS_ALL As Integer = &HFFFF

    'Public Const TOKEN_ASSIGN_PRIMARY As Integer = &H1
    'Public Const TOKEN_DUPLICATE As Integer = &H2
    'Public Const TOKEN_IMPERSONATE As Integer = &H4
    'Public Const TOKEN_QUERY As Integer = &H8
    'Public Const TOKEN_QUERY_SOURCE As Integer = &H10
    'Public Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
    'Public Const TOKEN_ADJUST_GROUPS As Integer = &H40
    'Public Const TOKEN_ADJUST_DEFAULT As Integer = &H80
    'Public Const TOKEN_ADJUST_SESSIONID As Integer = &H100
    'Public Const TOKEN_ALL_ACCESS_P As Integer = (STANDARD_RIGHTS_REQUIRED Or _
    '                              TOKEN_ASSIGN_PRIMARY Or _
    '                              TOKEN_DUPLICATE Or _
    '                              TOKEN_IMPERSONATE Or _
    '                              TOKEN_QUERY Or _
    '                              TOKEN_QUERY_SOURCE Or _
    '                              TOKEN_ADJUST_PRIVILEGES Or _
    '                              TOKEN_ADJUST_GROUPS Or _
    '                              TOKEN_ADJUST_DEFAULT)
    'Public Const TOKEN_ALL_ACCESS As Integer = TOKEN_ALL_ACCESS_P Or _
    '                  TOKEN_ADJUST_SESSIONID
    'Public Const TOKEN_READ As Integer = STANDARD_RIGHTS_READ Or TOKEN_QUERY
    'Public Const TOKEN_WRITE As Integer = STANDARD_RIGHTS_WRITE Or _
    '                              TOKEN_ADJUST_PRIVILEGES Or _
    '                              TOKEN_ADJUST_GROUPS Or _
    '                              TOKEN_ADJUST_DEFAULT
    'Public TOKEN_EXECUTE As Integer = STANDARD_RIGHTS_EXECUTE

    Private Const SE_PRIVILEGE_ENABLED As Integer = &H2
    Public Const SE_CHANGE_NOTIFY_NAME As String = "SeChangeNotifyPrivilege"
    Public Const SE_SHUTDOWN_PRIVILEGE As String = "SeShutdownPrivilege"
    Public Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
    Public Const SE_BACKUP_NAME As String = "SeBackupPrivilege"
    Public Const SE_TCB_PRIVILEGE As String = "SeTcbPrivilege"
    Public Const SE_DEBUG_PRIVILEGE As String = "SeDebugPrivilege"
    Public Const SE_TAKE_OWNERSHIP_NAME As String = "SeTakeOwnershipPrivilege"
    Public Const SE_SECURITY_PRIVILEGE As String = "SeSecurityPrivilege"
    Public Const SE_ASSIGNPRIMARYTOKEN_NAME As String = "SeAssignPrimaryTokenPrivilege"
    Public Const SE_AUDIT_NAME As String = "SeAuditPrivilege"
    Public Const SE_CREATE_PAGEFILE_NAME As String = "SeCreatePagefilePrivilege"
    Public Const SE_CREATE_PERMANENT_NAME As String = "SeCreatePermanentPrivilege"
    Public Const SE_CREATE_TOKEN_NAME As String = "SeCreateTokenPrivilege"
    Public Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Public Const SE_INC_BASE_PRIORITY_NAME As String = "SeIncreaseBasePriorityPrivilege"
    Public Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
    Public Const SE_LOAD_DRIVER_NAME As String = "SeLoadDriverPrivilege"
    Public Const SE_LOCK_MEMORY_NAME As String = "SeLockMemoryPrivilege"
    Public Const SE_MACHINE_ACCOUNT_NAME As String = "SeMachineAccountPrivilege"
    Public Const SE_PROF_SINGLE_PROCESS_NAME As String = "SeProfileSingleProcessPrivilege"
    Public Const SE_REMOTE_SHUTDOWN_NAME As String = "SeRemoteShutdownPrivilege"
    Public Const SE_SECURITY_NAME As String = "SeSecurityPrivilege"
    Public Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
    Public Const SE_SYSTEM_ENVIRONMENT_NAME As String = "SeSystemEnvironmentPrivilege"
    Public Const SE_SYSTEM_PROFILE_NAME As String = "SeSystemProfilePrivilege"
    Public Const SE_SYSTEMTIME_NAME As String = "SeSystemtimePrivilege"


    Public Enum SECURITY_MANDATORY_RID
        SECURITY_MANDATORY_UNTRUSTED_RID = &H0 'Untrusted. 
        SECURITY_MANDATORY_LOW_RID = &H1000 'Low integrity. 
        SECURITY_MANDATORY_MEDIUM_RID = &H2000 'Medium integrity. 
        SECURITY_MANDATORY_HIGH_RID = &H3000 'High integrity. 
        SECURITY_MANDATORY_SYSTEM_RID = &H4000 'System integrity. 
        SECURITY_MANDATORY_PROTECTED_PROCESS_RID = &H5000 'Protected process. 
    End Enum

    Public Enum SE_GROUP
        SE_GROUP_ENABLED = &H4
        SE_GROUP_ENABLED_BY_DEFAULT = &H2
        SE_GROUP_INTEGRITY = &H20
        SE_GROUP_INTEGRITY_ENABLED = &H40
        SE_GROUP_LOGON_ID = &HC0000000
        SE_GROUP_MANDATORY = &H1
        SE_GROUP_OWNER = &H8
        SE_GROUP_RESOURCE = &H20000000
        SE_GROUP_USE_FOR_DENY_ONLY = &H10
    End Enum

    Private Const DF_ALLOWOTHERACCOUNTHOOK As Integer = &H1

    Public Enum DesktopAccess
        DESKTOP_CREATEMENU = &H4
        DESKTOP_CREATEWINDOW = &H2
        DESKTOP_ENUMERATE = &H40
        DESKTOP_HOOKCONTROL = &H8
        DESKTOP_WRITEOBJECTS = &H80
        DESKTOP_READOBJECTS = &H1
        DESKTOP_SWITCHDESKTOP = &H100
        DESKTOP_JOURNALPLAYBACK = &H20
        DESKTOP_JOURNALRECORD = &H10

    End Enum

    Public Enum GenericAccess
        GENERIC_READ = (DesktopAccess.DESKTOP_ENUMERATE Or DesktopAccess.DESKTOP_READOBJECTS Or StandardAccess.STANDARD_RIGHTS_READ)
        GENERIC_WRITE = (DesktopAccess.DESKTOP_CREATEMENU Or DesktopAccess.DESKTOP_CREATEWINDOW Or DesktopAccess.DESKTOP_HOOKCONTROL Or DesktopAccess.DESKTOP_JOURNALPLAYBACK Or DesktopAccess.DESKTOP_JOURNALRECORD Or DesktopAccess.DESKTOP_WRITEOBJECTS Or StandardAccess.STANDARD_RIGHTS_WRITE)
        GENERIC_EXECUTE = (DesktopAccess.DESKTOP_SWITCHDESKTOP Or StandardAccess.STANDARD_RIGHTS_EXECUTE)
        GENERIC_ALL = (DesktopAccess.DESKTOP_CREATEMENU Or DesktopAccess.DESKTOP_CREATEWINDOW Or DesktopAccess.DESKTOP_ENUMERATE Or DesktopAccess.DESKTOP_HOOKCONTROL Or DesktopAccess.DESKTOP_JOURNALPLAYBACK Or DesktopAccess.DESKTOP_JOURNALRECORD Or DesktopAccess.DESKTOP_READOBJECTS Or DesktopAccess.DESKTOP_SWITCHDESKTOP Or DesktopAccess.DESKTOP_WRITEOBJECTS Or StandardAccess.STANDARD_RIGHTS_REQUIRED)
    End Enum

    Public Enum SID_NAME_USE
        SidTypeUser = 1
        SidTypeGroup
        SidTypeDomain
        SidTypeAlias
        SidTypeWellKnownGroup
        SidTypeDeletedAccount
        SidTypeInvalid
        SidTypeUnknown
        SidTypeComputer
        SidTypeLabel
    End Enum

    Public Enum WinStationAccess
        WINSTA_ACCESSCLIPBOARD = &H4
        WINSTA_ACCESSGLOBALATOMS = &H20
        WINSTA_ACCESSPUBLICATOMS = &H20
        WINSTA_CREATEDESKTOP = &H8
        WINSTA_ENUMDESKTOPS = &H1
        WINSTA_ENUMERATE = &H100
        WINSTA_EXITWINDOWS = &H40
        WINSTA_READATTRIBUTES = &H2
        WINSTA_READSCREEN = &H200
        WINSTA_WRITEATTRIBUTES = &H10
    End Enum

    Private Const UOI_FLAGS As Integer = 1
    Private Const UOI_NAME As Integer = 2
    Private Const UOI_TYPE As Integer = 3
    Private Const UOI_USER_SID As Integer = 4

    Public Enum SecurityObjectType As Integer
        SE_UNKNOWN_OBJECT_TYPE = 0
        SE_FILE_OBJECT
        SE_SERVICE
        SE_PRINTER
        SE_REGISTRY_KEY
        SE_LMSHARE
        SE_KERNEL_OBJECT
        SE_WINDOW_OBJECT
        SE_DS_OBJECT
        SE_DS_OBJECT_ALL
        SE_PROVIDER_DEFINED_OBJECT
        SE_WMIGUID_OBJECT
        SE_REGISTRY_WOW64_32
    End Enum

    Private Enum SecurityInfo As Integer
        OWNER_SECURITY_INFORMATION = 1
        GROUP_SECURITY_INFORMATION = 2
        DACL_SECURITY_INFORMATION = 4
        SACL_SECURITY_INFORMATION = 8
        PROTECTED_SACL_SECURITY_INFORMATION = 16
        PROTECTED_DACL_SECURITY_INFORMATION = 32
        UNPROTECTED_SACL_SECURITY_INFORMATION = 64
        UNPROTECTED_DACL_SECURITY_INFORMATION = 128
    End Enum

    Public Enum TOKEN_INFORMATION_CLASS
        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        MaxTokenInfoClass  '// MaxTokenInfoClass should always be the last enum
    End Enum

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure TOKEN_PRIMARY_GROUP
        Public PrimaryGroup As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure SID_AND_ATTRIBUTES
        Public Sid As IntPtr
        Public Attributes As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure TOKEN_MANDATORY_LABEL
        Public Label As SID_AND_ATTRIBUTES
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Private Structure Luid
        Dim Count As Integer
        Dim Luid As Long
        Dim Attr As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure SECURITY_ATTRIBUTES
        Dim nLength As Int32
        Dim lpSecurityDescriptor As Int32
        Dim bInheritHandle As Boolean
    End Structure
#End Region

#Region "API Declarations"

    Declare Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As Integer) As Integer

    <DllImport("kernel32.dll")> _
    Public Shared Function WaitForSingleObject( _
        ByVal hHandle As IntPtr, _
        ByVal dwMilliseconds As Integer) As Integer
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function GetModuleHandle( _
        ByVal lpModuleName As String) As IntPtr
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function VirtualAllocEx( _
        ByVal hProcess As IntPtr, _
        ByVal lpAddress As IntPtr, _
        ByVal dwSize As Integer, _
        ByVal flAllocationType As UInteger, _
        ByVal flProtect As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function VirtualFreeEx( _
        ByVal hProcess As IntPtr, _
        ByVal lpAddress As IntPtr, _
        ByVal dwSize As Integer, _
        ByVal dwFreeType As UInteger) As Boolean
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function WriteProcessMemory( _
        ByVal hProcess As IntPtr, _
        ByVal lpBaseAddress As IntPtr, _
        ByRef buffer As String, _
        ByVal dwSize As Integer, _
        ByVal lpNumberOfBytesWritten As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function CreateRemoteThread( _
        ByVal hProcess As IntPtr, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwStackSize As Integer, _
        ByVal lpStartAddress As IntPtr, _
        ByVal lpParameter As IntPtr, _
        ByVal dwCreationFlags As Integer, _
        ByRef lpThreadId As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll")> _
    Public Shared Function GetProcAddress( _
        ByVal hModule As IntPtr, _
        ByVal lpProcName As String) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function GetComputerName( _
        ByVal lpBuffer As StringBuilder, _
        ByRef nSize As Integer) As Integer
    End Function

    Public Shared Function GetComputerName() As String
        Dim length As Integer = 128
        Dim name As New StringBuilder(length)

        Dim result As Integer
        result = GetComputerName(name, length)
        If result = 0 Then
            Return ""
        Else
            Return name.ToString().Substring(0, length)
        End If
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function DuplicateTokenEx( _
        ByVal hToken As IntPtr, _
        ByVal access As UInteger, _
        ByRef tokenAttributes As SECURITY_ATTRIBUTES, _
        ByVal impersonationLevel As Integer, _
        ByVal tokenType As Integer, _
        ByRef hNewToken As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CloseHandle( _
         ByVal hObject As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CloseHandle( _
        ByVal hObject As Integer) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenProcess( _
        ByVal dwDesiredAccess As PROCESS_ACCESS, _
        ByVal bInheritHandle As Boolean, _
        ByVal dwProcessId As Integer) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function OpenProcessToken( _
        ByVal h As IntPtr, _
        ByVal acc As TokenAccess, _
        ByRef phtok As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function GetTokenInformation( _
        ByVal TokenHandle As IntPtr, _
        ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, _
        ByVal TokenInformation As IntPtr, _
        ByVal TokenInformationLength As Integer, _
        ByRef ReturnLength As Integer) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function SetTokenInformation( _
        ByVal TokenHandle As IntPtr, _
        ByVal TokenInformationClass As TOKEN_INFORMATION_CLASS, _
        ByVal TokenInformation As IntPtr, _
        ByVal TokenInformationLength As Integer) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function IsTokenRestricted( _
        ByVal phtok As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Private Shared Function LookupPrivilegeValue( _
        ByVal host As String, _
        ByVal name As String, _
        ByRef pluid As Long) As Boolean
    End Function

    <DllImport("advapi32.dll", ExactSpelling:=True, SetLastError:=True)> _
    Private Shared Function AdjustTokenPrivileges( _
        ByVal TokenHandle As IntPtr, _
        ByVal DisableAllPrivileges As Boolean, _
        ByRef NewState As Luid, _
        ByVal BufferLength As Integer, _
        ByVal PreviousState As IntPtr, _
        ByVal ReturnLength As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenWindowStation( _
        ByVal lpszWinSta As String, _
        ByVal fInherit As Boolean, _
        ByVal dwDesiredAccess As KernelApi.WinStationAccess) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenWindowStation( _
        ByVal lpszWinSta As String, _
        ByVal fInherit As Boolean, _
        ByVal dwDesiredAccess As KernelApi.GenericAccess) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function CloseWindowStation( _
        ByVal station As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenInputDesktop( _
        ByVal dwFlags As Integer, _
        ByVal fInherit As Boolean, _
        ByVal dwDesiredAccess As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenInputDesktop( _
        ByVal dwFlags As Integer, _
        ByVal fInherit As Boolean, _
        ByVal dwDesiredAccess As KernelApi.GenericAccess) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetThreadDesktop( _
        ByVal dwThread As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function CloseDesktop( _
        ByVal handle As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetUserObjectInformation( _
        ByVal hObj As IntPtr, _
        ByVal nIndex As Integer, _
        <MarshalAs(UnmanagedType.LPArray)> ByVal pvInfo As Byte(), _
        ByVal nLength As Integer, _
        ByRef lpnLengthNeeded As Integer) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetNamedSecurityInfo( _
        ByVal pObjectName As String, _
        ByVal ObjectType As SecurityObjectType, _
        ByVal SecurityInfo As SecurityInfo, _
        ByRef ppsidOwner As IntPtr, _
        ByRef ppsidGroup As IntPtr, _
        ByRef ppDacl As IntPtr, _
        ByRef ppSacl As IntPtr, _
        ByRef ppSecurityDescriptor As IntPtr) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure SHELLEXECUTEINFO
        Public cbSize As Integer
        Public fMask As Integer
        Public hwnd As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)> Public lpVerb As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpFile As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpParameters As String
        <MarshalAs(UnmanagedType.LPTStr)> Public lpDirectory As String
        Public nShow As WindowsApi.ShowWindowEnum
        Public hInstApp As IntPtr
        Public lpIDList As IntPtr
        <MarshalAs(UnmanagedType.LPTStr)> Public lpClass As String
        Public hkeyClass As IntPtr
        Public dwHotKey As Integer
        Public hIcon As IntPtr
        Public hProcess As IntPtr
    End Structure

    Public Enum PROCESS_ACCESS As Integer
        ' Specifies all possible access flags for the process object.
        PROCESS_ALL_ACCESS = &H1F0FFF

        ' Enables using the process handle in the CreateRemoteThread function
        ' to create a thread in the process.
        PROCESS_CREATE_THREAD = &H2

        ' Enables using the process handle as either the source or
        ' target process in the DuplicateHandle function to duplicate a handle.
        PROCESS_DUP_HANDLE = &H40

        ' Enables using the process handle in the GetExitCodeProcess and
        ' GetPriorityClass functions to read information from the process object.
        PROCESS_QUERY_INFORMATION = &H400

        ' Enables using the process handle in the SetPriorityClass function to
        ' set the priority class of the process.
        PROCESS_SET_INFORMATION = &H200

        ' Enables using the process handle in the TerminateProcess function to
        ' terminate the process.
        PROCESS_TERMINATE = &H1

        ' Enables using the process handle in the VirtualProtectEx and
        ' WriteProcessMemory functions to modify the virtual memory of the process.
        PROCESS_VM_OPERATION = &H8

        ' Enables using the process handle in the ReadProcessMemory function to
        ' read from the virtual memory of the process.
        PROCESS_VM_READ = &H10

        ' Enables using the process handle in the WriteProcessMemory function to
        ' write to the virtual memory of the process.
        PROCESS_VM_WRITE = &H20

        ' Enables using the process handle in any of the wait functions to wait
        ' for the process to terminate.
        SYNCHRONIZE = &H100000

        MAXIMUM_ALLOWED = &H2000000
    End Enum

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Structure STARTUPINFO
        Dim cb As Int32
        Dim lpReserved As Int32
        Dim lpDesktop As String
        Dim lpTitle As Int32
        Dim dwX As Int32
        Dim dwY As Int32
        Dim dwXSize As Int32
        Dim dwYSize As Int32
        Dim dwXCountChars As Int32
        Dim dwYCountChars As Int32
        Dim dwFillAttribute As Int32
        Dim dwFlags As Int32
        Dim wShowWindow As Int16
        Dim cbReserved2 As Int16
        Dim lpReserved2 As Int32
        Dim hStdInput As Int32
        Dim hStdOutput As Int32
        Dim hStdError As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure PROCESS_INFORMATION
        Dim hProcess As Int32
        Dim hThread As Int32
        Dim dwProcessId As Int32
        Dim dwThreadId As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure LUID_STRUCTURE
        Public LowPart As Integer
        Public HighPart As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure LUID_AND_ATTRIBUTES
        Public Luid As LUID_STRUCTURE
        Public Attributes As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure PRIVILEGE_SET
        Public PrivilegeCount As Integer
        Public Control As Integer
        <MarshalAs(UnmanagedType.ByValArray)> Public Privilege As LUID_AND_ATTRIBUTES()
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function CloseHandle(ByVal handle As HandleRef) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function CreateProcess( _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal lpApplicationName As String, _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Boolean, _
        ByVal dwCreationFlags As Integer, _
        ByVal lpEnvironment As IntPtr, _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal lpCurrentDirectory As String, _
        ByVal lpStartupInfo As STARTUPINFO, _
        ByVal lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function CreateProcessAsUserW( _
        ByVal hToken As IntPtr, _
        ByVal lpApplicationName As String, _
        ByVal lpCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Boolean, _
        ByVal dwCreationFlags As Int32, _
        ByVal lpEnvironment As IntPtr, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetStdHandle( _
        ByVal whichHandle As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function CreateFile(ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Integer, _
        ByVal dwShareMode As Integer, _
        ByVal lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwCreationDisposition As Integer, _
        ByVal dwFlagsAndAttributes As Integer, _
        ByVal hTemplateFile As HandleRef) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function CreateNamedPipe( _
        ByVal name As String, _
        ByVal openMode As UInt32, _
        ByVal pipeMode As UInt32, _
        ByVal maxInstances As UInt32, _
        ByVal outBufSize As UInt32, _
        ByVal inBufSize As UInt32, _
        ByVal timeout As UInt32, _
        ByVal lpPipeAttributes As SECURITY_ATTRIBUTES) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function GetConsoleOutputCP() As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LookupAccountName( _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal systemName As String, _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal accountName As String, _
        ByVal sid As IntPtr, _
        ByRef cbSid As Integer, _
        ByVal referencedDomainName As StringBuilder, _
        ByRef cbReferencedDomainName As Integer, _
        ByRef use As Integer) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LookupAccountName( _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal systemName As String, _
        <MarshalAs(UnmanagedType.LPTStr)> ByVal accountName As String, _
        <MarshalAs(UnmanagedType.LPArray)> ByVal sid As Byte(), _
        ByRef cbSid As Integer, _
        ByVal referencedDomainName As StringBuilder, _
        ByRef cbReferencedDomainName As Integer, _
        ByRef use As Integer) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LookupAccountSid( _
        ByVal systemName As String, _
        ByVal psid As IntPtr, _
        ByVal accountName As StringBuilder, _
        ByRef cbAccount As Integer, _
        ByVal domainName As StringBuilder, _
        ByRef cbDomainName As Integer, _
        ByRef use As Integer) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LookupAccountSid( _
        ByVal systemName As String, _
        <MarshalAs(UnmanagedType.LPArray)> ByVal psid As Byte(), _
        ByVal accountName As StringBuilder, _
        ByRef cbAccount As Integer, _
        ByVal domainName As StringBuilder, _
        ByRef cbDomainName As Integer, _
        ByRef use As SID_NAME_USE) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function ConvertStringSidToSid( _
        ByVal Sid As String, _
        ByRef StringSid As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function ConvertSidToStringSid( _
        <MarshalAs(UnmanagedType.LPArray)> ByVal sid As Byte(), _
        ByRef stringSID As IntPtr) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function ConvertSidToStringSid( _
        ByVal sid As IntPtr, _
        ByRef stringSID As IntPtr) As Integer
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LocalFree( _
        ByVal hMem As IntPtr) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Public Shared Function GetLengthSid( _
        ByVal sid As IntPtr) As Integer
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Private Shared Function CheckTokenMembership( _
        ByVal TokenHandle As IntPtr, _
        ByVal SidToCheck As IntPtr, _
        ByRef IsMember As Boolean) As Boolean
    End Function

    Public Enum DriveType
        Floppy = 1
        Removable
        Fixed
        Remote
        CDROM
        RAMDisk
    End Enum

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetDriveType( _
        ByVal nDrive As String) As DriveType
    End Function

    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function LogonUser( _
        ByVal lpszUsername As String, _
        ByVal lpszDomain As String, _
        ByVal lpszPassword As String, _
        ByVal dwLogonType As LogonType, _
        ByVal dwLogonProvider As LogonProvider, _
        ByRef phToken As IntPtr) As Integer
    End Function

#End Region

#Region "Public Interface"

    Public Shared Function LogonUser(ByVal username As String, ByVal password As String, ByVal type As LogonType) As IntPtr
        Utilities.EnableBlankPasswords(True)

        Dim token As IntPtr = IntPtr.Zero
        Dim result As Integer
        result = KernelApi.LogonUser(username, _
                        ".", _
                        password, _
                        type, _
                        LogonProvider.LOGON32_PROVIDER_DEFAULT, _
                        token)
        If result = 0 Then
            Throw New AccessViolationException("Error logging on user: " + (New Win32Exception(Marshal.GetLastWin32Error())).Message)
        End If

        Return token
    End Function

    Public Shared Function CheckTokenMembership(ByVal token As IntPtr, ByVal SidToCheck As IntPtr) As Boolean
        Dim IsMember As Boolean
        If Not CheckTokenMembership(token, SidToCheck, IsMember) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return IsMember
    End Function

    Public Shared Function CheckTokenMembership(ByVal token As IntPtr, ByVal SidToCheck As Byte()) As Boolean

        Dim hObject As GCHandle
        Try
            hObject = GCHandle.Alloc(SidToCheck, GCHandleType.Pinned)
            Dim ptrSid As IntPtr = hObject.AddrOfPinnedObject()
            Return CheckTokenMembership(token, ptrSid)
        Finally
            If hObject.IsAllocated Then
                hObject.Free()
            End If
        End Try
    End Function

    Public Shared Function ConvertStringSidToSid(ByVal sid As String) As IntPtr
        Dim ptrSid As IntPtr = IntPtr.Zero

        Dim result As Boolean = KernelApi.ConvertStringSidToSid(sid, ptrSid)
        If Not result Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return ptrSid
    End Function

    Public Shared Function ConvertSidToStringSid(ByVal bytes As Byte()) As String
        Dim stringSid As IntPtr = IntPtr.Zero
        Try
            If ConvertSidToStringSid(bytes, stringSid) = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            Return Marshal.PtrToStringAuto(stringSid)
        Finally
            Marshal.FreeHGlobal(stringSid)
        End Try
    End Function

    Public Shared Sub AddPrivilege(ByVal name As String)
        Dim tp As Luid
        Dim hProcess As IntPtr = IntPtr.Zero
        Dim hToken As IntPtr = IntPtr.Zero

        Try
            hProcess = GetCurrentProcess()
            If hProcess = IntPtr.Zero Then
                Throw New InvalidOperationException("GetCurrentProcess failed")
            End If

            If Not OpenProcessToken(hProcess, TokenAccess.TOKEN_ADJUST_PRIVILEGES Or TokenAccess.TOKEN_QUERY, hToken) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            If hToken = IntPtr.Zero Then
                Throw New InvalidOperationException("OpenProcessToken failed")
            End If

            If Not LookupPrivilegeValue(Nothing, name, tp.Luid) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            If tp.Luid = 0 Then
                Throw New InvalidOperationException("LookupPrivilegeValue failed")
            End If

            tp.Count = 1
            tp.Attr = SE_PRIVILEGE_ENABLED
            If Not AdjustTokenPrivileges(hToken, False, tp, 0, Nothing, Nothing) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        Finally
            CloseHandle(hProcess)
            CloseHandle(hToken)
        End Try
    End Sub

    Public Shared Function GetCurrentStationSID() As String
        Dim station As IntPtr
        Try
            station = OpenWindowStation("winsta0", True, WinStationAccess.WINSTA_ENUMERATE)
            Return GetCurrentSID(station)
        Finally
            CloseWindowStation(station)
        End Try
    End Function

    Public Shared Function GetDesktopName(ByVal sessionId As Integer) As String
        'Default
        'Winlogin
        'Disconnect (used by Terminal Services)
        'Screen-saver

        Dim desk As IntPtr = IntPtr.Zero
        Try
            desk = OpenInputDesktop(0, True, DesktopAccess.DESKTOP_READOBJECTS)
            If desk = IntPtr.Zero Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            Dim name As String = GetCurrentName(desk)

            'remove the trailing character.  not sure what that is
            Return name.Replace(ChrW(65533), "")
        Finally
            If desk <> IntPtr.Zero Then CloseDesktop(desk)
        End Try
    End Function

    Public Shared Function GetCurrentDesktopSID() As String
        Dim desk As IntPtr
        Dim length As Integer = 0
        Dim byteBuffer(255) As Byte

        Try
            desk = OpenInputDesktop(0, True, GenericAccess.GENERIC_ALL)
            If desk = IntPtr.Zero Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            Return GetCurrentSID(desk)
        Finally
            CloseDesktop(desk)
        End Try
    End Function

    Public Shared Function GetOwnerAccount(ByVal objectType As SecurityObjectType, ByVal path As String) As String
        Dim result As Integer
        Dim ptrSecDesc As IntPtr
        Dim ptrSID As IntPtr

        Try
            result = GetNamedSecurityInfo(path, objectType, SecurityInfo.OWNER_SECURITY_INFORMATION, _
                ptrSID, Nothing, Nothing, Nothing, ptrSecDesc)
            If result <> 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            Return AccountFromSID(ptrSID)
        Finally
            Marshal.FreeHGlobal(ptrSecDesc)
        End Try
    End Function

    Public Shared Function GetSID(ByVal name As String) As String
        Dim _sid As IntPtr
        Dim _sidLength As Integer
        Dim _domainLength As Integer
        Dim _use As Integer
        Dim _domain As New StringBuilder()

        Try
            If Not LookupAccountName(Nothing, name, _sid, _sidLength, _domain, _domainLength, _use) Then
                Dim _error As Integer = Marshal.GetLastWin32Error()
                If _error = 122 Then
                    'ignore
                ElseIf _error = 1332 Then
                    Return "(unknown)"
                Else
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If
            End If
            _sid = Marshal.AllocHGlobal(_sidLength)

            If Not LookupAccountName(Nothing, name, _sid, _sidLength, _domain, _domainLength, _use) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            Return ConvertSidToStringSid(_sid)
        Finally
            Marshal.FreeHGlobal(_sid)
        End Try
    End Function

    Public Shared Function GetByteSID(ByVal name As String) As Byte()
        Dim _sid As Byte() = Nothing
        Dim _sidLength As Integer
        Dim _domainLength As Integer
        Dim _use As Integer
        Dim _domain As New StringBuilder()

        If Not LookupAccountName(Nothing, name, _sid, _sidLength, _domain, _domainLength, _use) Then
            If Marshal.GetLastWin32Error() <> 122 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End If
        ReDim _sid(_sidLength - 1)
        If Not LookupAccountName(Nothing, name, _sid, _sidLength, _domain, _domainLength, _use) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return _sid
    End Function

    Public Shared Function GetOwnerSID(ByVal objectType As SecurityObjectType, ByVal path As String) As String
        Dim result As Integer
        Dim ptrSecDesc As IntPtr
        Dim ptrSID As IntPtr

        Try
            result = GetNamedSecurityInfo(path, objectType, SecurityInfo.OWNER_SECURITY_INFORMATION, _
                ptrSID, Nothing, Nothing, Nothing, ptrSecDesc)
            If result <> 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
            Return ConvertSidToStringSid(ptrSID)
        Finally
            Marshal.FreeHGlobal(ptrSecDesc)
        End Try
    End Function

    Public Shared Function GetSystemSid() As SecurityIdentifier
        Return New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)
        'Dim account As NTAccount = New NTAccount("SYSTEM")
        'Return New SecurityIdentifier(account.Translate(GetType(SecurityIdentifier)).Value)
    End Function

    Public Shared Function AdministratorsGroupSid() As SecurityIdentifier
        'Return New SecurityIdentifier(KernelApi.ConvertSidToStringSid(Constants.AdministratorsGroupSid))
        Return New SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, Nothing)
    End Function

    Public Shared Function AdministratorsGroupName() As String
        Dim name As String = Nothing
        LookupAccountSid(Constants.AdministratorsGroupSid, Nothing, name, Nothing)
        Return name
    End Function

    Public Shared Function LookupAccountSid(ByVal sid As SecurityIdentifier, ByRef domain As String, ByRef username As String, ByRef sidtype As KernelApi.SID_NAME_USE) As Boolean
        Dim byteSid(sid.BinaryLength) As Byte
        sid.GetBinaryForm(byteSid, 0)
        Return LookupAccountSid(byteSid, domain, username, sidtype)
    End Function

    Public Shared Function LookupAccountSid(ByVal byteSid As Byte(), ByRef domain As String, ByRef username As String, ByRef sidType As KernelApi.SID_NAME_USE) As Boolean
        Dim name As New StringBuilder(255)
        Dim domainName As New StringBuilder(255)
        Dim nameLength As Integer = name.Capacity
        Dim domainNameLength As Integer = domainName.Capacity

        Dim result As Integer
        result = LookupAccountSid(Nothing, byteSid, name, nameLength, domainName, domainNameLength, sidType)
        If result = 0 Then
            result = Marshal.GetLastWin32Error()
            If result = KernelApi.ERROR_NONE_MAPPED Then
                'if the sid is not found return false, no error thrown
                Return False
            Else
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End If

        domain = domainName.ToString()
        username = name.ToString()
        Return True
    End Function

    Public Shared Function GetUserAccount(ByVal handle As IntPtr) As String
        Dim result As Integer
        Dim sidLength As Integer
        Dim userName As String = ""
        Dim domain As String = ""

        'get SID for the associated handle into a byte array
        Dim byteSID(255) As Byte
        result = GetUserObjectInformation(handle, UOI_USER_SID, byteSID, 256, sidLength)
        ReDim Preserve byteSID(sidLength - 1)

        KernelApi.LookupAccountSid(byteSID, domain, userName, Nothing)

        Return domain + "\" + userName
    End Function

    Public Shared Function GetUserSID(ByVal handle As IntPtr) As String
        Dim length As Integer
        Dim result As Integer

        Dim byteSID(255) As Byte
        result = GetUserObjectInformation(handle, UOI_USER_SID, byteSID, 256, length)
        ReDim Preserve byteSID(length - 1)

        Return KernelApi.ConvertSidToStringSid(byteSID)
    End Function

    Public Shared Function TokenIntegrityLevel(ByVal token As IntPtr) As String
        Dim info As IntPtr = IntPtr.Zero
        Dim infoLength As Integer = 0

        Try
            Dim result As Boolean
            result = KernelApi.GetTokenInformation(token, _
                        KernelApi.TOKEN_INFORMATION_CLASS.TokenIntegrityLevel, _
                        IntPtr.Zero, _
                        0, _
                        infoLength)
            info = Marshal.AllocHGlobal(infoLength)

            result = KernelApi.GetTokenInformation(token, _
                        KernelApi.TOKEN_INFORMATION_CLASS.TokenIntegrityLevel, _
                        info, _
                        infoLength, _
                        infoLength)

            If Not result Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            Dim sidAttrib As KernelApi.TOKEN_MANDATORY_LABEL
            sidAttrib = CType(Marshal.PtrToStructure(info, GetType(KernelApi.TOKEN_MANDATORY_LABEL)), KernelApi.TOKEN_MANDATORY_LABEL)

            Return KernelApi.ConvertSidToStringSid(sidAttrib.Label.Sid)
        Finally
            Marshal.FreeHGlobal(info)
        End Try
    End Function

    Public Shared Function CreatePrimaryToken(ByVal userToken As IntPtr) As IntPtr
        Dim primaryToken As IntPtr = IntPtr.Zero

        If Not DuplicateTokenEx(userToken, TokenAccess.TOKEN_ALL_ACCESS, Nothing, ImpersonationLevel.SECURITY_IDENTIFICATION, TokenType.TOKEN_PRIMARY, primaryToken) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return primaryToken
    End Function

#End Region

#Region "Private Functions"

    Private Shared Function AccountFromSID(ByRef sid As IntPtr) As String
        Dim result As Integer
        Dim _bufferSize As Integer = 64
        Dim _account As New StringBuilder(_bufferSize)
        Dim _domain As New StringBuilder(_bufferSize)
        Dim _accountLen As Integer = _bufferSize
        Dim _domainLen As Integer = _bufferSize
        Dim use As Integer

        result = LookupAccountSid(Nothing, sid, _account, _accountLen, _domain, _domainLen, use)
        If result <> 1 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Return _domain.ToString() + "\" + _account.ToString()
    End Function

    Public Shared Function ConvertSidToStringSid(ByVal sid As IntPtr) As String
        Dim result As Integer
        Dim ptrSID As IntPtr = IntPtr.Zero
        Dim stringSID As String = ""
        Try
            result = ConvertSidToStringSid(sid, ptrSID)
            If result <> 1 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            stringSID = Marshal.PtrToStringAuto(ptrSID)
        Finally
            KernelApi.LocalFree(ptrSID)
        End Try
        Return stringSID
    End Function

    Private Shared Function GetCurrentName(ByVal handle As IntPtr) As String
        Dim length As Integer = 0
        Dim byteBuffer(255) As Byte
        Dim result As Integer

        result = GetUserObjectInformation(handle, UOI_NAME, byteBuffer, 256, length)
        If result <> 1 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        ReDim Preserve byteBuffer(length - 2)

        Return Encoding.Unicode.GetString(byteBuffer).Replace(vbNullChar, "")
    End Function

    Private Shared Function GetCurrentSID(ByVal handle As IntPtr) As String
        Dim length As Integer = 0
        Dim byteBuffer(255) As Byte
        Dim result As Integer

        result = GetUserObjectInformation(handle, UOI_USER_SID, byteBuffer, 256, length)
        If result <> 1 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        ReDim Preserve byteBuffer(length - 1)

        Dim ptrSID As IntPtr = IntPtr.Zero

        result = ConvertSidToStringSid(byteBuffer, ptrSID)
        If result <> 1 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Return Marshal.PtrToStringAuto(ptrSID)
    End Function
#End Region

End Class
