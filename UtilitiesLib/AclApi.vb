Imports System.Security.AccessControl
Imports System.IO
Imports System.Security.Principal
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class AclApi

    Public Enum Hives As Integer
        HKEY_CLASSES_ROOT = &H80000000
        HKEY_CURRENT_USER = &H80000001
        HKEY_LOCAL_MACHINE = &H80000002
        HKEY_USERS = &H80000003
    End Enum

    Public Const ERROR_SUCCESS As Integer = 0

    Public Const READ_CONTROL As Integer = &H20000
    Public Const WRITE_DAC As Integer = &H40000
    Public Const WRITE_OWNER As Integer = &H80000
    Public Const KEY_WRITE As Integer = &H20006

    Public Enum SE_OBJECT_TYPE As Integer
        SE_UNKNOWN_OBJECT_TYPE = 0
        SE_FILE_OBJECT
        SE_SERVICE
        SE_PRINTER
        SE_REGISTRY_KEY = 4
        SE_LMSHARE
        SE_KERNEL_OBJECT
        SE_WINDOW_OBJECT
        SE_DS_OBJECT
        SE_DS_OBJECT_ALL
        SE_PROVIDER_DEFINED_OBJECT
        SE_WMIGUID_OBJECT
        SE_REGISTRY_WOW64_32KEY
    End Enum

    Public Enum SECURITY_INFORMATION
        OWNER_SECURITY_INFORMATION = &H1
        GROUP_SECURITY_INFORMATION = &H2
        DACL_SECURITY_INFORMATION = &H4
        SACL_SECURITY_INFORMATION = &H8
    End Enum

    Private Structure ACL
        Dim AclRevision As Byte
        Dim Sbz1 As Byte
        Dim AclSize As Integer
        Dim AceCount As Integer
        Dim Sbz2 As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SECURITY_DESCRIPTOR
        Dim Revision As Byte
        Dim Sbz1 As Byte
        Dim Control As Integer
        Dim Owner As IntPtr
        Dim Group As IntPtr
        Dim Sacl As Integer
        Dim Dacl As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SECURITY_ATTRIBUTES
        Dim nLength As Integer
        Dim pSecurityDescriptor As IntPtr
        Dim bInheritHandle As Boolean
    End Structure

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function InitializeSecurityDescriptor( _
        ByVal pSecurityDescriptor As IntPtr, _
        ByVal dwRevision As Integer) As Boolean
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function RegOpenKeyEx( _
        ByVal hKey As Integer, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Integer, _
        ByVal samDesired As Integer, _
        ByRef phkResult As IntPtr) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function RegCloseKey( _
        ByVal hKey As IntPtr) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function RegLoadKey( _
        ByVal hKey As Integer, _
        ByVal lpSubKey As String, _
        ByVal lpFile As String) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function RegUnLoadKey( _
        ByVal hKey As Integer, _
        ByVal lpSubKey As String) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function RegSetKeySecurity( _
        ByVal hKey As Integer, _
        ByVal SecurityInformation As AclApi.SECURITY_INFORMATION, _
        ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function SetSecurityDescriptorOwner( _
        ByVal pSecurityDescriptor As IntPtr, _
        ByVal pOwner As IntPtr, _
        ByVal bOwnerDefaulted As Integer) As Boolean
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function SetFileSecurity( _
        ByVal lpFileName As String, _
        ByVal SecurityInformation As SECURITY_INFORMATION, _
        ByVal pSecurityDescriptor As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function SetSecurityInfo( _
        ByVal handle As IntPtr, _
        ByVal ObjectType As AclApi.SE_OBJECT_TYPE, _
        ByVal SecurityInformation As AclApi.SECURITY_INFORMATION, _
        ByVal psidOwner As IntPtr, _
        ByVal psidGroup As IntPtr, _
        ByVal pDacl As IntPtr, _
        ByVal pSacl As IntPtr) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function SetNamedSecurityInfo( _
        ByVal pObjectName As String, _
        ByVal ObjectType As AclApi.SE_OBJECT_TYPE, _
        ByVal SecurityInformation As SECURITY_INFORMATION, _
        ByVal psidOwner As IntPtr, _
        ByVal psidGroup As IntPtr, _
        ByVal pDacl As IntPtr, _
        ByVal pSacl As IntPtr) As Integer
    End Function

    <DllImport("advapi32.DLL", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function GetSecurityDescriptorDacl( _
        ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByRef lpbDaclPresent As Integer, _
        ByRef pDacl As IntPtr, _
        ByRef lpbDaclDefaulted As Integer) As Integer
    End Function

    Public Shared Sub SetOwner(ByVal info As DirectoryInfo, ByVal sid As SecurityIdentifier)
        Dim security As DirectorySecurity
        KernelApi.AddPrivilege(KernelApi.SE_TAKE_OWNERSHIP_NAME)
        security = info.GetAccessControl(AccessControlSections.Owner)
        security.SetOwner(sid)
        info.SetAccessControl(security)
    End Sub

    Public Shared Sub SetOwner(ByVal key As RegistryKey, ByVal sid As SecurityIdentifier)
        Dim security As RegistrySecurity
        KernelApi.AddPrivilege(KernelApi.SE_TAKE_OWNERSHIP_NAME)
        security = key.GetAccessControl(AccessControlSections.Owner)
        security.SetOwner(sid)
        key.SetAccessControl(security)
    End Sub

    Public Shared Sub SetOwner(ByVal pathname As String, ByVal sid As SecurityIdentifier)
        SetOwner(New DirectoryInfo(pathname), sid)
    End Sub

    Public Shared Sub SetOwner(ByVal pathname As String, ByVal username As String)
        Dim account As New NTAccount(username)
        If account Is Nothing Then
            Throw New ArgumentNullException("The account was not found: " + username)
        End If
        Dim sid As SecurityIdentifier = CType(account.Translate(GetType(SecurityIdentifier)), SecurityIdentifier)
        SetOwner(pathname, sid)
    End Sub

    Public Shared Sub FullControlRegistryKey()

    End Sub

    Public Shared Sub SetOwnerRegistryKey(ByVal root As Integer, ByVal keyName As String, ByVal sid As SecurityIdentifier)
        Dim hObject As GCHandle
        Dim sd As New SECURITY_DESCRIPTOR()
        'Dim ptrSecDesc As IntPtr
        Dim hKey As IntPtr = IntPtr.Zero
        Dim result As Integer

        Try
            KernelApi.AddPrivilege(KernelApi.SE_TAKE_OWNERSHIP_NAME)

            'ptrSecDesc = Marshal.AllocHGlobal(Marshal.SizeOf(sd))
            'Marshal.StructureToPtr(sd, ptrSecDesc, True)

            Dim byteSid(sid.BinaryLength - 1) As Byte
            sid.GetBinaryForm(byteSid, 0)

            Dim ptrSid As IntPtr
            hObject = GCHandle.Alloc(byteSid, GCHandleType.Pinned)
            ptrSid = hObject.AddrOfPinnedObject()

            result = AclApi.RegOpenKeyEx(root, keyName, 0, AclApi.WRITE_OWNER, hKey)
            If result <> 0 Then
                Throw New AccessViolationException("Administrator privileges are required")
            End If

            result = AclApi.SetSecurityInfo(hKey, _
                SE_OBJECT_TYPE.SE_REGISTRY_KEY, _
                SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION, _
                ptrSid, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero)
            If result <> 0 Then
                Throw New AccessViolationException("Administrator privileges are required")
            End If
        Finally
            'Marshal.FreeHGlobal(ptrSecDesc)
            If hObject.IsAllocated Then
                hObject.Free()
            End If

            If hKey <> IntPtr.Zero Then
                result = AclApi.RegCloseKey(hKey)
            End If
        End Try
    End Sub

    Public Shared Sub SetOwnerApi(ByVal pathname As String, ByVal sid As SecurityIdentifier)
        Dim hObject As GCHandle

        Try
            KernelApi.AddPrivilege(KernelApi.SE_TAKE_OWNERSHIP_NAME)

            Dim byteSid(sid.BinaryLength - 1) As Byte
            sid.GetBinaryForm(byteSid, 0)

            Dim ptrSid As IntPtr
            hObject = GCHandle.Alloc(byteSid, GCHandleType.Pinned)
            ptrSid = hObject.AddrOfPinnedObject()

            Dim result As Integer
            result = AclApi.SetNamedSecurityInfo(pathname, _
                SE_OBJECT_TYPE.SE_FILE_OBJECT, _
                SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION, _
                ptrSid, ptrSid, IntPtr.Zero, IntPtr.Zero)
            If result = 0 Then
                'success
            Else
                Throw New AccessViolationException("Administrator privileges are required")
            End If
        Finally
            If hObject.IsAllocated Then
                hObject.Free()
            End If
        End Try
    End Sub

    Public Shared Sub SetFullControl(ByVal info As FileInfo, ByVal sid As SecurityIdentifier)
        Dim security As FileSecurity
        Dim rule As FileSystemAccessRule

        security = info.GetAccessControl(AccessControlSections.Access)
        'allow inheritance
        security.SetAccessRuleProtection(False, True)

        rule = New FileSystemAccessRule(sid, _
                        FileSystemRights.FullControl, _
                        InheritanceFlags.None, _
                        PropagationFlags.None, _
                        AccessControlType.Allow)
        security.AddAccessRule(rule)

        info.SetAccessControl(security)
    End Sub

    Public Shared Sub SetFullControl(ByVal rootKey As RegistryKey, ByVal subKeyName As String, ByVal sid As SecurityIdentifier, ByVal removeOrphans As Boolean)
        Using subKey As RegistryKey = rootKey.CreateSubKey(subKeyName)
            AclApi.SetFullControl(subKey, sid, removeOrphans)
        End Using
    End Sub

    Public Shared Sub SetFullControl(ByVal key As RegistryKey, ByVal sid As SecurityIdentifier, ByVal removeOrphans As Boolean)
        Dim security As RegistrySecurity
        Dim rule As RegistryAccessRule

        security = key.GetAccessControl(AccessControlSections.Access)
        'allow inheritance
        security.SetAccessRuleProtection(False, True)
        key.SetAccessControl(security)

        rule = New RegistryAccessRule(sid, _
                        RegistryRights.FullControl, _
                        InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, _
                        PropagationFlags.None, _
                        AccessControlType.Allow)
        security.ResetAccessRule(rule)

        key.SetAccessControl(security)
        If removeOrphans Then
            AclApi.RemoveOrphanedSids(key)
        End If
    End Sub

    Public Shared Sub RemoveOrphanedSids(ByVal info As DirectoryInfo)
        Dim doUpdate As Boolean
        Dim security As DirectorySecurity

        security = info.GetAccessControl(AccessControlSections.Access)
        For Each rule As FileSystemAccessRule In security.GetAccessRules(True, False, GetType(SecurityIdentifier))
            Dim sid As SecurityIdentifier
            sid = CType(rule.IdentityReference, SecurityIdentifier)
            If Not KernelApi.LookupAccountSid(sid, Nothing, Nothing, Nothing) Then
                security.PurgeAccessRules(sid)
                doUpdate = True
            End If
        Next
        If doUpdate Then
            info.SetAccessControl(security)
        End If
    End Sub

    Public Shared Sub RemoveOrphanedSids(ByVal info As FileInfo)
        Dim doUpdate As Boolean
        Dim security As FileSecurity

        security = info.GetAccessControl(AccessControlSections.Access)
        For Each rule As FileSystemAccessRule In security.GetAccessRules(True, False, GetType(SecurityIdentifier))
            Dim sid As SecurityIdentifier
            sid = CType(rule.IdentityReference, SecurityIdentifier)
            If Not KernelApi.LookupAccountSid(sid, Nothing, Nothing, Nothing) Then
                security.PurgeAccessRules(sid)
                doUpdate = True
            End If
        Next
        If doUpdate Then
            info.SetAccessControl(security)
        End If
    End Sub

    Public Shared Sub RemoveOrphanedSids(ByVal key As RegistryKey)
        Dim doUpdate As Boolean
        Dim security As RegistrySecurity

        security = key.GetAccessControl(AccessControlSections.Access)
        For Each rule As RegistryAccessRule In security.GetAccessRules(True, False, GetType(SecurityIdentifier))
            Dim sid As SecurityIdentifier
            sid = CType(rule.IdentityReference, SecurityIdentifier)
            If Not KernelApi.LookupAccountSid(sid, Nothing, Nothing, Nothing) Then
                security.PurgeAccessRules(sid)
                doUpdate = True
            End If
        Next
        If doUpdate Then
            key.SetAccessControl(security)
        End If
    End Sub

    Public Shared Sub SetOwnerRecursive(ByVal info As DirectoryInfo, ByVal sid As SecurityIdentifier)
        Try
            AclApi.SetOwner(info, sid)
            For Each subInfo As DirectoryInfo In info.GetDirectories()
                SetOwnerRecursive(subInfo, sid)
            Next
        Catch ex As Exception
            'Logger.WriteEntry("Unable to set owner: " + info.FullName + vbCrLf + ex.Message, EventLogEntryType.Warning)
        End Try
    End Sub

    Public Shared Sub SetFullControl(ByVal info As DirectoryInfo, ByVal sid As SecurityIdentifier)
        Dim security As DirectorySecurity
        Dim rule As FileSystemAccessRule

        security = info.GetAccessControl(AccessControlSections.Access)
        rule = New FileSystemAccessRule(sid, _
                        FileSystemRights.FullControl, _
                        InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, _
                        PropagationFlags.None, _
                        AccessControlType.Allow)
        security.ResetAccessRule(rule)
        info.SetAccessControl(security)
    End Sub

    Public Shared Sub RemoveAllAccess(ByVal info As DirectoryInfo)
        Dim security As DirectorySecurity
        security = info.GetAccessControl(AccessControlSections.Access)

        'remove all sids except system
        For Each rule As FileSystemAccessRule In security.GetAccessRules(True, True, GetType(SecurityIdentifier))
            If rule.IdentityReference <> KernelApi.GetSystemSid() Then
                security.PurgeAccessRules(rule.IdentityReference)
            End If
        Next

        info.SetAccessControl(security)
    End Sub

    Public Shared Sub ReplaceOwner(ByVal info As DirectoryInfo, ByVal oldSid As SecurityIdentifier, ByVal newSid As SecurityIdentifier)
        Dim ownerSid As SecurityIdentifier

        KernelApi.AddPrivilege(KernelApi.SE_TAKE_OWNERSHIP_NAME)

        Try
            Dim dirSec As DirectorySecurity
            dirSec = info.GetAccessControl(AccessControlSections.Owner)
            ownerSid = CType(dirSec.GetOwner(GetType(SecurityIdentifier)), SecurityIdentifier)
            If ownerSid = oldSid Then
                dirSec.SetOwner(newSid)
                info.SetAccessControl(dirSec)
            End If
            For Each subDir As DirectoryInfo In info.GetDirectories()
                ReplaceOwner(subDir, oldSid, newSid)
            Next
        Catch ex As Exception
            Logger.WriteEntry("Error replacing owner on: " + info.FullName + "; " + ex.Message)
        End Try

        Dim f As FileInfo = Nothing
        Try
            For Each f In info.GetFiles()
                Dim fileSec As FileSecurity
                fileSec = f.GetAccessControl(AccessControlSections.Owner)
                ownerSid = CType(fileSec.GetOwner(GetType(SecurityIdentifier)), SecurityIdentifier)
                If ownerSid = oldSid Then
                    fileSec.SetOwner(newSid)
                    f.SetAccessControl(fileSec)
                End If
            Next
        Catch ex As Exception
            If f Is Nothing Then
                Logger.WriteEntry("Error retrieving owner on: " + info.FullName + "; " + ex.Message)
            Else
                Logger.WriteEntry("Error replacing owner on: " + f.FullName + "; " + ex.Message)
            End If
        End Try
    End Sub
End Class

