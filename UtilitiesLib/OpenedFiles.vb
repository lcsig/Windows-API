Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Text
Imports System.Threading
Imports Microsoft.Win32.SafeHandles

Public Class OpenedFiles

    Public Const PROCESS_DUP_HANDLE As Integer = &H40

    Public Class FileDetail
        Public ProcessId As Integer
        Public FileHandle As IntPtr
        Public FileName As String
    End Class

    Protected Enum NT_STATUS
        STATUS_SUCCESS = &H0
        STATUS_BUFFER_OVERFLOW = &H80000005
        STATUS_INFO_LENGTH_MISMATCH = &HC0000004
    End Enum

    Protected Enum SYSTEM_INFORMATION_CLASS
        SystemBasicInformation = 0
        SystemPerformanceInformation = 2
        SystemTimeOfDayInformation = 3
        SystemProcessInformation = 5
        SystemProcessorPerformanceInformation = 8
        SystemHandleInformation = 16
        SystemInterruptInformation = 23
        SystemExceptionInformation = 33
        SystemRegistryQuotaInformation = 37
        SystemLookasideInformation = 45
    End Enum

    Protected Enum OBJECT_INFORMATION_CLASS
        ObjectBasicInformation = 0
        ObjectNameInformation = 1
        ObjectTypeInformation = 2
        ObjectAllTypesInformation = 3
        ObjectHandleInformation = 4
    End Enum

    <Flags()> _
    Protected Enum ProcessAccessRights
        PROCESS_DUP_HANDLE = &H40
    End Enum

    <Flags()> _
    Protected Enum DuplicateHandleOptions
        DUPLICATE_CLOSE_SOURCE = &H1
        DUPLICATE_SAME_ACCESS = &H2
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure UNICODE_STRING
        Public Length As UShort
        Public MaximumLength As UShort
        Public Buffer As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SYSTEM_HANDLE_ENTRY
        Public OwnerPid As Integer
        Public ObjectType As Byte
        Public HandleFlags As Byte
        Public HandleValue As Short
        Public ObjectPointer As Integer
        Public AccessMask As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure GENERIC_MAPPING
        Public GenericRead As Integer
        Public GenericWrite As Integer
        Public GenericExecute As Integer
        Public GenericAll As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure OBJECT_NAME_INFORMATION
        Public Name As UNICODE_STRING
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure OBJECT_BASIC_INFORMATION
        Public Attributes As Integer
        Public GrantedAccess As Integer
        Public HandleCount As Integer
        Public PointerCount As Integer
        Public PagedPoolUsage As Integer
        Public NonPagedPoolUsage As Integer
        Public Reserved1 As Integer
        Public Reserved2 As Integer
        Public Reserved3 As Integer
        Public NameInformationLength As Integer
        Public TypeInformationLength As Integer
        Public SecurityDescriptorLength As Integer
        Public CreateTime As ComTypes.FILETIME
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure OBJECT_TYPE_INFORMATION
        Public Name As UNICODE_STRING
        Public ObjectCount As Integer
        Public HandleCount As Integer
        Public Reserved1 As Integer
        Public Reserved2 As Integer
        Public Reserved3 As Integer
        Public Reserved4 As Integer
        Public PeakObjectCount As Integer
        Public PeakHandleCount As Integer
        Public Reserved5 As Integer
        Public Reserved6 As Integer
        Public Reserved7 As Integer
        Public Reserved8 As Integer
        Public InvalidAttributes As Integer
        Public GenericMapping As GENERIC_MAPPING
        Public ValidAccess As Integer
        Public Unknown As Byte
        Public MaintainHandleDatabase As Byte
        Public PoolType As Integer
        Public PagedPoolUsage As Integer
        Public NonPagedPoolUsage As Integer
    End Structure


    <SecurityPermission(SecurityAction.LinkDemand, UnmanagedCode:=True)> _
    Protected Class SafeObjectHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid

        Private Sub New()
            MyBase.New(True)
        End Sub

        Protected Overrides Function ReleaseHandle() As Boolean
            Return NativeMethods.CloseHandle(MyBase.handle)
        End Function

        Protected Sub New(ByVal preexistingHandle As IntPtr, ByVal ownsHandle As Boolean)
            MyBase.New(ownsHandle)
            MyBase.SetHandle(preexistingHandle)
        End Sub
    End Class

    <SecurityPermission(SecurityAction.LinkDemand, UnmanagedCode:=True)> _
    Protected Class SafeProcessHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid

        Private Sub New()
            MyBase.New(True)
        End Sub

        Protected Overrides Function ReleaseHandle() As Boolean
            Return NativeMethods.CloseHandle(MyBase.handle)
        End Function

        Public Sub New(ByVal preexistingHandle As IntPtr, ByVal ownsHandle As Boolean)
            MyBase.New(ownsHandle)
            MyBase.SetHandle(preexistingHandle)
        End Sub
    End Class


    Protected Class NativeMethods

        <DllImport("ntdll.dll", SetLastError:=True)> _
        Public Shared Function NtQuerySystemInformation( _
                <[In]()> ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                <[In]()> ByVal SystemInformation As IntPtr, _
                <[In]()> ByVal SystemInformationLength As Integer, _
                <Out()> ByRef ReturnLength As Integer _
            ) As NT_STATUS
        End Function

        <DllImport("ntdll.dll", SetLastError:=True)> _
        Public Shared Function NtQueryObject( _
                <[In]()> ByVal Handle As IntPtr, _
                <[In]()> ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, _
                <[In]()> ByVal ObjectInformation As IntPtr, _
                <[In]()> ByVal ObjectInformationLength As Integer, _
                <Out()> ByRef ReturnLength As Integer _
            ) As NT_STATUS
        End Function

        <DllImport("ntdll.dll", SetLastError:=True)> _
        Public Shared Function NtQueryObject2( _
                <[In]()> ByVal Handle As IntPtr, _
                <[In]()> ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, _
                <[In]()> ByVal ObjectInformation As IntPtr, _
                <[In]()> ByVal ObjectInformationLength As Integer, _
                <Out()> ByRef ReturnLength As Integer _
            ) As SafeProcessHandle
        End Function

        <DllImport("kernel32.dll", SetLastError:=True, EntryPoint:="RtlZeroMemory")> _
        Public Shared Sub ZeroMemory( _
                <[In]()> ByVal Destination As IntPtr, _
                <[In]()> ByVal length As Integer _
            )
        End Sub

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function OpenProcess( _
                ByVal dwDesiredAccess As ProcessAccessRights, _
                ByVal bInheritHandle As Integer, _
                ByVal dwProcessId As Integer _
            ) As Integer
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function DuplicateHandle( _
                <[In]()> ByVal hSourceProcessHandle As Integer, _
                <[In]()> ByVal hSourceHandle As Integer, _
                <[In]()> ByVal hTargetProcessHandle As Integer, _
                <Out()> ByRef lpTargetHandle As Integer, _
                <[In]()> ByVal dwDesiredAccess As Integer, _
                <[In]()> ByVal bInheritHandle As Integer, _
                <[In]()> ByVal dwOptions As Integer _
            ) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function GetCurrentProcess() As Integer
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function GetProcessId(ByVal process As IntPtr) As Integer
        End Function

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)> _
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function CloseHandle(ByVal hObject As IntPtr) As Boolean
        End Function

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)> _
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function CloseHandle(ByVal hObject As Integer) As Boolean
        End Function

        <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)> _
        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function CloseHandle(ByVal hObject As SafeProcessHandle) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Public Shared Function QueryDosDevice( _
                ByVal lpDeviceName As String, _
                ByVal lpTargetPath As StringBuilder, _
                ByVal ucchMax As Integer _
            ) As Integer
        End Function

    End Class

    Protected Enum SystemHandleType
        OB_TYPE_UNKNOWN = 0
        OB_TYPE_TYPE = 1
        OB_TYPE_DIRECTORY
        OB_TYPE_SYMBOLIC_LINK
        OB_TYPE_TOKEN
        OB_TYPE_PROCESS
        OB_TYPE_THREAD
        OB_TYPE_UNKNOWN_7
        OB_TYPE_EVENT
        OB_TYPE_EVENT_PAIR
        OB_TYPE_MUTANT
        OB_TYPE_UNKNOWN_11
        OB_TYPE_SEMAPHORE
        OB_TYPE_TIMER
        OB_TYPE_PROFILE
        OB_TYPE_WINDOW_STATION
        OB_TYPE_DESKTOP
        OB_TYPE_SECTION
        OB_TYPE_KEY
        OB_TYPE_PORT
        OB_TYPE_WAITABLE_PORT
        OB_TYPE_UNKNOWN_21
        OB_TYPE_UNKNOWN_22
        OB_TYPE_UNKNOWN_23
        OB_TYPE_UNKNOWN_24
        OB_TYPE_IO_COMPLETION
        OB_TYPE_FILE
        'OB_TYPE_CONTROLLER
        'OB_TYPE_DEVICE
        'OB_TYPE_DRIVER
    End Enum

    Public Shared Function GetOpenHandles() As SYSTEM_HANDLE_ENTRY()
        Dim entries() As SYSTEM_HANDLE_ENTRY = Nothing
        Dim ptr As IntPtr = IntPtr.Zero

        Try
            RuntimeHelpers.PrepareConstrainedRegions()

            Dim ret As NT_STATUS
            Dim length As Integer = &H100
            Dim returnLength As Integer
            ptr = Marshal.AllocHGlobal(length)
            ret = NativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemHandleInformation, ptr, length, returnLength)

            'verify length is good
            Do While (ret = NT_STATUS.STATUS_INFO_LENGTH_MISMATCH)
                length = length + 64000
                Marshal.FreeHGlobal(ptr)
                ptr = Marshal.AllocHGlobal(length)
                ret = NativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemHandleInformation, ptr, length, returnLength)
            Loop

            If ret = NT_STATUS.STATUS_SUCCESS Then
                Dim handleCount As Integer = Marshal.ReadInt32(ptr)
                Dim offset As Integer = Marshal.SizeOf(GetType(System.Int32))
                Dim size As Integer = Marshal.SizeOf(GetType(SYSTEM_HANDLE_ENTRY))

                ReDim entries(handleCount - 1)

                Dim i As Integer = 0
                For i = 0 To handleCount - 1
                    Dim handleEntry As SYSTEM_HANDLE_ENTRY
                    handleEntry = CType(Marshal.PtrToStructure(New IntPtr(ptr.ToInt32() + offset), GetType(SYSTEM_HANDLE_ENTRY)), SYSTEM_HANDLE_ENTRY)

                    entries(i) = handleEntry

                    offset += size
                Next
            End If
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try

        Return entries
    End Function

    Public Shared Function GetOpenFileHandles() As List(Of FileDetail)

        Dim objectTypeNumber As Integer = OpenedFiles.GetObjectTypeNumber("File")

        GetOpenFileHandles = New List(Of FileDetail)

        Dim entries() As SYSTEM_HANDLE_ENTRY = OpenedFiles.GetOpenHandles()
        For Each entry As SYSTEM_HANDLE_ENTRY In entries
            If entry.ObjectType = objectTypeNumber Then ' And entry.OwnerPid = 2028 Then
                Dim d As FileDetail
                d = RetrieveObjectThreaded(entry)
                If Not String.IsNullOrEmpty(d.FileName) Then
                    GetOpenFileHandles.Add(d)
                End If
            End If
        Next
    End Function

    Private Shared Function GetObjectTypeNumber(ByVal typeName As String) As Integer
        Dim buffObjTypes As IntPtr = IntPtr.Zero
        Dim cbReqLength As Integer = 0
        Dim cTypeCount As Integer = 0
        Dim x As Integer = 0
        Dim lpTypeInfo As IntPtr = IntPtr.Zero
        Dim typeInfo As New OBJECT_TYPE_INFORMATION()
        Dim strType As String = ""
        If Utilities.IsVista() Then
            Dim b As IntPtr = Marshal.AllocHGlobal(4)
            NativeMethods.NtQueryObject(IntPtr.Zero, OBJECT_INFORMATION_CLASS.ObjectAllTypesInformation, b, 4, cbReqLength)
            Marshal.FreeHGlobal(b)
        Else
            NativeMethods.NtQueryObject(IntPtr.Zero, OBJECT_INFORMATION_CLASS.ObjectAllTypesInformation, IntPtr.Zero, 0, cbReqLength)
        End If

        buffObjTypes = Marshal.AllocHGlobal(cbReqLength)
        NativeMethods.NtQueryObject(IntPtr.Zero, OBJECT_INFORMATION_CLASS.ObjectAllTypesInformation, buffObjTypes, cbReqLength, cbReqLength)
        cTypeCount = Marshal.ReadInt32(buffObjTypes, 0)
        lpTypeInfo = New IntPtr(buffObjTypes.ToInt32() + 4)

        For x = 0 To cTypeCount - 1
            typeInfo = CType(Marshal.PtrToStructure(lpTypeInfo, typeInfo.GetType()), OBJECT_TYPE_INFORMATION)
            strType = Marshal.PtrToStringUni(typeInfo.Name.Buffer, typeInfo.Name.Length >> 1)
            If (typeName.Equals(strType)) Then
                Return x + 1
            End If

            lpTypeInfo = New IntPtr(lpTypeInfo.ToInt32() + 96 + typeInfo.Name.MaximumLength)
            lpTypeInfo = New IntPtr(lpTypeInfo.ToInt32() + (lpTypeInfo.ToInt32() Mod 4))
        Next

        Marshal.FreeHGlobal(buffObjTypes)

        If (typeName.Equals("Driver")) Then
            Return 24
        ElseIf (typeName.Equals("IoCompletion")) Then
            Return 25
        ElseIf (typeName.Equals("File")) Then
            Return 26
        End If
        Return 0
    End Function

    Public Shared Function RetrieveObjectThreaded(ByVal Handle As SYSTEM_HANDLE_ENTRY) As FileDetail
        Dim del As New DelegateRetrieveObject(AddressOf RetrieveObject)
        Dim iar As IAsyncResult
        iar = del.BeginInvoke(Handle, Nothing, Nothing)
        iar.AsyncWaitHandle.WaitOne(100, False)
        If iar.IsCompleted Then
            Return del.EndInvoke(iar)
        Else
            Return New FileDetail()
        End If
    End Function

    Public Delegate Function DelegateRetrieveObject(ByVal Handle As SYSTEM_HANDLE_ENTRY) As FileDetail

    Public Shared Function RetrieveObject(ByVal Handle As SYSTEM_HANDLE_ENTRY) As FileDetail
        Dim detail As New FileDetail()
        Dim length As Integer = 0
        Dim retv As Integer = 0
        Dim ret2 As Integer = 0
        Dim hHandle As Integer = 0

        Dim ObjBasic As New OBJECT_BASIC_INFORMATION()
        Dim ObjName As New OBJECT_NAME_INFORMATION()

        Dim m_ObjectTypeName As String
        Dim m_ObjectName As String
        Dim BufferObjName As IntPtr = IntPtr.Zero
        Dim BufferObjBasic As IntPtr = IntPtr.Zero

        Try
            OpenProcessForHandle(Handle.OwnerPid)

            NativeMethods.DuplicateHandle(m_hProcess, Handle.HandleValue, NativeMethods.GetCurrentProcess(), hHandle, 0, 0, DuplicateHandleOptions.DUPLICATE_SAME_ACCESS)
            If hHandle = 0 Then
                Return detail
            End If

            BufferObjBasic = Marshal.AllocHGlobal(Marshal.SizeOf(ObjBasic))
            ret2 = NativeMethods.NtQueryObject(New IntPtr(hHandle), OBJECT_INFORMATION_CLASS.ObjectBasicInformation, BufferObjBasic, Marshal.SizeOf(ObjBasic), retv)
            ObjBasic = CType(Marshal.PtrToStructure(BufferObjBasic, ObjBasic.GetType()), OBJECT_BASIC_INFORMATION)

            m_ObjectTypeName = "File"
            If ObjBasic.NameInformationLength = 0 Then
                length = 512
            Else
                length = ObjBasic.NameInformationLength + 2
            End If

            BufferObjName = Marshal.AllocCoTaskMem(length)
            NativeMethods.ZeroMemory(BufferObjName, length)
            ret2 = NativeMethods.NtQueryObject(New IntPtr(hHandle), OBJECT_INFORMATION_CLASS.ObjectNameInformation, BufferObjName, length, retv)
            ObjName = CType(Marshal.PtrToStructure(BufferObjName, ObjName.GetType()), OBJECT_NAME_INFORMATION)
            m_ObjectName = Marshal.PtrToStringUni(New IntPtr(BufferObjName.ToInt32() + 8))

            m_ObjectName = GetRegularFileNameFromDevice(m_ObjectName)

            'return FileDetail object
            detail.FileHandle = New IntPtr(Handle.HandleValue)
            detail.ProcessId = Handle.OwnerPid
            detail.FileName = m_ObjectName

            Return detail
        Finally
            Marshal.FreeHGlobal(BufferObjBasic)
            Marshal.FreeCoTaskMem(BufferObjName)
            NativeMethods.CloseHandle(hHandle)
        End Try
    End Function

    Private Shared m_lastPID As Integer
    Private Shared m_hProcess As Integer
    Public Shared Sub OpenProcessForHandle(ByVal ProcessID As Integer)
        If ProcessID <> m_lastPID Then
            NativeMethods.CloseHandle(m_hProcess)
            m_hProcess = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_DUP_HANDLE, 0, ProcessID)
            m_lastPID = ProcessID
        End If
    End Sub

    Public Shared Function GetRegularFileNameFromDevice(ByVal rawName As String) As String
        Dim filename As String = rawName

        For Each drivePath As String In Environment.GetLogicalDrives()
            Dim drive As String = drivePath.Substring(0, 2)
            Dim lpTargetPath As StringBuilder = New StringBuilder(260)

            If NativeMethods.QueryDosDevice(drive, lpTargetPath, 260) = 0 Then
                Return rawName
            End If

            Dim targetPath As String = lpTargetPath.ToString()

            If (filename.StartsWith(targetPath)) Then
                filename = filename.Replace(targetPath, drive)
                Return filename
            End If
        Next
        Return ""
    End Function

    Private Shared Function GetFileNameFromHandle(ByVal handle As IntPtr) As String
        Dim ptr As IntPtr = IntPtr.Zero
        Dim length As Integer = &H200

        Dim filename As String = ""

        Try
            RuntimeHelpers.PrepareConstrainedRegions()
            ptr = Marshal.AllocHGlobal(length)

            Dim ret As NT_STATUS
            ret = NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, length, length)

            'verify length is good
            Do While (ret = NT_STATUS.STATUS_BUFFER_OVERFLOW)
                length = length + 64000
                Marshal.FreeHGlobal(ptr)
                ptr = Marshal.AllocHGlobal(length)
                ret = NativeMethods.NtQueryObject(handle, OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, length, length)
            Loop

            If ret = NT_STATUS.STATUS_SUCCESS Then
                filename = Marshal.PtrToStringUni(New IntPtr((ptr.ToInt32() + 8)), CInt((length - 9) / 2))
            End If
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try
        Return filename
    End Function

End Class

