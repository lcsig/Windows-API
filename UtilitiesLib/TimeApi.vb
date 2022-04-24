Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Microsoft.Win32.SafeHandles
Imports System.Runtime.InteropServices.ComTypes

Public Class TimeApi

    Public Structure SYSTEMTIME
        Dim Year As Integer
        Dim Month As Integer
        Dim DayOfWeek As Integer
        Dim Day As Integer
        Dim Hour As Integer
        Dim Minute As Integer
        Dim Second As Integer
        Dim Milliseconds As Integer
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function FileTimeToLocalFileTime( _
        <[In]()> ByRef lpFileTime As ComTypes.FILETIME, _
        <Out()> ByRef lpLocalFileTime As ComTypes.FILETIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function FileTimeToSystemTime( _
        <[In]()> ByRef lpFileTime As ComTypes.FILETIME, _
        <Out()> ByRef lpSystemTime As SYSTEMTIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SystemTimeToFileTime( _
        <[In]()> ByRef lpSystemTime As SYSTEMTIME, _
        <Out()> ByRef lpFileTime As ComTypes.FILETIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function LocalFileTimeToFileTime( _
        <[In]()> ByRef lpLocalFileTime As ComTypes.FILETIME, _
        <Out()> ByRef lpFileTime As ComTypes.FILETIME) As Boolean
    End Function

    Public Shared Function DateTimeToFileTime(ByVal d As DateTime) As ComTypes.FILETIME
        Dim ft As New ComTypes.FILETIME()
        Dim tmp As Long = d.ToFileTime()
        'Logger.WriteEntry(d.ToString())
        'Dim A As Integer = CInt(tmp) 'And &HFFFFFFFF
        ft.dwLowDateTime = CInt(tmp >> 32) 'And &HFFFFFFFF
        ft.dwHighDateTime = CInt(tmp >> 32)
        Return ft
    End Function

    'Public Shared Function DateTimeToFileTime(ByVal d As DateTime) As ComTypes.FILETIME
    '    Dim ft As New ComTypes.FILETIME()
    '    Dim tmp As Long = d.ToFileTime()
    '    ft.dwLowDateTime = CInt(tmp And &HFFFFFFFF)
    '    ft.dwHighDateTime = CInt(tmp >> 32)
    '    Return ft
    'End Function

    Public Shared Function DateTimeToFileTimeUtc(ByVal d As DateTime) As ComTypes.FILETIME
        Dim ft As New ComTypes.FILETIME()
        Dim tmp As Long = d.ToFileTimeUtc()
        ft.dwLowDateTime = CInt(tmp And &HFFFFFFFF)
        ft.dwHighDateTime = CInt(tmp >> 32)
        Return ft
    End Function

    Public Shared Function FileTimeToDateTime(ByVal ft As ComTypes.FILETIME) As DateTime
        Dim dt As DateTime = DateTime.MaxValue
        Dim tmp As Long
        tmp = (CLng(ft.dwHighDateTime) << 32) + ft.dwLowDateTime

        Try
            dt = DateTime.FromFileTime(tmp)
        Catch ex As ArgumentOutOfRangeException
            dt = DateTime.MaxValue
        End Try
        Return dt
    End Function

    Public Shared Function FileTimeToDateTimeUtc(ByVal ft As ComTypes.FILETIME) As DateTime
        Dim dt As DateTime = DateTime.MaxValue
        Dim tmp As Long
        tmp = (CLng(ft.dwHighDateTime) << 32) + ft.dwLowDateTime

        Try
            dt = DateTime.FromFileTimeUtc(tmp)
        Catch ex As ArgumentOutOfRangeException
            dt = DateTime.MaxValue
        End Try
        Return dt
    End Function

    'Public Shared Function SystemTimeToFileTimeUtc(ByVal st As SYSTEMTIME) As FILETIME
    '    Dim ft As New FILETIME
    '    If Not TimeApi.SystemTimeToFileTime(st, ft) Then
    '        Throw New Win32Exception(Marshal.GetLastWin32Error())
    '    End If
    '    If LocalFileTimeToFileTime(ft, ft) Then
    '        Throw New Win32Exception(Marshal.GetLastWin32Error())
    '    End If
    '    Return ft
    'End Function

    'Public Shared Function DateTimeToSystemTime(ByVal d As DateTime) As SYSTEMTIME
    '    Dim st As SYSTEMTIME
    '    st.Month = d.Month - 1
    '    st.Day = d.Day - 1
    '    st.DayOfWeek = d.DayOfWeek - 1
    '    st.Year = d.Year - 1
    '    st.Hour = d.Hour - 1
    '    st.Minute = d.Minute - 1
    '    st.Second = d.Second - 1
    '    st.Milliseconds = d.Millisecond - 1
    '    Return st
    'End Function

End Class
