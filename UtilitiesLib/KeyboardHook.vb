Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class KeyboardHook

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function UnhookWindowsHookEx( _
        ByVal hHook As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function SetWindowsHookEx( _
        ByVal idHook As Integer, _
        ByVal lpfn As KeyboardHookDelegate, _
        ByVal hmod As Integer, _
        ByVal dwThreadId As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function GetAsyncKeyState( _
        ByVal vKey As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function CallNextHookEx( _
        ByVal hHook As Integer, _
        ByVal nCode As Integer, _
        ByVal wParam As Integer, _
        ByRef lParam As KBDLLHOOKSTRUCT) As Integer
    End Function

    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    ' Low-Level Keyboard Constants
    Private Const HC_ACTION As Integer = 0
    Private Const LLKHF_EXTENDED As Integer = &H1
    Private Const LLKHF_INJECTED As Integer = &H10
    Private Const LLKHF_ALTDOWN As Integer = &H20
    Private Const LLKHF_UP As Integer = &H80

    ' Virtual Keys
    Public Const VK_TAB As Integer = &H9
    Public Const VK_CONTROL As Integer = &H11
    Public Const VK_ESCAPE As Integer = &H1B
    Public Const VK_DELETE As Integer = &H2E
    Public Const VK_MENU As Integer = &H12

    Private Const WH_KEYBOARD_LL As Integer = 13&
    Private hKeyboard As Integer

    Public Event KeyUp(ByVal key As Keys, ByVal ctrl As Boolean, ByVal alt As Boolean, ByVal shift As Boolean)

    Public Function KeyboardCallback(ByVal Code As Integer, _
        ByVal wParam As Integer, _
        ByRef lParam As KBDLLHOOKSTRUCT) As Integer

        Try
            Dim keyUp As Boolean = CBool(lParam.flags And LLKHF_UP)
            If keyUp Then
                Dim alt As Boolean = (GetAsyncKeyState(Keys.Menu) <> 0)
                Dim ctrl As Boolean = (GetAsyncKeyState(Keys.ControlKey) <> 0)
                Dim shift As Boolean = (GetAsyncKeyState(Keys.ShiftKey) <> 0)

                RaiseEvent KeyUp(CType(lParam.vkCode, Keys), ctrl, alt, shift)
            End If

            Return CallNextHookEx(hKeyboard, Code, wParam, lParam)
        Catch ex As Exception
            Logger.WriteEntry("Error in KeyboardCallback: " + ex.Message, EventLogEntryType.Error)
        End Try

    End Function

    Public Delegate Function KeyboardHookDelegate( _
        ByVal Code As Integer, _
        ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) _
        As Integer

    <MarshalAs(UnmanagedType.FunctionPtr)> Private callback As KeyboardHookDelegate

    Public Sub HookKeyboard()
        Dim hInstance As IntPtr
        hInstance = Marshal.GetHINSTANCE([Assembly].GetExecutingAssembly().GetModules()(0))
        callback = New KeyboardHookDelegate(AddressOf KeyboardCallback)

        hKeyboard = SetWindowsHookEx( _
            WH_KEYBOARD_LL, callback, _
            hInstance.ToInt32, 0)

        Call CheckHooked()
    End Sub

    Public Sub CheckHooked()
        If (Hooked()) Then
            Logger.WriteEntry("Keyboard hooked", EventLogEntryType.Information)
        Else
            Logger.WriteEntry("Keyboard hook failed: " & Err.LastDllError, EventLogEntryType.Error)
        End If
    End Sub

    Private Function Hooked() As Boolean
        Hooked = hKeyboard <> 0
    End Function

    Public Sub UnhookKeyboard()
        If (Hooked()) Then
            Call UnhookWindowsHookEx(hKeyboard)
        End If
    End Sub

End Class
