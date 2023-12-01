
' Create, install, and uninstall a hook for a modeless dialog box

' or PropertyManager page for intercepting keystrokes and

' handling accelerator keys
'为无模式对话框创建、安装和卸载挂钩"或用于拦截击键和 "处理快捷键的 PropertyManager"页面
Namespace WinHooks

    Public Class HookEventArgs

        Inherits EventArgs
        Public HookCode As Integer   ' Hook code
        Public wParam As IntPtr   ' WPARAM argument
        Public lParam As IntPtr  ' LPARAM argument

    End Class

    ' Hook Types  
    '挂钩类型
    Public Enum HookType

        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14

    End Enum

    Public Class LocalWindowsHook

        ' Filter function delegate 筛选器函数委托
        Public Delegate Function HookProc(ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

        ' Internal properties 内部属性
        Protected m_hhook As IntPtr = IntPtr.Zero
        Protected m_filterFunc As HookProc = Nothing
        Protected m_hookType As HookType

        ' Event delegate
        Public Delegate Sub HookEventHandler(ByVal sender As Object, ByVal e As HookEventArgs)

        ' Event: HookInvoked  钩子初始化
        Public Event HookInvoked As HookEventHandler

        Protected Sub OnHookInvoked(ByVal e As HookEventArgs)

            RaiseEvent HookInvoked(Me, e)

        End Sub

        ' Class constructor(s)类构造函数

        Public Sub New(ByVal hook As HookType)

            m_hookType = hook

            m_filterFunc = New HookProc(AddressOf Me.CoreHookProc)

        End Sub

        Public Sub New(ByVal hook As HookType, ByVal func As HookProc)

            m_hookType = hook

            m_filterFunc = func

        End Sub

        ' Default filter function

        Public Function CoreHookProc(ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

            If code < 0 Then Return CallNextHookEx(m_hhook, code, wParam, lParam)

            ' Let clients determine what to do 让客户决定该怎么做

            Dim e As HookEventArgs = New HookEventArgs()

            e.HookCode = code

            e.wParam = wParam

            e.lParam = lParam

            OnHookInvoked(e)

            ' Yield to the next hook in the chain 屈服于链中的下一个钩子

            CoreHookProc = CallNextHookEx(m_hhook, code, wParam, lParam)

        End Function

        Public Sub Install() 'Install the hook 安装挂钩

            m_hhook = SetWindowsHookEx(m_hookType, m_filterFunc, IntPtr.Zero, AppDomain.GetCurrentThreadId())

        End Sub

        Public Sub Uninstall() ' Uninstall the hook 卸载挂钩

            UnhookWindowsHookEx(m_hhook)

        End Sub

        ' Win32 Imports
        ' Win32: SetWindowsHookEx()
        Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal code As HookType, ByVal func As HookProc, ByVal hInstance As IntPtr, ByVal threadID As Integer) As Integer

        ' Win32: UnhookWindowsHookEx()
        Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhook As IntPtr) As Integer

        ' Win32: CallNextHookEx()
        Declare Function CallNextHookEx Lib "user32" (ByVal hhook As IntPtr, ByVal code As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    End Class

End Namespace