Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class BackgroundApp

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As DwmWindowAttribute, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, EntryPoint:="SetWinEventHook", CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function SetWinEventHook(eventMin As UInteger, eventMax As UInteger, hmodWinEventProc As IntPtr, lpfnWinEventProc As WinEventDelegate, idProcess As UInteger, idThread As UInteger, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, EntryPoint:="UnhookWinEvent", CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function UnhookWinEvent(hWinEventHook As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Public Delegate Sub WinEventDelegate(hWinEventHook As IntPtr, eventType As UInteger, hWnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)

    Enum DwmWindowAttribute As UInteger
        NCRenderingEnabled = 1
        NCRenderingPolicy
        TransitionsForceDisabled
        AllowNCPaint
        CaptionButtonBounds
        NonClientRtlLayout
        ForceIconicRepresentation
        Flip3DPolicy
        ExtendedFrameBounds
        HasIconicBitmap
        DisallowPeek
        ExcludedFromPeek
        Cloak
        Cloaked
        FreezeRepresentation
        PassiveUpdateMode
        UseHostBackdropBrush
        UseImmersiveDarkMode = 20
        WindowCornerPreference = 33
        BorderColor
        CaptionColor
        TextColor
        VisibleFrameBorderThickness
        SystemBackdropType
        Last
    End Enum

    Enum DwmWindowCornerPreference As Long
        DoNotRound = 1
        Round = 2
        SemiRound = 3
    End Enum

    Private Const WINEVENT_OUTOFCONTEXT As UInteger = 0
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = 3

    Private ReadOnly WinEventProcDelegate As New WinEventDelegate(AddressOf WinEventProc)
    Private ReadOnly WinEventHook As IntPtr = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, WinEventProcDelegate, 0, 0, WINEVENT_OUTOFCONTEXT)
    Private ReadOnly TaskSchedeuler As New Scheduler("NoRoundCorners11", "cmd430", "Starts NoRoundCorners11 on Logon", Nothing, Nothing, False)

    Private Sub BackgroundApp_Load(sender As Object, e As EventArgs) Handles Me.Load
        If TaskSchedeuler.GetTask() Is Nothing Then
            TaskSchedeuler.AddTask()
        Else
            TaskSchedeuler.UpdateTask()
        End If
        If LoadWithWindows() Then
            StartWithWindowsToolStripMenuItem.Checked = True
        End If
        AddHandler StartWithWindowsToolStripMenuItem.CheckedChanged, AddressOf StartWithWindowsToolStripMenuItem_CheckedChanged
    End Sub

    Public Function LoadWithWindows() As Boolean
        Return TaskSchedeuler.GetTask().Enabled
    End Function

    Private Sub WinEventProc(hWinEventHook As IntPtr, eventType As UInteger, HWND As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
        DwmSetWindowAttribute(HWND, DwmWindowAttribute.WindowCornerPreference, DwmWindowCornerPreference.DoNotRound, Marshal.SizeOf(GetType(Integer)))
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not LoadWithWindows() Then
            TaskSchedeuler.DeleteTask()
        End If
        UnhookWinEvent(WinEventHook)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub StartWithWindowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartWithWindowsToolStripMenuItem.Click
        StartWithWindowsToolStripMenuItem.Checked = False
    End Sub

    Private Sub StartWithWindowsToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs)
        RemoveHandler StartWithWindowsToolStripMenuItem.CheckedChanged, AddressOf StartWithWindowsToolStripMenuItem_CheckedChanged
        If LoadWithWindows() Then
            TaskSchedeuler.ToggleTask()
            StartWithWindowsToolStripMenuItem.Checked = False
        Else
            TaskSchedeuler.ToggleTask()
            StartWithWindowsToolStripMenuItem.Checked = True
        End If
        AddHandler StartWithWindowsToolStripMenuItem.CheckedChanged, AddressOf StartWithWindowsToolStripMenuItem_CheckedChanged
    End Sub

End Class
