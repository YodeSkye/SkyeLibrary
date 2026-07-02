
Imports System.Runtime.InteropServices
Imports System.Text

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        Public Const GW_HWNDFIRST As UInteger = 0
        Public Const GW_HWNDLAST As UInteger = 1
        Public Const GW_HWNDNEXT As UInteger = 2
        Public Const GW_HWNDPREV As UInteger = 3
        Public Const GW_OWNER As UInteger = 4
        Public Const GW_CHILD As UInteger = 5
        Public Const GWL_HWNDPARENT As Integer = -8
        Public Const GWL_STYLE As Integer = -16
        Public Const GWL_EXSTYLE As Integer = -20
        Public Const GWLP_WNDPROC As Integer = -4
        Public Const GWLP_HINSTANCE As Integer = -6
        Public Const GWLP_ID As Integer = -12
        Public Const WS_VISIBLE As Integer = &H10000000
        Public Const WS_CHILD As Integer = &H40000000
        Public Const WS_POPUP As Integer = &H80000000
        Public Const WS_MINIMIZEBOX As Integer = 131072 '&H20000 'Turn on the WS_MINIMIZEBOX style flag for borderless windows so you can minimize/restore from the taskbar.
        Public Const WS_MAXIMIZE As Integer = 16777216 '&H1000000
        Public Const WS_EX_TOOLWINDOW As Integer = 128 '&H80
        Public Const WS_EX_TOPMOST As Integer = &H8
        Public Const WS_EX_APPWINDOW As Integer = &H40000
        Public Const WS_EX_NOACTIVATE As Integer = &H8000000
        Public Const WS_EX_TRANSPARENT As Integer = &H20
        Public Const RDW_INVALIDATE As UInteger = &H1
        Public Const RDW_ERASE As UInteger = &H4
        Public Const RDW_ALLCHILDREN As UInteger = &H80
        Public Const RDW_UPDATENOW As UInteger = &H100
        Public Const RDW_FRAME As UInteger = &H400
        Public Const SW_HIDE As Integer = 0
        Public Const SW_SHOWNORMAL As Integer = 1
        Public Const SW_SHOWMINIMIZED As Integer = 2
        Public Const SW_SHOWMAXIMIZED As Integer = 3
        Public Const SW_SHOWNOACTIVATE As Integer = 4
        Public Const SWP_NOACTIVATE As UInteger = &H10
        Public Const SWP_SHOWWINDOW As UInteger = &H40
        Public Const SWP_HIDEWINDOW As UInteger = &H80
        Public Const SWP_NOZORDER As UInteger = &H4
        Public Const SWP_NOMOVE As UInteger = &H2
        Public Const SWP_NOSIZE As UInteger = &H1
        Public Const SWP_FRAMECHANGED As UInteger = &H20
        Public Shared ReadOnly HWND_TOP As New IntPtr(0)
        Public Shared ReadOnly HWND_BOTTOM As New IntPtr(1)
        Public Shared ReadOnly HWND_TOPMOST As New IntPtr(-1)
        Public Shared ReadOnly HWND_NOTOPMOST As New IntPtr(-2)
        <StructLayout(LayoutKind.Sequential)>
        Public Structure COMBOBOXINFO
            Public cbSize As Integer
            Public rcItem As RECT
            Public rcButton As RECT
            Public stateButton As Integer
            Public hwndCombo As IntPtr
            Public hwndEdit As IntPtr
            Public hwndList As IntPtr
        End Structure

        ' API FUNCTIONS
        Public Delegate Function WndProcDelegate(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
        <DllImport("user32.dll")>
        Public Shared Function GetComboBoxInfo(hWnd As IntPtr, ByRef pcbi As COMBOBOXINFO) As Boolean
        End Function

        ' METHODS
        ''' <summary>
        ''' To Remove A Borderless Window From Windows TaskSwitcher
        ''' </summary>
        ''' <param name="handle">Your Form's Handle</param>
        ''' <returns>True Is Success, Otherwise False</returns>
        Public Shared Function HideFormInTaskSwitcher(handle As IntPtr) As Boolean
            Try
                Dim exStyle As Integer = GetWindowLong(handle, GWL_EXSTYLE)
                Dim result As Integer = SetWindowLong(handle, GWL_EXSTYLE, exStyle Or WS_EX_TOOLWINDOW)
                Return result <> 0
            Catch
                Return False
            End Try
        End Function
        ''' <summary>
        ''' Gets all window handles for a given process ID.
        ''' </summary>
        ''' <param name="pid">The process ID.</param>
        ''' <returns>A list of window handles.</returns>
        Public Shared Function GetWindowsForProcess(pid As Integer) As List(Of IntPtr)
            Dim results As New List(Of IntPtr)
            EnumWindows(Function(hWnd, lParam)
                            Dim windowPid As UInteger
                            Dim threadId As UInteger = GetWindowThreadProcessId(hWnd, windowPid)
                            If windowPid = pid Then
                                results.Add(hWnd)
                            End If
                            Return True
                        End Function,
                    IntPtr.Zero)
            Return results
        End Function
        ''' <summary>
        ''' Gets the caption text of a window given its handle.
        ''' </summary>
        ''' <param name="hWnd">The handle of the window.</param>
        ''' <returns>The caption text of the window.</returns>
        Public Shared Function GetCaption(hWnd As IntPtr) As String
            Dim sb As New StringBuilder(512)
            Dim length As Integer = GetWindowText(hWnd, sb, sb.Capacity)
            If length > 0 Then
                Return sb.ToString()
            Else
                Return String.Empty
            End If
        End Function
        ''' <summary>
        ''' Automatically selects the appropriate GetWindowLong function based on the process architecture (32-bit or 64-bit).
        ''' </summary>
        ''' <param name="hWnd">The handle of the window.</param>
        ''' <param name="nIndex">The zero-based offset of the value to be retrieved.</param>
        ''' <returns>The requested window information.</returns>
        Public Shared Function GetWindowLongAuto(hWnd As IntPtr, nIndex As Integer) As IntPtr
            If IntPtr.Size = 4 Then
                Return New IntPtr(GetWindowLong(hWnd, nIndex))
            Else
                Return GetWindowLongPtr(hWnd, nIndex)
            End If
        End Function
        ''' <summary>
        ''' Automatically selects the appropriate SetWindowLong function based on the process architecture (32-bit or 64-bit).
        ''' </summary>
        ''' <param name="hWnd">The handle of the window.</param>
        ''' <param name="nIndex">The zero-based offset of the value to be set.</param>
        ''' <param name="dwNewLong">The new value to be set.</param>
        ''' <returns>The previous value of the specified window information.</returns>
        Public Shared Function SetWindowLongAuto(hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
            If IntPtr.Size = 4 Then
                Return New IntPtr(SetWindowLong(hWnd, nIndex, dwNewLong.ToInt32()))
            Else
                Return SetWindowLongPtr(hWnd, nIndex, dwNewLong)
            End If
        End Function

    End Class

End Namespace
#Enable Warning CA1401
