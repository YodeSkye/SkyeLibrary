Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        Public Const MAX_PATH As Integer = 260
        Public Const SHGFI_ICON As Integer = 256
        Public Const SHGFI_SMALLICON As Integer = 1
        Public Const SHGFI_LARGEICON As Integer = 0
        Public Const SHGFI_DISPLAYNAME As Integer = 512
        Public Const SHGFI_TYPENAME As Integer = 1024
        Public Const SHGFI_USEFILEATTRIBUTES As Integer = &H10
        Public Const FILE_ATTRIBUTE_NORMAL As Integer = &H80

        Public Structure SHFILEINFO
            Dim hIcon As IntPtr
            Dim iIcon As Integer
            Dim dwAttributes As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)>
            Dim szDisplayName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
            Dim szTypeName As String
        End Structure

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function SHGetFileInfo(pszPath As String,
                                             dwFileAttributes As Integer,
                                             ByRef psfi As SHFILEINFO,
                                             cbFileInfo As Integer,
                                             uFlags As Integer) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        ''' <summary>
        ''' Retrieves the icon associated with a file path, optionally returning the large or small version.
        ''' </summary>
        ''' <param name="filepath">The full path to the file.</param>
        ''' <param name="getLargeIcon">True to retrieve the large icon; False for small.</param>
        ''' <returns>The extracted icon, or Nothing if retrieval fails.</returns>
        Public Shared Function GetApplicationIcon(filepath As String, Optional getlargeicon As Boolean = False) As Icon
            filepath = filepath.Substring(0, Math.Min(filepath.Length, MAX_PATH))
            Dim shinfo As New SHFILEINFO()
            Dim flags As Integer = SHGFI_ICON Or If(getlargeicon, SHGFI_LARGEICON, SHGFI_SMALLICON)

            Dim status As IntPtr = SHGetFileInfo(filepath, 0, shinfo, Marshal.SizeOf(shinfo), flags)
            If status = IntPtr.Zero OrElse shinfo.hIcon = IntPtr.Zero Then Return Nothing

            Try
                Using tempIcon = Icon.FromHandle(shinfo.hIcon)
                    Return DirectCast(tempIcon.Clone(), Icon)
                End Using
            Finally
                DestroyIcon(shinfo.hIcon)
            End Try

        End Function

    End Class

End Namespace
#Enable Warning CA1401
