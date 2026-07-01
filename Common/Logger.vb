
Namespace Skye

    Partial Public Class Common

        ' Logger
        ''' <summary>
        ''' Provides a simple logging mechanism that writes timestamped messages to a log file. 
        ''' The log file is automatically rotated when it exceeds a specified number of lines.
        ''' </summary>
        Public Class Log

            ' Declarations
            Private Shared _appName As String
            Private Shared _logFilePath As String
            Private Shared _lineCount As Integer
            Private Shared _maxLines As Integer = 5000
            Private Shared _viewer As Skye.UI.Log.LogViewer
            Private Shared ReadOnly _lock As New Object()

            ' Public Properties
            Public Shared Property LogFilePath As String
                Get
                    If String.IsNullOrWhiteSpace(_logFilePath) Then
                        Throw New InvalidOperationException("Log.Initialize must be called before accessing LogFilePath.")
                    End If
                    Return _logFilePath
                End Get
                Set(value As String)
                    If String.IsNullOrWhiteSpace(value) Then
                        Throw New ArgumentException("LogFilePath cannot be empty.")
                    End If

                    _logFilePath = value
                    EnsureFolderExistsIfNeeded()
                    RecountLines()
                End Set
            End Property
            Public Shared ReadOnly Property LineCount As Integer
                Get
                    Return _lineCount
                End Get
            End Property
            Public Shared Property MaxLines As Integer
                Get
                    Return _maxLines
                End Get
                Set(value As Integer)
                    If value < 100 Then
                        Throw New ArgumentException("MaxLines must be at least 100.")
                    End If
                    _maxLines = value
                End Set
            End Property

            ' Public Methods
            Public Shared Sub Initialize(appName As String, Optional folder As String = Nothing)
                If String.IsNullOrWhiteSpace(appName) Then
                    Throw New ArgumentException("appName cannot be empty.")
                End If

                _appName = appName

                If String.IsNullOrWhiteSpace(folder) Then
                    Dim temp As String = IO.Path.GetTempPath()
                    _logFilePath = IO.Path.Combine(temp, $"{_appName}Log.txt")
                Else
                    EnsureFolderExists(folder)
                    _logFilePath = IO.Path.Combine(folder, $"{_appName}Log.txt")
                End If

                RecountLines()
            End Sub
            Public Shared Sub Write(message As String)
                If String.IsNullOrWhiteSpace(_logFilePath) Then
                    Throw New InvalidOperationException("Log.Initialize must be called before writing logs.")
                End If

                If message Is Nothing Then
                    message = String.Empty
                End If

                SyncLock _lock
                    Try
                        ' Count lines in the message
                        Dim lineCountInMessage As Integer = message.Split({vbCrLf}, StringSplitOptions.None).Length

                        ' Rotate if needed
                        If _lineCount + lineCountInMessage >= _maxLines Then
                            Rotate()
                        End If

                        ' Build timestamped first line
                        Dim ts As String = DateTime.Now.ToString("yyyy-MM-dd @ HH:mm:ss")
                        Dim firstLine As String = $"{ts} --> "

                        ' Write to file
                        Using writer As New IO.StreamWriter(_logFilePath, append:=True)
                            writer.Write(firstLine)
                            writer.Write(message)
                            writer.Write(vbCrLf)
                        End Using

                        ' Update line counter
                        _lineCount += lineCountInMessage

                    Catch
                        ' Logging must NEVER throw
                    End Try
                End SyncLock
            End Sub
            Public Shared Sub Clear()
                SyncLock _lock
                    Try
                        If String.IsNullOrWhiteSpace(_logFilePath) Then Return

                        ' Overwrite the file with nothing
                        IO.File.WriteAllText(_logFilePath, String.Empty)

                        ' Reset line counter
                        _lineCount = 0

                    Catch
                        ' Logging must NEVER throw
                    End Try
                End SyncLock
            End Sub
            Public Shared Sub OpenLocation()
                Try
                    If String.IsNullOrWhiteSpace(_logFilePath) Then Return

                    Dim argument As String = "/select,""" & _logFilePath & """"
                    Process.Start("explorer.exe", argument)

                Catch
                    ' Never throw from logging utilities
                End Try
            End Sub
            Public Shared Sub ShowViewer()
                Try
                    If _viewer Is Nothing OrElse _viewer.IsDisposed Then
                        _viewer = New Skye.UI.Log.LogViewer()
                    End If

                    If _viewer.Visible Then
                        _viewer.BringToFront()
                    Else
                        _viewer.StartPosition = FormStartPosition.CenterScreen
                        _viewer.Show()
                    End If
                Catch
                    ' Never throw from logging utilities
                End Try
            End Sub

            ' Helpers
            Private Shared Sub EnsureFolderExists(path As String)
                If Not IO.Directory.Exists(path) Then
                    IO.Directory.CreateDirectory(path)
                End If
            End Sub
            Private Shared Sub EnsureFolderExistsIfNeeded()
                Dim folder As String = IO.Path.GetDirectoryName(_logFilePath)
                If Not String.IsNullOrWhiteSpace(folder) Then
                    EnsureFolderExists(folder)
                End If
            End Sub
            Private Shared Sub RecountLines()
                If IO.File.Exists(_logFilePath) Then
                    _lineCount = IO.File.ReadLines(_logFilePath).Count()
                Else
                    _lineCount = 0
                End If
            End Sub
            Private Shared Sub Rotate()
                Try
                    If IO.File.Exists(_logFilePath) Then
                        Dim folder As String = IO.Path.GetDirectoryName(_logFilePath)
                        Dim baseName As String = IO.Path.GetFileNameWithoutExtension(_logFilePath)
                        Dim stamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                        Dim overflowName As String = $"{baseName}_{stamp}.txt"
                        Dim overflowPath As String = IO.Path.Combine(folder, overflowName)
                        IO.File.Move(_logFilePath, overflowPath)
                    End If
                Catch
                    ' Rotation must NEVER throw
                End Try
                _lineCount = 0
            End Sub

        End Class

        ' Self, Safe Logger
        ''' <summary>
        ''' Provides a simple logging mechanism that is safe to call from anywhere in the library without risk of throwing exceptions.
        ''' </summary>
        Friend Shared Sub SafeLogWrite(message As String)
            If String.IsNullOrWhiteSpace(Log.LogFilePath) Then Return ' Logging not initialized, so we can't log this message. Just return.
            Try
                Log.Write("SKYELIBRARY --> " & message)
            Catch
                ' Swallow any exceptions to ensure this method is safe to call from anywhere
            End Try
        End Sub

    End Class

End Namespace
