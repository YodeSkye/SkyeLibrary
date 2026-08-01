Imports System.IO
Imports System.Reflection
Imports System.Threading

Namespace Skye

    Partial Public Class Common

        Public Class StorageManager

            Private Const PUBLISHER_FOLDER As String = "Skye"
            Private Const MAX_RETRIES As Integer = 3
            Private Const RETRY_DELAY_MS As Integer = 1000 ' 1 second delay between attempts

            ''' <summary>
            ''' Resolves storage path to: LocalAppData\Skye\&lt;AppName&gt;
            ''' Returns the valid path string on success, or String.Empty if the path is inaccessible after retries.
            ''' </summary>
            Public Shared Function GetAppDirectory() As String
                Dim appName As String = GetExecutingAppName()
                Dim primaryPath As String = GetPrimaryPath(appName)

                ' Ensure path ends with a trailing backslash so "UserPath & fileName" works cleanly
                If Not primaryPath.EndsWith(Path.DirectorySeparatorChar) Then
                    primaryPath &= Path.DirectorySeparatorChar
                End If

                ' Step 1: Migrate legacy Documents\Skye files for this app if present
                MigrateLegacyData(primaryPath, appName)

                ' Step 2: Test write access with a retry loop for transient locks
                For attempt As Integer = 1 To MAX_RETRIES
                    If CanWriteToDirectory(primaryPath) Then
                        Return primaryPath
                    End If

                    ' If write failed, pause briefly before trying again
                    If attempt < MAX_RETRIES Then
                        Thread.Sleep(RETRY_DELAY_MS)
                    End If
                Next

                ' Step 3: Edge Case Catch — Log the failure and return String.Empty for the app to handle
                Skye.Common.SafeLogWrite($"Storage Manager Error: Unable to access or write to path after {MAX_RETRIES} attempts: {primaryPath}")
                Return String.Empty
            End Function

            Private Shared Function GetPrimaryPath(appName As String) As String
                Dim localAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)

                If String.IsNullOrWhiteSpace(localAppData) Then
                    localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA")
                End If

                Return Path.Combine(localAppData, PUBLISHER_FOLDER, appName)
            End Function
            Private Shared Function GetExecutingAppName() As String
                Try

                    Dim entryAssembly As Assembly = Assembly.GetEntryAssembly()
                    If entryAssembly IsNot Nothing Then
                        Return entryAssembly.GetName().Name
                    End If

                    Dim processName As String = System.Diagnostics.Process.GetCurrentProcess().ProcessName
                    Return processName.Replace(".vshost", "")

                Catch
                    Return "DefaultApp"
                End Try
            End Function
            Private Shared Function CanWriteToDirectory(path As String) As Boolean
                Try
                    Directory.CreateDirectory(path)
                    Dim testFile As String = IO.Path.Combine(path, $".write_test_{Guid.NewGuid():N}.tmp")
                    File.WriteAllText(testFile, "test")
                    File.Delete(testFile)
                    Return True
                Catch
                    Return False
                End Try
            End Function

            Private Shared Sub MigrateLegacyData(newPath As String, appName As String)
                Try
                    Dim docsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    Dim legacyRootPath As String = Path.Combine(docsPath, PUBLISHER_FOLDER)

                    If Directory.Exists(legacyRootPath) Then
                        Directory.CreateDirectory(newPath)

                        ' Grab ALL files starting with appName (with or without extensions)
                        Dim searchPattern As String = $"{appName}*"
                        Dim legacyFiles As String() = Directory.GetFiles(legacyRootPath, searchPattern, SearchOption.TopDirectoryOnly)

                        ' RUNTIME CHECK: Detect if the calling/entry application is a Debug build
                        Dim isCallingAppDebug As Boolean = IsAssemblyDebug(Assembly.GetEntryAssembly())

                        For Each filePath In legacyFiles
                            Dim fileName As String = Path.GetFileName(filePath)
                            Dim hasDevInName As Boolean = fileName.Contains("DEV", StringComparison.OrdinalIgnoreCase)

                            If isCallingAppDebug Then
                                ' RUNNING IN DEBUG MODE: Only migrate files containing "DEV"
                                If Not hasDevInName Then Continue For
                            Else
                                ' RUNNING IN RELEASE MODE: Ignore "DEV" files completely
                                If hasDevInName Then Continue For
                            End If

                            Dim destinationPath As String = Path.Combine(newPath, fileName)

                            If File.Exists(filePath) AndAlso Not File.Exists(destinationPath) Then
                                File.Move(filePath, destinationPath)
                                Skye.Common.SafeLogWrite($"Storage Manager Moved {fileName} -> {destinationPath}")
                            End If
                        Next

                        ' Clean up legacy folder only if completely empty
                        If Directory.GetFiles(legacyRootPath).Length = 0 AndAlso Directory.GetDirectories(legacyRootPath).Length = 0 Then
                            Try
                                ' 1. Ensure our process isn't holding current directory lock
                                Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory)
                                ' 2. Wait for any locks to clear
                                System.Threading.Thread.Sleep(100)
                                ' 3. Clear attributes on the folder
                                Dim dirInfo As New DirectoryInfo(legacyRootPath)
                                dirInfo.Attributes = FileAttributes.Normal
                                ' 4. Attempt removal
                                dirInfo.Delete(False)
                                Skye.Common.SafeLogWrite($"Storage Manager Migration removed empty legacy directory: {legacyRootPath}")
                            Catch ex As IOException
                                ' Folder either wasn't empty yet (other app files remain) or OS handle was briefly busy.
                                ' Non-critical — leaves it alone.
                            Catch ex As Exception
                                ' Handles permission or unexpected exceptions silently
                            End Try
                        End If
                    End If
                Catch ex As Exception
                    Skye.Common.SafeLogWrite($"Storage Manager Migration Warning {ex.Message}")
                End Try
            End Sub
            ''' <summary>
            ''' Inspects an assembly at runtime to see if it was built in Debug mode.
            ''' </summary>
            Private Shared Function IsAssemblyDebug(asm As Assembly) As Boolean
                If asm Is Nothing Then Return False

                Try
                    Dim attribs = asm.GetCustomAttributes(GetType(System.Diagnostics.DebuggableAttribute), False)
                    If attribs IsNot Nothing AndAlso attribs.Length > 0 Then
                        Dim debugAttr = CType(attribs(0), System.Diagnostics.DebuggableAttribute)
                        ' Checks if JIT tracking/optimization is configured for Debugging
                        Return debugAttr.IsJITTrackingEnabled OrElse debugAttr.DebuggingFlags.HasFlag(System.Diagnostics.DebuggableAttribute.DebuggingModes.DisableOptimizations)
                    End If
                Catch
                End Try

                Return False
            End Function

        End Class

    End Class

End Namespace
