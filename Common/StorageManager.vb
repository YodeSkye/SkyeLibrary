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

                        ' Grab candidate files starting with appName (e.g., "SkyeClip*.*")
                        Dim searchPattern As String = $"{appName}*.*"
                        Dim legacyFiles As String() = Directory.GetFiles(legacyRootPath, searchPattern, SearchOption.TopDirectoryOnly)

                        For Each filePath In legacyFiles
                            Dim fileName As String = Path.GetFileName(filePath)

#If DEBUG Then
                            ' IN DEBUG MODE: Only migrate DEV files (skips live production files)
                            If Not fileName.Contains("DEV", StringComparison.OrdinalIgnoreCase) Then Continue For
#Else
                            ' IN RELEASE MODE: Ignore DEV files completely (leaves your debug data alone)
                            If fileName.Contains("DEV", StringComparison.OrdinalIgnoreCase) Then Continue For
#End If

                            Dim destinationPath As String = Path.Combine(newPath, fileName)

                            If Not File.Exists(destinationPath) Then
                                File.Move(filePath, destinationPath)
                                Debug.WriteLine($"[Skye Migration] Moved {fileName} -> {destinationPath}")
                            End If
                        Next

                        ' Clean up legacy folder only if completely empty
                        If Directory.GetFiles(legacyRootPath).Length = 0 AndAlso Directory.GetDirectories(legacyRootPath).Length = 0 Then
                            Directory.Delete(legacyRootPath)
                        End If
                    End If
                Catch ex As Exception
                    Skye.Common.SafeLogWrite($"[Skye Migration Warning] {ex.Message}")
                End Try
            End Sub

        End Class

    End Class

End Namespace
