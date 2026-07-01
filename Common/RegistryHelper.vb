
Imports Microsoft.Win32

Namespace Skye

    Partial Public Class Common

        ''' <summary>
        ''' Provides helper methods for reading and writing application settings 
        ''' under HKEY_CURRENT_USER in the Windows Registry.
        ''' 
        ''' Before using any methods, your application must assign a value to the 
        ''' <see cref="BaseKey"/> property. This should be a full registry path 
        ''' beginning with "Software\", for example: "Software\YourAppName".
        ''' 
        ''' All values are stored beneath this BaseKey, ensuring each application 
        ''' has its own isolated settings root.
        ''' </summary>
        Public Class RegistryHelper

            Private Shared _baseKey As String = Nothing
            Public Shared Property BaseKey As String
                Get
                    If String.IsNullOrWhiteSpace(_baseKey) Then
                        Throw New InvalidOperationException("RegistryHelper.BaseKey must be set before using registry functions.")
                    End If
                    Return _baseKey
                End Get
                Set(value As String)
                    If String.IsNullOrWhiteSpace(value) Then
                        Throw New ArgumentException("BaseKey cannot be empty.")
                    End If
                    _baseKey = value
                End Set
            End Property

            ' Methods
            Public Shared Function GetInt(name As String, defaultValue As Integer) As Integer
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Return CInt(key.GetValue(name, defaultValue))
                End Using
            End Function
            Public Shared Sub SetInt(name As String, value As Integer)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, value, RegistryValueKind.DWord)
                End Using
            End Sub

            Public Shared Function GetBool(name As String, defaultValue As Boolean) As Boolean
                Return GetInt(name, If(defaultValue, 1, 0)) = 1
            End Function
            Public Shared Sub SetBool(name As String, value As Boolean)
                SetInt(name, If(value, 1, 0))
            End Sub

            Public Shared Function GetString(name As String, defaultValue As String) As String
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Return CStr(key.GetValue(name, defaultValue))
                End Using
            End Function
            Public Shared Sub SetString(name As String, value As String)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, value, RegistryValueKind.String)
                End Using
            End Sub

            Public Shared Function GetDateTime(name As String, defaultValue As DateTime) As DateTime
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Dim obj = key.GetValue(name, Nothing)

                    If TypeOf obj Is Long Then
                        Return New DateTime(CType(obj, Long))
                    End If

                    Return defaultValue
                End Using
            End Function
            Public Shared Sub SetDateTime(name As String, value As DateTime)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, value.Ticks, RegistryValueKind.QWord)
                End Using
            End Sub

            Public Shared Function GetBytes(name As String) As Byte()
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Dim obj = key.GetValue(name)
                    If TypeOf obj Is Byte() Then
                        Return CType(obj, Byte())
                    End If
                    Return Nothing
                End Using
            End Function
            Public Shared Sub SetBytes(name As String, data As Byte())
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, data, RegistryValueKind.Binary)
                End Using
            End Sub

            Public Shared Function GetLong(name As String, defaultValue As Long) As Long
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Dim obj = key.GetValue(name, Nothing)
                    If TypeOf obj Is Long Then
                        Return CType(obj, Long)
                    End If
                    Return defaultValue
                End Using
            End Function
            Public Shared Sub SetLong(name As String, value As Long)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, value, RegistryValueKind.QWord)
                End Using
            End Sub

            Public Shared Function GetStringArray(name As String, defaultValue As String()) As String()
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Dim obj = key.GetValue(name, Nothing)
                    If TypeOf obj Is String() Then
                        Return CType(obj, String())
                    End If
                    Return defaultValue
                End Using
            End Function
            Public Shared Sub SetStringArray(name As String, values As String())
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.SetValue(name, values, RegistryValueKind.MultiString)
                End Using
            End Sub

            Public Shared Function ValueExists(name As String) As Boolean
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Return key.GetValue(name) IsNot Nothing
                End Using
            End Function
            Public Shared Function GetValueNames() As String()
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Return key.GetValueNames()
                End Using
            End Function
            Public Shared Function GetValuesWithPrefix(prefix As String) As IEnumerable(Of String)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    Return key.GetValueNames().Where(Function(n) n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToArray()
                End Using
            End Function
            Public Shared Sub DeleteValue(name As String)
                Using key = Registry.CurrentUser.CreateSubKey(BaseKey)
                    key.DeleteValue(name, False)
                End Using
            End Sub

            Public Shared Function GetStringFromHKCU(subKey As String, valueName As String, defaultValue As String) As String
                Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(subKey, writable:=False)
                    If key Is Nothing Then Return defaultValue
                    Dim v = key.GetValue(valueName)
                    If v Is Nothing Then Return defaultValue
                    Return CStr(v)
                End Using
            End Function
            Public Shared Sub SetStringInHKCU(subKey As String, valueName As String, valueData As String)
                Using key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(subKey)
                    key.SetValue(valueName, valueData, Microsoft.Win32.RegistryValueKind.String)
                End Using
            End Sub
            Public Shared Sub DeleteValueInHKCU(subKey As String, valueName As String)
                Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(subKey, writable:=True)
                    key?.DeleteValue(valueName, throwOnMissingValue:=False)
                End Using
            End Sub

        End Class

    End Class

End Namespace
