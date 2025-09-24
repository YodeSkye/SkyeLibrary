
''' <summary>
''' Common Utility Functions
''' </summary>
Public Class Common

    'Declarations
    ''' <summary>
    ''' File Size Units for FormatFileSize Function.
    ''' </summary>
    Public Enum FormatFileSizeUnits
        Auto
        Bytes
        KiloBytes
        MegaBytes
        GigaBytes
        TeraBytes
    End Enum

    'Functions
    ''' <summary>
    ''' Formats a file's or other object's size in bytes into a human-readable string using the specified units.
    ''' </summary>
    ''' <param name="sizeinbytes">The size in bytes</param>
    ''' <param name="unit">One of FormatFileSizeUnits</param>
    ''' <param name="decimalDigits">OPTIONAL the number of decimal places to return. Default is 2.</param>
    ''' <param name="omitThousandSeparators">OPTIONAL Whether to omit thousands separator. Default is False.</param>
    ''' <returns>A String representing the clean, human-readable size.</returns>
    Public Shared Function FormatFileSize(sizeinbytes As Long, unit As FormatFileSizeUnits, Optional decimalDigits As Integer = 2, Optional omitThousandSeparators As Boolean = False) As String

        'Simple Error Checking
        If sizeinbytes <= 0 Then Return "0 B"

        'Auto-Select Best Units Of Measure
        If unit = FormatFileSizeUnits.Auto Then
            Select Case sizeinbytes
                Case Is < 1024
                    unit = FormatFileSizeUnits.Bytes
                    decimalDigits = 0
                Case Is < 1048576 : unit = FormatFileSizeUnits.KiloBytes
                Case Is < 1073741824 : unit = FormatFileSizeUnits.MegaBytes
                Case Is < 1099511627776 : unit = FormatFileSizeUnits.GigaBytes
                Case Else : unit = FormatFileSizeUnits.TeraBytes
            End Select
        End If

        'Evaluate The Decimal Value
        Dim value As Decimal
        Dim suffix As String = ""
        Select Case unit
            Case FormatFileSizeUnits.Bytes
                value = CDec(sizeinbytes)
                suffix = " B"
            Case FormatFileSizeUnits.KiloBytes
                value = CDec(sizeinbytes / 1024)
                suffix = " KB"
            Case FormatFileSizeUnits.MegaBytes
                value = CDec(sizeinbytes / 1048576)
                suffix = " MB"
            Case FormatFileSizeUnits.GigaBytes
                value = CDec(sizeinbytes / 1073741824)
                suffix = " GB"
        End Select

        'Get The String Representation
        Dim format As String
        If omitThousandSeparators Then
            format = "F" & decimalDigits.ToString
        Else
            format = "N" & decimalDigits.ToString
        End If
        Return value.ToString(format) & suffix

    End Function

    ''' <summary>
    ''' Generates a log time string in the format HH:mm:ss.ffff (or HH:mm:ss if fractionalseconds is False)
    ''' representing the elapsed time between starttime and stoptime.
    ''' If starttime is greater than stoptime, it is assumed that the time period spans midnight.
    ''' </summary>
    ''' <param name="starttime">The time the operation started.</param>
    ''' <param name="stoptime">The time the operation ended.</param>
    ''' <param name="fractionalseconds">OPTIONAL show 4 decimal places on seconds. Default is True</param>
    ''' <returns>A human-readable String containing the total time of the operation.</returns>
    Public Shared Function GenerateLogTime(starttime As TimeSpan, stoptime As TimeSpan, Optional fractionalseconds As Boolean = True) As String
        Dim time As TimeSpan
        If starttime > stoptime Then
            time = stoptime + (New TimeSpan(24, 0, 0) - starttime)
        Else
            time = stoptime - starttime
        End If
        'If fractionalseconds Then
        '    Return New Date(time.Ticks).ToString("HH:mm:ss.ffff")
        'Else
        '    Return New Date(time.Ticks).ToString("HH:mm:ss")
        'End If
        If fractionalseconds Then
            Return $"{Math.Floor(time.TotalHours):00}:{time.Minutes:00}:{time.Seconds:00}.{time.Milliseconds:0000}"
        Else
            Return $"{Math.Floor(time.TotalHours):00}:{time.Minutes:00}:{time.Seconds:00}"
        End If
    End Function

    ''' <summary>
    ''' Returns a random integer between min and max, inclusive.
    ''' </summary>
    ''' <param name="min">The minimum value, inclusive.</param>
    ''' <param name="max">The maximum value, inclusive.</param>
    ''' <param name="current">OPTIONAL The current value you wish to avoid, Null by default</param>
    ''' <returns>A random Integer between min and max, inclusive, that is not equal to the current if provided.</returns>
    Public Shared Function GetRandom(ByVal min As Integer, ByVal max As Integer, Optional current As Integer? = Nothing) As Integer

        'Declarations
        Static Generator As System.Random = New System.Random() 'by making Generator static, we preserve the same instance between calls (i.e., do not create new instances with the same seed over and over)
        Dim result As Integer

        'Error Checking
        If max <> Integer.MaxValue Then max += 1
        If min >= max Then Return CInt(min)

        'Get Random Number
        Do
            result = Generator.Next(min, max)
        Loop While current.HasValue AndAlso result = current
        Return result

    End Function

End Class
