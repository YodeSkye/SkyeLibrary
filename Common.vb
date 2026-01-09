
Imports System.Text
Imports System.Text.RegularExpressions

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
        Static Generator As New System.Random() 'by making Generator static, we preserve the same instance between calls (i.e., do not create new instances with the same seed over and over)
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

    ''' <summary>
    ''' Returns a truncated version of the input string with an ellipsis appended 
    ''' if its length exceeds the specified maximum. Null values return an empty string.
    ''' </summary>
    ''' <param name="s">The input string to truncate.</param>
    ''' <param name="n">The maximum allowed length of the string.</param>
    ''' <returns>
    ''' The original string if its length is less than or equal to <paramref name="n"/>, 
    ''' otherwise the first <paramref name="n"/> characters followed by "…".
    ''' </returns>
    Public Shared Function Trunc(s As String, n As Integer) As String
        If s Is Nothing Then Return ""
        If s.Length <= n Then Return s
        Return String.Concat(s.AsSpan(0, n), "…")
    End Function

    ''' <summary>
    ''' Removes trailing ANSI null (0x00) bytes from the end of a byte array.
    ''' </summary>
    ''' <param name="bytes">The input byte array to trim.</param>
    ''' <returns>
    ''' A new byte array with trailing null bytes removed. 
    ''' Returns an empty array if the input is null, empty, or contains only nulls.
    ''' </returns>
    Public Shared Function TrimAnsiNull(bytes As Byte()) As Byte()
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return Array.Empty(Of Byte)()
        Dim i = Array.FindLastIndex(bytes, Function(b) b <> 0)
        If i < 0 Then Return Array.Empty(Of Byte)()
        Dim out(i + 1) As Byte
        Buffer.BlockCopy(bytes, 0, out, 0, i + 1)
        Return out
    End Function

    ''' <summary>
    ''' Removes trailing Unicode null pairs (0x0000) from the end of a byte array.
    ''' </summary>
    ''' <param name="bytes">The input byte array to trim.</param>
    ''' <returns>
    ''' A new byte array with trailing 0x0000 pairs removed. 
    ''' Returns an empty array if the input is null, empty, or contains only null pairs.
    ''' </returns>
    Public Shared Function TrimUnicodeNull(bytes As Byte()) As Byte()
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return Array.Empty(Of Byte)()
        ' Remove trailing 0x0000 pair
        Dim len = bytes.Length
        While len >= 2 AndAlso bytes(len - 2) = 0 AndAlso bytes(len - 1) = 0
            len -= 2
        End While
        Dim out(len - 1) As Byte
        If len > 0 Then Buffer.BlockCopy(bytes, 0, out, 0, len)
        Return out
    End Function

    ''' <summary>
    ''' Extracts plain text from an RTF-encoded byte array by loading it into a RichTextBox.
    ''' </summary>
    ''' <param name="bytes">The RTF-encoded byte array.</param>
    ''' <returns>
    ''' The plain text content of the RTF. Returns an empty string if the input is null or empty.
    ''' </returns>
    Public Shared Function ExtractPlainTextFromRTF(bytes As Byte()) As String
        Dim rtfString = Encoding.Default.GetString(bytes)
        Dim rtfBox As New RichTextBox With {.Rtf = rtfString}
        Return rtfBox.Text
    End Function

    ''' <summary>
    ''' Extracts plain text from an HTML-encoded byte array by locating the StartFragment/EndFragment markers.
    ''' </summary>
    ''' <param name="bytes">The HTML-encoded byte array.</param>
    ''' <returns>
    ''' The plain text content between <!--StartFragment--> and <!--EndFragment--> if found,
    ''' otherwise returns the entire HTML string.
    ''' </returns>
    Public Shared Function ExtractPlainTextFromHTML(bytes As Byte()) As String
        Dim htmlString = Encoding.Default.GetString(bytes)
        Dim startIdx = htmlString.IndexOf("<!--StartFragment-->")
        Dim endIdx = htmlString.IndexOf("<!--EndFragment-->")
        If startIdx >= 0 AndAlso endIdx > startIdx Then
            Return htmlString.Substring(startIdx + 20, endIdx - (startIdx + 20))
        End If
        Return htmlString
    End Function

    ''' <summary>
    ''' Converts an RTF string to plain text.
    ''' </summary>
    ''' <param name="rtf">The RTF string to convert.</param>
    ''' <returns>The plain text representation of the RTF string.</returns>
    Public Shared Function RTFToPlainText(rtf As String) As String
        Using rtb As New RichTextBox()
            rtb.Rtf = rtf
            Return rtb.Text
        End Using
    End Function

    ''' <summary>
    ''' Converts an HTML string to plain text by stripping HTML tags.
    ''' </summary>
    ''' <param name="html">The HTML string to convert.</param>
    ''' <returns>The plain text representation of the HTML string.</returns>
    Public Shared Function HTMLToPlainText(html As String) As String
        Dim noTags = Regex.Replace(html, "<.*?>", " ")
        Return Regex.Replace(noTags, "\s+", " ").Trim()
    End Function

End Class
