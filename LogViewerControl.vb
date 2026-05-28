
Imports System.IO
Imports System.Text

Namespace UI.Log

    Public Class LogViewerControl

        ' Declarations
        Public Structure LogViewerColors

            Public Back As Color
            Public Fore As Color

            Public TextBoxBack As Color
            Public TextBoxFore As Color

            Public ButtonBack As Color
            Public ButtonFore As Color

            Public TooltipBack As Color
            Public TooltipFore As Color
            Public TooltipBorder As Color

        End Structure
        Private _currentcolors As New LogViewerColors With {
                .Back = BackColor,
                .Fore = ForeColor,
                .TextBoxBack = BackColor,
                .TextBoxFore = ForeColor,
                .ButtonBack = BackColor,
                .ButtonFore = ForeColor,
                .TooltipBack = BackColor,
                .TooltipBorder = ForeColor,
                .TooltipFore = ForeColor
            }
        Private _lastSearchText As String = String.Empty
        Private _isBlinking As Boolean = False
        Private _highlightAll As Boolean = False

        Public Overrides Property Font As Font
            Get
                Return MyBase.Font
            End Get
            Set(value As Font)
                MyBase.Font = value
                ApplyFontToChildren(Me, value)
                Tip.Font = value
                Invalidate()
            End Set
        End Property

        ' Form Events
        Public Sub New()
            InitializeComponent()
            Skye.UI.ThemeManager.RegisterComponent(Tip)
            SetRTBPadding(5, 0)
        End Sub
        Private Sub LogViewerControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            RefreshContent()
            RTBLog.Select()
        End Sub

        ' Control Events
        Private Sub PanelTop_Paint(sender As Object, e As PaintEventArgs) Handles PanelTop.Paint
            Using p As New Pen(Color.FromArgb(200, 200, 200), 1)
                e.Graphics.DrawLine(p, 0, PanelTop.Height - 1, PanelTop.Width, PanelTop.Height - 1)
            End Using
        End Sub
        Private Sub TxtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBoxSearch.KeyDown
            If e.KeyCode = Keys.Enter AndAlso e.Shift Then
                e.Handled = True
                e.SuppressKeyPress = True
                StartSearchTopDown()
                If _highlightAll Then
                    ApplyHighlightAll()
                End If
            ElseIf e.KeyCode = Keys.Enter Then
                e.Handled = True
                e.SuppressKeyPress = True
                StartSearchBottomUp()
                If _highlightAll Then
                    ApplyHighlightAll()
                End If
            End If
        End Sub
        Private Sub TxtBoxSearch_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxSearch.TextChanged
            If String.IsNullOrWhiteSpace(TxtBoxSearch.Text) Then
                _lastSearchText = String.Empty
                _highlightAll = False
                ClearHighlights(False)
                UpdateHighlightButtonVisual()
            End If
        End Sub
        Private Sub BtnFindPrev_Click(sender As Object, e As EventArgs) Handles BtnPrevious.Click
            FindPrevious()
        End Sub
        Private Sub BtnFindNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
            FindNext()
        End Sub
        Private Sub BtnHighlightAll_Click(sender As Object, e As EventArgs) Handles BtnHighlightAll.Click
            If String.IsNullOrWhiteSpace(TxtBoxSearch.Text) Then
                IndicateNotFound()
                Return
            End If
            _highlightAll = Not _highlightAll
            UpdateHighlightButtonVisual()

            If _highlightAll Then
                ApplyHighlightAll()
            Else
                ClearHighlights()
            End If
        End Sub
        Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
            ClearHighlights()
            _highlightAll = False
            RefreshContent()
        End Sub
        Private Sub RTBLog_BackColorChanged(sender As Object, e As EventArgs) Handles RTBLog.BackColorChanged
            HandleColorChange()
        End Sub
        Private Sub RTBLog_ForeColorChanged(sender As Object, e As EventArgs) Handles RTBLog.ForeColorChanged
            HandleColorChange()
        End Sub

        ' Handlers

        ' Methods
        Public Sub RefreshContent()
            Try
                'Debug.Print($"Loading log from: {Skye.Common.Log.LogFilePath}")
                Dim path As String = Skye.Common.Log.LogFilePath
                If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                    RTBLog.Clear()
                    Return
                End If

                Using sr As New StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                    RTBLog.Text = sr.ReadToEnd()
                End Using
                RTBLog.SelectionStart = RTBLog.TextLength
                RTBLog.SelectionLength = 0
                RTBLog.ScrollToCaret()

            Catch
                ' Viewer must not throw
            End Try
        End Sub
        Private Sub StartSearchTopDown()
            Dim term = TxtBoxSearch.Text
            _lastSearchText = term

            Dim matches = GetAllMatches(term)
            If matches.Count = 0 Then
                IndicateNotFound()
                Return
            End If

            Dim first = matches(0)

            RTBLog.SelectionStart = first
            RTBLog.SelectionLength = term.Length
            RTBLog.ScrollToCaret()
        End Sub
        Private Sub StartSearchBottomUp()
            Dim term = TxtBoxSearch.Text
            _lastSearchText = term

            Dim matches = GetAllMatches(term)
            If matches.Count = 0 Then
                IndicateNotFound()
                Return
            End If

            Dim last = matches(matches.Count - 1)

            RTBLog.SelectionStart = last
            RTBLog.SelectionLength = term.Length
            RTBLog.ScrollToCaret()
        End Sub
        Private Sub FindNext()
            Dim term = _lastSearchText
            If String.IsNullOrWhiteSpace(term) Then
                StartSearchTopDown()
                Return
            End If

            Dim matches = GetAllMatches(term)
            If matches.Count = 0 Then Return

            Dim current = RTBLog.SelectionStart

            ' Find the first match AFTER the current one
            For Each m In matches
                If m > current Then
                    RTBLog.SelectionStart = m
                    RTBLog.SelectionLength = term.Length
                    RTBLog.ScrollToCaret()
                    Return
                End If
            Next

            ' If none after, stay at last match
            Dim last = matches(matches.Count - 1)
            RTBLog.SelectionStart = last
            RTBLog.SelectionLength = term.Length
            RTBLog.ScrollToCaret()
        End Sub
        Private Sub FindPrevious()
            Dim term = _lastSearchText
            If String.IsNullOrWhiteSpace(term) Then
                StartSearchTopDown()
                Return
            End If

            Dim matches = GetAllMatches(term)
            If matches.Count = 0 Then Return

            Dim current = RTBLog.SelectionStart

            Dim prev As Integer = -1

            For Each m In matches
                If m < current Then
                    prev = m
                Else
                    Exit For
                End If
            Next

            If prev = -1 Then
                prev = matches(0)
            End If

            RTBLog.SelectionStart = prev
            RTBLog.SelectionLength = term.Length
            RTBLog.ScrollToCaret()
        End Sub
        Private Sub ApplyHighlightAll()
            Dim term As String = TxtBoxSearch.Text
            If String.IsNullOrWhiteSpace(term) Then
                ClearHighlights()
                Return
            End If
            Dim selStart = RTBLog.SelectionStart
            Dim selLength = RTBLog.SelectionLength
            RTBLog.SuspendLayout()

            ClearHighlights(False)

            Dim index As Integer = 0
            While index < RTBLog.TextLength
                index = RTBLog.Find(term, index, RichTextBoxFinds.None)
                If index = -1 Then Exit While

                RTBLog.SelectionStart = index
                RTBLog.SelectionLength = term.Length
                Dim highlight = Blend(RTBLog.ForeColor, RTBLog.BackColor, 0.6)
                highlight = Color.FromArgb(120, highlight)
                RTBLog.SelectionBackColor = highlight

                index += term.Length
            End While

            ' Restore user selection
            RTBLog.SelectionStart = selStart
            RTBLog.SelectionLength = selLength

            RTBLog.ResumeLayout()
        End Sub

        ' Helpers
        Private Function GetAllMatches(term As String) As List(Of Integer)
            Dim results As New List(Of Integer)
            If String.IsNullOrWhiteSpace(term) Then Return results

            Dim text As String = RTBLog.Text
            Dim idx As Integer = text.IndexOf(term, 0, StringComparison.OrdinalIgnoreCase)

            While idx <> -1
                results.Add(idx)
                idx = text.IndexOf(term, idx + term.Length, StringComparison.OrdinalIgnoreCase)
            End While

            Return results
        End Function
        Private Sub ApplyFontToChildren(parent As Control, f As Font)
            For Each ctl As Control In parent.Controls
                ctl.Font = f
                If ctl.HasChildren Then
                    ApplyFontToChildren(ctl, f)
                End If
            Next
        End Sub
        Private Async Sub IndicateNotFound()
            If _isBlinking Then Return ' prevent overlap
            _isBlinking = True

            Dim oldColor = TxtBoxSearch.BackColor
            For i As Integer = 1 To 3
                TxtBoxSearch.BackColor = Color.MistyRose
                Await Task.Delay(120)
                TxtBoxSearch.BackColor = oldColor
                Await Task.Delay(120)
            Next

            _isBlinking = False
        End Sub
        Private Sub SetRTBPadding(left As Integer, right As Integer)
            Dim wParam As Integer = Skye.WinAPI.EC_LEFTMARGIN Or Skye.WinAPI.EC_RIGHTMARGIN
            Dim lParam As Integer = (right << 16) Or (left And &HFFFF)
            Skye.WinAPI.SendMessage(RTBLog.Handle, Skye.WinAPI.EM_SETMARGINS, CType(wParam, IntPtr), CType(lParam, IntPtr))
        End Sub
        Private Sub UpdateHighlightButtonVisual()
            If _highlightAll Then
                'BtnHighlightAll.BackColor = Blend(_currentcolors.ButtonBack, _currentcolors.ButtonFore, 0.6)
                BtnHighlightAll.BackColor = BlendWithMinimumContrast(_currentcolors.ButtonBack, _currentcolors.ButtonFore, 0.25, 40) ' minimum brightness delta
                BtnHighlightAll.ForeColor = _currentcolors.ButtonFore
                BtnHighlightAll.FlatAppearance.BorderColor = Blend(_currentcolors.ButtonBack, _currentcolors.ButtonFore, 0.4)
            Else
                BtnHighlightAll.BackColor = _currentcolors.ButtonBack
                BtnHighlightAll.ForeColor = _currentcolors.ButtonFore
                BtnHighlightAll.FlatAppearance.BorderColor = _currentcolors.ButtonFore
            End If
        End Sub
        Public Sub ApplyColors(colors As LogViewerColors)
            _currentcolors = colors

            ' Backgrounds
            BackColor = colors.Back
            ForeColor = colors.Fore
            RTBLog.BackColor = colors.Back
            RTBLog.ForeColor = colors.Fore

            ' Search box
            TxtBoxSearch.BackColor = colors.TextBoxBack
            TxtBoxSearch.ForeColor = colors.TextBoxFore

            ' Buttons
            BtnPrevious.BackColor = colors.ButtonBack
            BtnPrevious.ForeColor = colors.ButtonFore
            BtnNext.BackColor = colors.ButtonBack
            BtnNext.ForeColor = colors.ButtonFore
            UpdateHighlightButtonVisual()
            BtnRefresh.BackColor = colors.ButtonBack
            BtnRefresh.ForeColor = colors.ButtonFore

            ' Tooltip
            Tip.BackColor = colors.TooltipBack
            Tip.ForeColor = colors.TooltipFore
            Tip.BorderColor = colors.TooltipBorder

        End Sub
        Private Sub ClearHighlights(Optional restoreSelection As Boolean = True)
            Dim selStart = RTBLog.SelectionStart
            Dim selLength = RTBLog.SelectionLength

            RTBLog.SelectAll()
            RTBLog.SelectionBackColor = RTBLog.BackColor
            RTBLog.SelectionColor = RTBLog.ForeColor
            RTBLog.SelectionLength = 0
            RTBLog.SelectionStart = RTBLog.TextLength
            RTBLog.ScrollToCaret()

            If restoreSelection Then
                RTBLog.SelectionStart = selStart
                RTBLog.SelectionLength = selLength
            End If
        End Sub
        Private Function Blend(c1 As Color, c2 As Color, amount As Double) As Color
            ' amount = how much of c2 to add in (0.0–1.0)
            Dim a = Math.Max(0.0, Math.Min(1.0, amount))
            Return Color.FromArgb(
                        CInt(c1.R * (1 - a) + c2.R * a),
                        CInt(c1.G * (1 - a) + c2.G * a),
                        CInt(c1.B * (1 - a) + c2.B * a)
                    )
        End Function
        Private Function BlendWithMinimumContrast(baseColor As Color, mixColor As Color, amount As Double, minDelta As Integer) As Color
            ' First do the normal blend
            Dim result = Blend(baseColor, mixColor, amount)

            ' Compute brightness difference
            Dim delta = Math.Abs(result.GetBrightness() - baseColor.GetBrightness())

            ' If too subtle, force a stronger shift
            If delta < (minDelta / 255.0) Then
                ' Decide direction: lighter or darker
                If baseColor.GetBrightness() > 0.5 Then
                    ' Base is light → push darker
                    result = ControlPaint.Dark(baseColor, 0.2)
                Else
                    ' Base is dark → push lighter
                    result = ControlPaint.Light(baseColor, 0.2)
                End If
            End If

            Return result
        End Function
        Private Sub SetRedraw(ctrl As Control, enable As Boolean)
            Skye.WinAPI.SendMessage(ctrl.Handle, Skye.WinAPI.WM_SETREDRAW, CType(If(enable, 1, 0), IntPtr), IntPtr.Zero)
        End Sub
        Private Sub HandleColorChange()
            SetRedraw(RTBLog, False)
            '
            'Clear old highlights because they were blended with the old theme
            ClearHighlights()

            ' If toggle is ON, reapply highlight using new theme colors
            If _highlightAll Then
                ApplyHighlightAll()
            End If

            SetRedraw(RTBLog, True)
            RTBLog.Invalidate()
        End Sub

    End Class

End Namespace