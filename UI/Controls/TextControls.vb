
Imports System.ComponentModel
Imports System.Text

Namespace Skye.UI

	''' <summary>
	''' Changes basic .NET label to OPTIONALLY copy on double-click.
	''' Also adds text orientation options: Horizontal (normal), Vertical (rotated), and Vertical Stacked (characters top-to-bottom).
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class Label
		Inherits System.Windows.Forms.Label

		' Declarations
		Public Enum TextOrientation
			Horizontal
			Vertical
			Stacked
		End Enum
		Private _orientation As TextOrientation = TextOrientation.Horizontal
		Public Event CustomDraw As PaintEventHandler

		' Properties
		<DefaultValue(False)>
		Public Property CopyOnDoubleClick As Boolean
		<DefaultValue(False)>
		Public Shadows Property AutoSize As Boolean
			Get
				Return MyBase.AutoSize
			End Get
			Set(value As Boolean)
				MyBase.AutoSize = value
				If value Then
					AdjustSizeForOrientation()
				End If
			End Set
		End Property
		<Category("Layout")>
		<Description("Controls how text is rendered: Horizontal (Normal), Vertical (Rotated), or Vertical Stacked (Characters Top-to-Bottom). NOTE: Paint Event only fires when Orientation is set to Horizontal.")>
		<DefaultValue(TextOrientation.Horizontal)>
		Public Property Orientation As TextOrientation
			Get
				Return _orientation
			End Get
			Set(value As TextOrientation)
				If _orientation <> value Then
					_orientation = value
					AdjustSizeForOrientation()
					Invalidate()   ' force redraw
					Update()       ' optional: immediate refresh
				End If
			End Set
		End Property

		' Events
		Protected Overrides Sub DefWndProc(ByRef m As System.Windows.Forms.Message)
			If LicenseManager.UsageMode = LicenseUsageMode.Runtime Then
				Select Case m.Msg
					Case WinAPI.WM_LBUTTONDBLCLK
						' Suppress default double-click behavior unless explicitly allowed
						If CopyOnDoubleClick Then
							MyBase.DefWndProc(m)
						Else
							m.Result = IntPtr.Zero
						End If
					Case Else
						MyBase.DefWndProc(m)
				End Select
			End If
		End Sub
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			If Orientation = TextOrientation.Horizontal Then
				MyBase.OnPaint(e) ' normal label behavior
			Else
				' your custom vertical/stacked drawing
				MyBase.OnPaintBackground(e)

				Dim g As Graphics = e.Graphics
				Dim sf As StringFormat = GetStringFormatFromContentAlignment(TextAlign)

				Select Case Orientation
					Case TextOrientation.Horizontal
						g.DrawString(Text, Font, New SolidBrush(Me.ForeColor), ClientRectangle, sf)
					Case TextOrientation.Vertical
						g.TranslateTransform(CSng(Width / 2), CSng(Height / 2))
						g.RotateTransform(-90)
						g.DrawString(Text, Font, New SolidBrush(Me.ForeColor), 0, 0, sf)
						g.ResetTransform()
					Case TextOrientation.Stacked
						Dim chars() As Char = Me.Text.ToCharArray()
						Dim lineHeight As Single = TextRenderer.MeasureText("X", Me.Font).Height
						Dim y As Single = ((Height - (chars.Length * lineHeight)) / 2) + lineHeight / 2
						For Each c As Char In chars
							g.DrawString(c.ToString(), Font, New SolidBrush(ForeColor), New PointF(CSng(Width / 2), y), sf)
							y += lineHeight
						Next
				End Select

			End If
		End Sub
		Protected Overrides Sub OnTextChanged(e As EventArgs)
			MyBase.OnTextChanged(e)
			If Me.AutoSize Then AdjustSizeForOrientation()
		End Sub
		Protected Overrides Sub OnFontChanged(e As EventArgs)
			MyBase.OnFontChanged(e)
			If Me.AutoSize Then AdjustSizeForOrientation()
		End Sub

		' Methods
		Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
			If AutoSize Then
				Using g As Graphics = Me.CreateGraphics()
					Select Case _orientation
						Case TextOrientation.Horizontal
							Return TextRenderer.MeasureText(Me.Text, Me.Font)
						Case TextOrientation.Stacked
							Dim lineHeight As Integer = TextRenderer.MeasureText("X", Me.Font).Height
							Return New Size(Me.Width, lineHeight * Me.Text.Length)
						Case TextOrientation.Vertical
							Dim sf As New StringFormat(StringFormat.GenericTypographic)
							Dim size As SizeF = g.MeasureString(Me.Text, Me.Font, Integer.MaxValue, sf)
							Return New Size(CInt(size.Height), CInt(size.Width))
					End Select
				End Using
			End If
			Return MyBase.GetPreferredSize(proposedSize)
		End Function
		Private Sub AdjustSizeForOrientation()
			Using g As Graphics = Me.CreateGraphics()
				Select Case _orientation
					Case TextOrientation.Horizontal
						Dim size As SizeF = TextRenderer.MeasureText(Text, Font)
						Width = CInt(size.Width)
						Height = CInt(size.Height)
					Case TextOrientation.Stacked
						Dim lineHeight As Single = TextRenderer.MeasureText("X", Font).Height
						Dim neededHeight As Integer = CInt(lineHeight * Text.Length)
						Height = neededHeight
					Case TextOrientation.Vertical
						Dim size As SizeF = TextRenderer.MeasureText(Text, Font)
						Width = CInt(size.Height) ' Rotated text swaps width/height
						Height = CInt(size.Width)
					Case Else
						' Normal orientation: let the designer/user control size
				End Select
			End Using
		End Sub
		Private Shared Function GetStringFormatFromContentAlignment(align As ContentAlignment) As StringFormat
			Dim sf As New StringFormat()

			Select Case align
				Case ContentAlignment.TopLeft
					sf.Alignment = StringAlignment.Far
					sf.LineAlignment = StringAlignment.Far
				Case ContentAlignment.TopCenter
					sf.Alignment = StringAlignment.Center
					sf.LineAlignment = StringAlignment.Far
				Case ContentAlignment.TopRight
					sf.Alignment = StringAlignment.Near
					sf.LineAlignment = StringAlignment.Far
				Case ContentAlignment.MiddleLeft
					sf.Alignment = StringAlignment.Far
					sf.LineAlignment = StringAlignment.Center
				Case ContentAlignment.MiddleCenter
					sf.Alignment = StringAlignment.Center
					sf.LineAlignment = StringAlignment.Center
				Case ContentAlignment.MiddleRight
					sf.Alignment = StringAlignment.Near
					sf.LineAlignment = StringAlignment.Center
				Case ContentAlignment.BottomLeft
					sf.Alignment = StringAlignment.Far
					sf.LineAlignment = StringAlignment.Near
				Case ContentAlignment.BottomCenter
					sf.Alignment = StringAlignment.Center
					sf.LineAlignment = StringAlignment.Near
				Case ContentAlignment.BottomRight
					sf.Alignment = StringAlignment.Near
					sf.LineAlignment = StringAlignment.Near
			End Select

			Return sf
		End Function

	End Class

	''' <summary>
	''' A TextBox that only allows numeric input, with options for decimal points and negative values. Also includes min/max value enforcement and a ValueCommitted event.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class NumericTextBox
		Inherits System.Windows.Forms.TextBox

		' Declarations
		Private _isNormalizing As Boolean = False
		<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Property Value As Decimal?
			Get
				Dim v As Decimal
				If Decimal.TryParse(Me.Text, v) Then Return v
				Return Nothing
			End Get
			Set(value As Decimal?)
				If value.HasValue Then
					Me.Text = value.Value.ToString()
				Else
					Me.Text = ""
				End If
			End Set
		End Property
		<Category("Behavior"), Description("Value Committed Event"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Event ValueCommitted(sender As Object, value As Decimal?)

		' Designer Properties
		<Category("Behavior"), Description("Allow Decimal Point."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property AllowDecimal As Boolean = False
		<Category("Behavior"), Description("Allow Negative Indicator."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property AllowNegative As Boolean = False
		<Category("Behavior"), Description("Minimum allowed value."), DefaultValue(GetType(Decimal), "-79228162514264337593543950335"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Minimum As Decimal = Decimal.MinValue
		<Category("Behavior"), Description("Maximum allowed value."), DefaultValue(GetType(Decimal), "79228162514264337593543950335"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Maximum As Decimal = Decimal.MaxValue
		<Category("Appearance"), Description("Show thousands separators."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property UseThousandsSeparator As Boolean = False
		<Category("Appearance"), Description("Format as currency."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property CurrencyMode As Boolean = False
		<Category("Appearance"), Description("Currency symbol to display."), DefaultValue("$"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property CurrencySymbol As String = "$"

		' Control Events
		Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
			MyBase.OnKeyPress(e)

			' Allow control keys
			If Char.IsControl(e.KeyChar) Then Return

			' Digits always allowed
			If Char.IsDigit(e.KeyChar) Then Return

			' Decimal point
			If AllowDecimal AndAlso e.KeyChar = "."c AndAlso Not Me.Text.Contains("."c) Then Return

			' Negative sign
			If AllowNegative AndAlso e.KeyChar = "-"c AndAlso Me.SelectionStart = 0 AndAlso Not Me.Text.Contains("-"c) Then Return

			' Otherwise block
			e.Handled = True
		End Sub
		Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

			' If user is editing formatted text, strip formatting first
			If Not _isNormalizing Then
				If Me.Text.Contains(","c) OrElse (CurrencyMode AndAlso Me.Text.StartsWith(CurrencySymbol)) Then

					Dim original = Me.Text
					Dim originalPos = Me.SelectionStart

					' Count how many formatting characters are before the caret
					Dim removedBefore As Integer = 0

					' Currency symbol at the start
					If CurrencyMode AndAlso original.StartsWith(CurrencySymbol, StringComparison.Ordinal) Then
						If originalPos > 0 Then
							removedBefore += CurrencySymbol.Length
						End If
					End If

					' Commas before the caret
					For i = 0 To Math.Min(originalPos - 1, original.Length - 1)
						If original(i) = ","c Then
							removedBefore += 1
						End If
					Next

					' Build raw text (no commas, no currency symbol)
					Dim raw = original.Replace(",", "")
					If CurrencyMode Then
						raw = raw.Replace(CurrencySymbol, "")
					End If

					' New caret position = old position - chars removed before it
					Dim newPos = Math.Max(0, originalPos - removedBefore)

					Me.Text = raw
					Me.SelectionStart = Math.Min(newPos, Me.Text.Length)
				End If
			End If

			If e.KeyCode = Keys.Enter Then
				NormalizeText()
				RaiseValueCommitted()
				e.SuppressKeyPress = True   ' stops the ding
				MyBase.OnKeyDown(e)         ' lets the key bubble up
				Return
			End If

			MyBase.OnKeyDown(e)
		End Sub
		Protected Overrides Sub OnTextChanged(e As EventArgs)
			MyBase.OnTextChanged(e)

			' Skip filtering during normalization
			If _isNormalizing Then Return

			' Skip filtering if text is already formatted
			If Me.Text.Contains(","c) OrElse (CurrencyMode AndAlso Me.Text.StartsWith(CurrencySymbol)) Then Return

			Dim original = Me.Text
			Dim filtered = FilterText(original)

			If filtered <> original Then
				Dim pos = Me.SelectionStart
				Me.Text = filtered
				Me.SelectionStart = Math.Min(pos, Me.Text.Length)
			End If
		End Sub
		Protected Overrides Sub OnLeave(e As EventArgs)
			MyBase.OnLeave(e)
			NormalizeText()
			RaiseValueCommitted()
		End Sub

		' Methods
		Private Sub RaiseValueCommitted()
			RaiseEvent ValueCommitted(Me, Me.Value)
		End Sub
		Private Function FilterText(input As String) As String
			If String.IsNullOrEmpty(input) Then Return String.Empty

			Dim sb As New System.Text.StringBuilder()
			Dim hasDecimal As Boolean = False
			Dim hasNegative As Boolean = False

			For i = 0 To input.Length - 1
				Dim ch = input(i)

				If Char.IsDigit(ch) Then
					sb.Append(ch)
					Continue For
				End If

				If AllowDecimal AndAlso ch = "."c AndAlso Not hasDecimal Then
					sb.Append("."c)
					hasDecimal = True
					Continue For
				End If

				If AllowNegative AndAlso ch = "-"c AndAlso i = 0 AndAlso Not hasNegative Then
					sb.Append("-"c)
					hasNegative = True
					Continue For
				End If
			Next

			Return sb.ToString()
		End Function
		Private Sub NormalizeText()
			_isNormalizing = True
			Try
				Dim t = Me.Text.Trim()

				' If text ends with a decimal point, append a zero
				If t.EndsWith("."c) Then
					t &= "0"
				End If

				If t = "" Then
					Me.Text = ""
					Return
				End If

				' Lone decimal → "0."
				If t = "." Then
					Me.Text = "0."
					Return
				End If

				' Lone negative → "-0"
				If t = "-" Then
					Me.Text = "-0"
					Return
				End If

				' "-." → "-0."
				If t = "-." Then
					Me.Text = "-0."
					Return
				End If

				' DECIMAL MODE
				If AllowDecimal Then
					Dim value As Decimal
					If Decimal.TryParse(t, value) Then

						' Clamp
						If value < Minimum Then value = Minimum
						If value > Maximum Then value = Maximum

						' Apply formatting
						Me.Text = FormatValue(value)
					End If

					Return
				End If

				' INTEGER MODE
				Dim intValue As Integer
				If Integer.TryParse(t, intValue) Then

					' Clamp
					If intValue < Minimum Then intValue = CInt(Minimum)
					If intValue > Maximum Then intValue = CInt(Maximum)

					' Apply formatting
					Me.Text = FormatValue(intValue)
				End If
			Catch
			Finally
				_isNormalizing = False
			End Try
		End Sub
		Private Function FormatValue(value As Decimal) As String
			If CurrencyMode Then
				' Currency formatting: $1,234.56
				Return CurrencySymbol & value.ToString("N2")
			End If

			If UseThousandsSeparator Then
				' Thousands only: 1,234.56 or 1,234
				If AllowDecimal Then
					Return value.ToString("N2")
				Else
					Return value.ToString("N0")
				End If
			End If

			' No formatting
			Return value.ToString()
		End Function

	End Class

	''' <summary>
	''' Improved And Extended RichTextBox control that prevents the cursor from changing to the I-beam when hovering over the control. It can cause cursor "blinking". This is especially useful in scenarios where the RichTextBox is used for display purposes only and should not allow text selection or editing.
	''' Also includes a SetAlignment method to easily set text alignment. Now includes methods to append and insert both RTF and plain text at specified positions.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class RichTextBox
		Inherits System.Windows.Forms.RichTextBox

		Protected Overrides Sub WndProc(ByRef m As Message)
			Const WM_SETCURSOR As Integer = &H20

			If m.Msg = WM_SETCURSOR Then
				Cursor.Current = Cursors.Default
				m.Result = IntPtr.Zero
				Return
			End If

			MyBase.WndProc(m)
		End Sub

		''' <summary>
		''' Sets the text alignment of the RichTextBox.
		''' </summary>
		''' <param name="align">The alignment to set.</param>
		Public Sub SetAlignment(align As HorizontalAlignment)
			Me.SelectAll()
			Me.SelectionAlignment = align
			Me.DeselectAll()
		End Sub
		''' <summary>
		''' Appends RTF formatted text to the end of the RichTextBox.
		''' </summary>
		''' <param name="rtf">The RTF formatted text to append.</param>
		Public Sub AppendRTF(rtf As String)
			If String.IsNullOrEmpty(rtf) Then Exit Sub
			Me.Select(Me.TextLength, 0)
			Me.SelectedRtf = rtf
			Me.Select(Me.TextLength, 0)
			Me.ScrollToCaret()
		End Sub
		''' <summary>
		''' Inserts RTF formatted text at the specified index in the RichTextBox.
		''' </summary>
		''' <param name="index">The zero-based index at which to insert the text.</param>
		''' <param name="rtf">The RTF formatted text to insert.</param>
		Public Sub InsertRTFAt(index As Integer, rtf As String)
			If String.IsNullOrEmpty(rtf) Then Exit Sub
			If index < 0 OrElse index > Me.TextLength Then Exit Sub
			Me.Select(index, 0)
			Me.SelectedRtf = rtf
			Me.Select(index, 0)
			Me.ScrollToCaret()
		End Sub
		''' <summary>
		''' Appends plain text to the end of the RichTextBox.
		''' </summary>
		''' <param name="text">The plain text to append.</param>
		Public Sub AppendPlainText(text As String)
			If text Is Nothing Then Exit Sub
			Me.Select(Me.TextLength, 0)
			Me.SelectedText = text
			Me.Select(Me.TextLength, 0)
			Me.ScrollToCaret()
		End Sub
		''' <summary>
		''' Inserts plain text at the specified index in the RichTextBox.
		''' </summary>
		''' <param name="index">The zero-based index at which to insert the text.</param>
		''' <param name="text">The plain text to insert.</param>
		Public Sub InsertPlainTextAt(index As Integer, text As String)
			If text Is Nothing Then Exit Sub
			If index < 0 OrElse index > Me.TextLength Then Exit Sub
			Me.Select(index, 0)
			Me.SelectedText = text
			Me.Select(index, 0)
			Me.ScrollToCaret()
		End Sub

	End Class

	''' <summary>
	''' Simple, Old-Style Context Menu for TextBoxes.
	''' Contains everything needed for basic Cut &amp; Paste functionality from ContextMenu &amp; Keyboard.
	''' 1.	Must create New Instance of TextBoxContextMenu either in Designer or Programmatically.
	''' 2.	Set the ContextMenu Property of the TextBox to the Instance of TextBoxContextMenu.
	''' 3.	Handle PreviewKeyDownEvent on TextBox &amp; call ShortcutKeys Function.
	''' 4.	Set ShortcutsEnabled property to False.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class TextBoxContextMenu
		Inherits System.Windows.Forms.ContextMenuStrip

		'Declarations
		Private WithEvents MIUndo As ToolStripMenuItem
		Private WithEvents MICut As ToolStripMenuItem
		Private WithEvents MICopy As ToolStripMenuItem
		Private WithEvents MIPaste As ToolStripMenuItem
		Private WithEvents MIDelete As ToolStripMenuItem
		Private ReadOnly MISeparatorProperCase As New ToolStripSeparator
		Private WithEvents MIProperCase As ToolStripMenuItem
		Private WithEvents MISelectAll As ToolStripMenuItem

		'Properties
		<DefaultValue(False)>
		Public Property ShowExtendedTools As Boolean

		'Events
		Public Sub New()
			MyBase.New
			MIUndo = New ToolStripMenuItem("Undo", Resources.ImageEditUndo16, AddressOf UndoClick)
			Me.Items.Add(MIUndo)
			Me.Items.Add(New ToolStripSeparator())
			MICut = New ToolStripMenuItem("Cut", Resources.ImageEditCut16, AddressOf CutClick)
			Me.Items.Add(MICut)
			MICopy = New ToolStripMenuItem("Copy", Resources.ImageEditCopy16, AddressOf CopyClick)
			Me.Items.Add(MICopy)
			MIPaste = New ToolStripMenuItem("Paste", Resources.ImageEditPaste16, AddressOf PasteClick)
			Me.Items.Add(MIPaste)
			MIDelete = New ToolStripMenuItem("Delete", Resources.ImageEditDelete16, AddressOf DeleteClick)
			Me.Items.Add(MIDelete)
			Me.Items.Add(MISeparatorProperCase)
			MIProperCase = New ToolStripMenuItem("Proper Case", Resources.ImageEditProperCase16, AddressOf ProperCaseClick)
			Me.Items.Add(MIProperCase)
			Me.Items.Add(New ToolStripSeparator())
			MISelectAll = New ToolStripMenuItem("Select All", Resources.ImageEditSelectAll16, AddressOf SelectAllClick)
			Me.Items.Add(MISelectAll)
		End Sub
		Protected Overrides Sub OnHandleCreated(e As EventArgs)
			MyBase.OnHandleCreated(e)

			Try
				Encoding.GetEncoding("windows-1252")
			Catch ex As ArgumentException
				Debug.WriteLine("SkyeControl Warning: Encoding 'windows-1252' not available. Call System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance) in your app initializer, and add the System.Text.Encoding.CodePages Nuget package to your project to enable legacy text parsing.")
			End Try

		End Sub
		Protected Overrides Sub OnOpening(e As CancelEventArgs)
			MyBase.OnOpening(e)

			Dim txbx As TextBox = TryCast(SourceControl, TextBox)
			If txbx Is Nothing Then Return
			MIUndo.Enabled = txbx.CanUndo AndAlso Not txbx.ReadOnly
			MICut.Enabled = txbx.SelectedText.Length > 0 AndAlso Not txbx.ReadOnly
			MICopy.Enabled = txbx.SelectedText.Length > 0
			MIPaste.Enabled = Clipboard.ContainsText() AndAlso Not txbx.ReadOnly
			MIDelete.Enabled = txbx.SelectedText.Length > 0 AndAlso Not txbx.ReadOnly
			MISeparatorProperCase.Visible = ShowExtendedTools
			MIProperCase.Visible = ShowExtendedTools
			MISelectAll.Enabled = txbx.Text.Length > 0 AndAlso txbx.SelectedText.Length < txbx.Text.Length

			If txbx.SelectedText.Length > 0 Then txbx.Focus()

		End Sub
		<System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static",
										Justification:="Kept as instance for API consistency")>
		Public Sub ShortcutKeys(ByRef sender As TextBox, e As PreviewKeyDownEventArgs)
			If e.Control Then
				Select Case e.KeyCode
					Case Keys.A : SelectAll(sender)
					Case Keys.C : Copy(sender)
					Case Keys.D : Delete(sender)
					Case Keys.V : Paste(sender)
					Case Keys.X : Cut(sender)
					Case Keys.Z : Undo(sender)
				End Select
			End If
		End Sub
		Private Sub UndoClick(ByVal sender As Object, ByVal e As EventArgs)
			Undo(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub CutClick(ByVal sender As Object, ByVal e As EventArgs)
			Cut(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub CopyClick(ByVal sender As Object, ByVal e As EventArgs)
			Copy(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub PasteClick(ByVal sender As Object, ByVal e As EventArgs)
			Paste(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub DeleteClick(ByVal sender As Object, ByVal e As EventArgs)
			Delete(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub ProperCaseClick(ByVal sender As Object, ByVal e As EventArgs)
			ProperCase(TryCast(SourceControl, TextBox))
		End Sub
		Private Sub SelectAllClick(ByVal sender As Object, ByVal e As EventArgs)
			SelectAll(TryCast(SourceControl, TextBox))
		End Sub

		'Procedures
		Private Shared Sub Undo(txbx As TextBox)
			txbx.Undo()
			If txbx.FindForm IsNot Nothing Then txbx.FindForm.Validate()
		End Sub
		Private Shared Sub Cut(txbx As TextBox)
			txbx.Cut()
			If txbx.FindForm IsNot Nothing Then txbx.FindForm.Validate()
		End Sub
		Private Shared Sub Copy(txbx As TextBox)
			txbx.Copy()
		End Sub
		Private Shared Sub Paste(txbx As TextBox)
			txbx.Paste()
			If txbx.FindForm IsNot Nothing Then txbx.FindForm.Validate()
		End Sub
		Private Shared Sub Delete(txbx As TextBox)
			If Not txbx.ReadOnly Then
				txbx.SelectedText = String.Empty
				txbx.FindForm()?.Validate()
			End If
		End Sub
		Private Shared Sub ProperCase(txbx As TextBox)
			txbx.Focus()
			txbx.Text = StrConv(txbx.Text, VbStrConv.ProperCase)
			If txbx.FindForm IsNot Nothing Then txbx.FindForm.Validate()
		End Sub
		Private Shared Sub SelectAll(txbx As TextBox)
			txbx.SelectAll()
			txbx.Focus()
		End Sub

	End Class

	''' <summary>
	''' Simple, Old-Style Context Menu for RichTextBoxes.
	''' Contains everything needed for basic Cut &amp; Paste functionality from ContextMenu &amp; Keyboard.
	''' 1.	Must create New Instance of RichTextBoxContextMenu either in Designer or Programmatically.
	''' 2.	Set the ContextMenu Property of the RichTextBox to the Instance of RichTextBoxContextMenu.
	''' 3.	Handle PreviewKeyDownEvent on RichTextBox &amp; call ShortcutKeys Function. If Using Skye.UI.RichTextBox, you can still call the ShortcutKeys function, just perform a CType on the sender to System.Windows.Forms.RichTextBox. This is because Skye.UI.RichTextBox inherits from System.Windows.Forms.RichTextBox and the ShortcutKeys function is designed to work with that type.
	''' 4.	Set ShortcutsEnabled property of the RichTextBox to False.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class RichTextBoxContextMenu
		Inherits System.Windows.Forms.ContextMenuStrip

		'Declarations
		Private WithEvents MIUndo As ToolStripMenuItem
		Private WithEvents MICut As ToolStripMenuItem
		Private WithEvents MICopy As ToolStripMenuItem
		Private WithEvents MIPaste As ToolStripMenuItem
		Private WithEvents MIDelete As ToolStripMenuItem
		Private WithEvents MISelectAll As ToolStripMenuItem

		'Events
		Public Sub New()
			MyBase.New
			MIUndo = New ToolStripMenuItem("Undo", Resources.ImageEditUndo16, AddressOf UndoClick)
			Me.Items.Add(MIUndo)
			Me.Items.Add(New ToolStripSeparator())
			MICut = New ToolStripMenuItem("Cut", Resources.ImageEditCut16, AddressOf CutClick)
			Me.Items.Add(MICut)
			MICopy = New ToolStripMenuItem("Copy", Resources.ImageEditCopy16, AddressOf CopyClick)
			Me.Items.Add(MICopy)
			MIPaste = New ToolStripMenuItem("Paste", Resources.ImageEditPaste16, AddressOf PasteClick)
			Me.Items.Add(MIPaste)
			MIDelete = New ToolStripMenuItem("Delete", Resources.ImageEditDelete16, AddressOf DeleteClick)
			Me.Items.Add(MIDelete)
			Me.Items.Add(New ToolStripSeparator())
			MISelectAll = New ToolStripMenuItem("Select All", Resources.ImageEditSelectAll16, AddressOf SelectAllClick)
			Me.Items.Add(MISelectAll)
		End Sub
		Protected Overrides Sub OnOpening(e As CancelEventArgs)
			MyBase.OnOpening(e)
			Dim rtb As RichTextBox = TryCast(SourceControl, RichTextBox)
			If rtb Is Nothing Then Return
			MIUndo.Enabled = rtb.CanUndo AndAlso Not rtb.ReadOnly
			MICut.Enabled = rtb.SelectedText.Length > 0 AndAlso Not rtb.ReadOnly
			MICopy.Enabled = rtb.SelectedText.Length > 0
			MIPaste.Enabled = Clipboard.ContainsText() AndAlso Not rtb.ReadOnly
			MIDelete.Enabled = rtb.SelectedText.Length > 0 AndAlso Not rtb.ReadOnly
			MISelectAll.Enabled = rtb.Text.Length > 0 AndAlso rtb.SelectedText.Length < rtb.Text.Length
			If rtb.SelectedText.Length > 0 Then rtb.Focus()
		End Sub
		<System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static",
										Justification:="Kept as instance for API consistency")>
		Public Sub ShortcutKeys(ByRef sender As System.Windows.Forms.RichTextBox, e As PreviewKeyDownEventArgs)
			If e.Control Then
				Select Case e.KeyCode
					Case Keys.A : SelectAll(sender)
					Case Keys.C : Copy(sender)
					Case Keys.D : Delete(sender)
					Case Keys.V : Paste(sender)
					Case Keys.X : Cut(sender)
					Case Keys.Z : Undo(sender)
				End Select
			End If
		End Sub
		Private Sub UndoClick(ByVal sender As Object, ByVal e As EventArgs)
			Undo(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub
		Private Sub CutClick(ByVal sender As Object, ByVal e As EventArgs)
			Cut(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub
		Private Sub CopyClick(ByVal sender As Object, ByVal e As EventArgs)
			Copy(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub
		Private Sub PasteClick(ByVal sender As Object, ByVal e As EventArgs)
			Paste(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub
		Private Sub DeleteClick(ByVal sender As Object, ByVal e As EventArgs)
			Delete(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub
		Private Sub SelectAllClick(ByVal sender As Object, ByVal e As EventArgs)
			SelectAll(TryCast(SourceControl, System.Windows.Forms.RichTextBox))
		End Sub

		'Procedures
		Private Shared Sub Undo(rtb As System.Windows.Forms.RichTextBox)
			rtb.Undo()
			If rtb.FindForm IsNot Nothing Then rtb.FindForm.Validate()
		End Sub
		Private Shared Sub Cut(rtb As System.Windows.Forms.RichTextBox)
			rtb.Cut()
			If rtb.FindForm IsNot Nothing Then rtb.FindForm.Validate()
		End Sub
		Private Shared Sub Copy(rtb As System.Windows.Forms.RichTextBox)
			rtb.Copy()
		End Sub
		Private Shared Sub Paste(rtb As System.Windows.Forms.RichTextBox)
			rtb.Paste()
			If rtb.FindForm IsNot Nothing Then rtb.FindForm.Validate()
		End Sub
		Private Shared Sub Delete(rtb As System.Windows.Forms.RichTextBox)
			If Not rtb.ReadOnly Then
				rtb.SelectedText = String.Empty
				rtb.FindForm()?.Validate()
			End If
		End Sub
		Private Shared Sub SelectAll(rtb As System.Windows.Forms.RichTextBox)
			rtb.SelectAll()
			rtb.Focus()
		End Sub

	End Class

End Namespace
