
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Text
Imports Skye.UI.ToolTipEX

Namespace UI



	''' <summary>
	''' A ComboBox that supports custom items with text, images, and color swatches. Fixes the stock combobox dropdownlist draw background issues, making it theme-ready. Also includes hover effects for improved user experience.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ComboBox
		Inherits System.Windows.Forms.ComboBox

		' Declarations
		Private _hovering As Boolean = False
		Private _editBox As TextBox = Nothing

		' Designer Properties
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Shadows Property DrawMode As DrawMode
			Get
				Return MyBase.DrawMode
			End Get
			Set(value As DrawMode)
				' Lock it to OwnerDrawFixed
				MyBase.DrawMode = DrawMode.OwnerDrawFixed
			End Set
		End Property

		' Events
		Protected Overrides Sub WndProc(ByRef m As Message)

			' ---------------------------------------------------------
			' 0. Suppress native repaint on focus changes
			' ---------------------------------------------------------
			If m.Msg = WinAPI.WM_SETFOCUS OrElse m.Msg = WinAPI.WM_KILLFOCUS Then
				WinAPI.SendMessage(Me.Handle, WinAPI.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)
				MyBase.WndProc(m)
				WinAPI.SendMessage(Me.Handle, WinAPI.WM_SETREDRAW, CType(1, IntPtr), IntPtr.Zero)
				Me.Invalidate()
				Return
			End If

			' ---------------------------------------------------------
			' 1. Full custom painting for DropDownList
			' ---------------------------------------------------------
			If Me.DropDownStyle = ComboBoxStyle.DropDownList AndAlso m.Msg = WinAPI.WM_PAINT Then
				Dim ps As New WinAPI.PAINTSTRUCT
				Dim hdc As IntPtr = WinAPI.BeginPaint(Me.Handle, ps)
				Try
					Using g As Graphics = Graphics.FromHdc(hdc)
						PaintDDLCombo(g)
					End Using
				Finally
					WinAPI.EndPaint(Me.Handle, ps)
				End Try
				Return
			End If

			' ---------------------------------------------------------
			' 2. Let Windows paint normally (textbox, caret, selection)
			' ---------------------------------------------------------
			MyBase.WndProc(m)

			' ---------------------------------------------------------
			' 3. Post‑paint: unify visuals for DropDown mode
			' ---------------------------------------------------------
			If m.Msg = WinAPI.WM_PAINT OrElse m.Msg = WinAPI.WM_NCPAINT Then
				Using g As Graphics = Graphics.FromHwnd(Me.Handle)

					If Me.DropDownStyle = ComboBoxStyle.DropDown Then

						' Paint textbox background OVER the native white border
						PaintDDTextBoxBackground(g)

						' Paint arrow background
						PaintDDArrowBackground(g)

						' Paint arrow
						DrawUnifiedArrow(g, Me.ClientRectangle)

						' Paint border last
						PaintDDBorder(g)

					End If

				End Using
			End If

		End Sub
		Public Sub New()
			MyBase.New()
			DrawMode = DrawMode.OwnerDrawFixed
			DoubleBuffered = True
			SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
			UpdateStyles()
		End Sub
		Protected Overrides Sub OnHandleCreated(e As EventArgs)
			MyBase.OnHandleCreated(e)

			If Not DesignMode Then
				Me.DrawMode = DrawMode.OwnerDrawFixed
			End If

			If Me.DropDownStyle = ComboBoxStyle.DropDown Then
				For Each c As Control In Me.Controls
					If TypeOf c Is TextBox Then
						_editBox = DirectCast(c, TextBox)
						_editBox.BorderStyle = BorderStyle.None
						_editBox.BackColor = Me.BackColor
						_editBox.ForeColor = Me.ForeColor
						Exit For
					End If
				Next
			End If

		End Sub
		Protected Overrides Sub OnCreateControl()
			MyBase.OnCreateControl()

			If Not Me.DesignMode Then RecalculateDropdownHeight()

		End Sub
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			MyBase.OnPaint(e)
			PaintDDLCombo(e.Graphics)
		End Sub
		Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
			If e.Index < 0 Then Return

			If Not Me.Enabled Then
				e.Graphics.FillRectangle(New SolidBrush(DisabledColor(Me.BackColor)), e.Bounds)
				Using b As New SolidBrush(DisabledColor(Me.ForeColor))
					e.Graphics.DrawString(Me.Items(e.Index).ToString(), Me.Font, b, e.Bounds)
				End Using
				Return
			End If

			Dim g = e.Graphics

			' --- ClearType + quality flags ---
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
			'g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.PixelOffsetMode = Drawing2D.PixelOffsetMode.None
			g.CompositingQuality = Drawing2D.CompositingQuality.AssumeLinear

			Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected

			' --- Background color (theme aware) ---
			Dim bg As Color = If(isSelected, SystemColors.Highlight, Me.BackColor)
			Dim textColor As Color = If(isSelected, SystemColors.HighlightText, Me.ForeColor)

			e.DrawBackground()

			' --- Now draw your theme background on top ---
			Using b As New SolidBrush(bg)
				g.FillRectangle(b, e.Bounds)
			End Using

			' --- Try to get ComboItem ---
			Dim item = TryCast(Me.Items(e.Index), ComboItem)

			' --- Fallback branch (non-ComboItem) ---
			If item Is Nothing Then
				'e.DrawBackground()
				Using b As New SolidBrush(bg)
					g.FillRectangle(b, e.Bounds)
				End Using

				' Text rectangle
				Dim textRect As New Rectangle(
					e.Bounds.Left + 4,
					e.Bounds.Top,
					e.Bounds.Width - 4,
					e.Bounds.Height)

				Dim sf As New StringFormat() With {
					.LineAlignment = StringAlignment.Center,
					.Alignment = StringAlignment.Near,
					.Trimming = StringTrimming.EllipsisCharacter,
					.FormatFlags = StringFormatFlags.NoWrap
				}

				Using b As New SolidBrush(textColor)
					g.DrawString(Me.Items(e.Index).ToString(), Me.Font, b, textRect, sf)
				End Using

				e.DrawFocusRectangle()
				Return
			End If

			' --- ComboItem branch ---
			Dim iconSize As Integer = e.Bounds.Height - 4
			Dim iconX As Integer = e.Bounds.Left + 4
			Dim iconY As Integer = e.Bounds.Top + 2
			Dim textX As Integer = iconX

			' Image
			If item.Image IsNot Nothing Then
				g.DrawImage(item.Image, New Rectangle(iconX, iconY, iconSize, iconSize))
				textX += iconSize + 6
			ElseIf item.Swatch.HasValue Then ' Swatch
				Dim swatchRect As New Rectangle(iconX, iconY, iconSize, iconSize)
				Using b As New SolidBrush(item.Swatch.Value)
					g.FillRectangle(b, swatchRect)
				End Using
				Using p As New Pen(Color.Black)
					g.DrawRectangle(p, swatchRect)
				End Using
				textX += iconSize + 6
			End If

			' --- Text ---
			Dim textRect2 As New Rectangle(textX, e.Bounds.Top,
								   e.Bounds.Width - textX, e.Bounds.Height)

			Dim sf2 As New StringFormat() With {
				.LineAlignment = StringAlignment.Center,
				.Alignment = StringAlignment.Near,
				.Trimming = StringTrimming.EllipsisCharacter,
				.FormatFlags = StringFormatFlags.NoWrap
			}

			Using b As New SolidBrush(textColor)
				g.DrawString(item.Text, Me.Font, b, textRect2, sf2)
			End Using

			e.DrawFocusRectangle()
		End Sub
		Protected Overrides Sub OnMouseEnter(e As EventArgs)
			MyBase.OnMouseEnter(e)
			_hovering = True
			Invalidate()
		End Sub
		Protected Overrides Sub OnMouseLeave(e As EventArgs)
			MyBase.OnMouseLeave(e)
			_hovering = False
			Invalidate()
		End Sub
		Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
			MyBase.OnMouseMove(e)
			Dim overNow As Boolean = Me.ClientRectangle.Contains(e.Location)
			If overNow <> _hovering Then
				_hovering = overNow
				Me.Invalidate()
			End If
		End Sub
		Protected Overrides Sub OnDataSourceChanged(e As EventArgs)
			MyBase.OnDataSourceChanged(e)
			RecalculateDropdownHeight()
		End Sub
		Protected Overrides Sub OnSelectedIndexChanged(e As EventArgs)
			MyBase.OnSelectedIndexChanged(e)
			RecalculateDropdownHeight()
		End Sub
		Protected Overrides Sub OnDropDown(e As EventArgs)
			MyBase.OnDropDown(e)
			RecalculateDropdownHeight()
		End Sub
		Protected Overrides Sub OnDropDownClosed(e As EventArgs)
			MyBase.OnDropDownClosed(e)
			_hovering = False
			Invalidate()
		End Sub

		' Methods
		Private Sub PaintDDLCombo(g As Graphics)
			g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

			Dim rc As Rectangle = Me.ClientRectangle
			Dim bgColor As Color = Me.BackColor
			Dim textColor As Color = Me.ForeColor
			Dim arrowColor As Color = Me.ForeColor
			If Not Me.Enabled Then
				bgColor = DisabledColor(bgColor)
				textColor = DisabledColor(textColor)
				arrowColor = DisabledColor(arrowColor)
			End If

			' ---------------------------------------------------------
			' 1. Background
			' ---------------------------------------------------------
			If DropDownStyle = ComboBoxStyle.DropDownList Then
				' You own the whole background
				Dim bg As Color = If(_hovering, ControlPaint.Light(bgColor, 0.25F), bgColor)
				Using b As New SolidBrush(bg)
					g.FillRectangle(b, rc)
				End Using
			Else
				' DropDown mode → DO NOT paint the text area.
				' Only paint the dropdown button area.
				Dim buttonWidth As Integer = SystemInformation.VerticalScrollBarWidth
				Dim buttonRect As New Rectangle(rc.Right - buttonWidth, rc.Top, buttonWidth, rc.Height)

				' IMPORTANT: Only fill the button rectangle, NOT rc.
				Using b As New SolidBrush(bgColor)
					g.FillRectangle(b, buttonRect)
				End Using
			End If

			' ---------------------------------------------------------
			' 2. Draw the selected item text (DropDownList ONLY)
			' ---------------------------------------------------------
			If DropDownStyle = ComboBoxStyle.DropDownList Then
				Dim text As String = ""
				Dim textX As Integer = 4
				Dim selectedItem = TryCast(Me.SelectedItem, ComboItem)
				If selectedItem IsNot Nothing Then
					text = selectedItem.Text
					Dim iconSize As Integer = rc.Height - 6
					Dim iconX As Integer = 4
					Dim iconY As Integer = (rc.Height - iconSize) \ 2
					If selectedItem.Image IsNot Nothing Then
						g.DrawImage(selectedItem.Image, New Rectangle(iconX, iconY, iconSize, iconSize))
						textX = iconX + iconSize + 6
					ElseIf selectedItem.Swatch.HasValue Then
						Dim swatchRect As New Rectangle(iconX, iconY, iconSize, iconSize)
						Using b As New SolidBrush(selectedItem.Swatch.Value)
							g.FillRectangle(b, swatchRect)
						End Using
						Using p As New Pen(Color.Black)
							g.DrawRectangle(p, swatchRect)
						End Using
						textX = iconX + iconSize + 6
					End If
				ElseIf Me.SelectedItem IsNot Nothing Then
					text = Me.SelectedItem.ToString()
				End If
				Dim arrowWidth As Integer = SystemInformation.VerticalScrollBarWidth
				Dim textRect As New Rectangle(textX, rc.Top, rc.Width - textX - arrowWidth - 4, rc.Height)
				Dim sf As New StringFormat With {
					.FormatFlags = StringFormatFlags.NoWrap,
					.Trimming = StringTrimming.EllipsisCharacter,
					.LineAlignment = StringAlignment.Center,
					.Alignment = StringAlignment.Near
				}

				Using b As New SolidBrush(textColor)
					g.DrawString(text, Me.Font, b, textRect, sf)
				End Using
			End If

			' ---------------------------------------------------------
			' 3. Draw the arrow (BOTH modes)
			' ---------------------------------------------------------
			DrawUnifiedArrow(g, rc)

			' ---------------------------------------------------------
			' 4. Border
			' ---------------------------------------------------------
			Dim borderColor As Color = If(Me.Enabled, ControlPaint.Dark(Me.BackColor), DisabledColor(ControlPaint.Dark(Me.BackColor)))
			Using p As New Pen(borderColor)
				g.DrawRectangle(p, 0, 0, rc.Width - 1, rc.Height - 1)
			End Using

		End Sub
		Private Sub PaintDDBorder(g As Graphics)
			Dim rc As Rectangle = Me.ClientRectangle
			Dim borderColor As Color = If(Me.Enabled, ControlPaint.Dark(Me.BackColor), DisabledColor(ControlPaint.Dark(Me.BackColor)))
			Using p As New Pen(borderColor)
				g.DrawRectangle(p, 0, 0, rc.Width - 1, rc.Height - 1)
			End Using
		End Sub
		Private Sub PaintDDTextBoxBackground(g As Graphics)
			If Me.DropDownStyle <> ComboBoxStyle.DropDown Then Exit Sub

			Const border As Integer = 1
			Dim textRect As New Rectangle(border, border, Me.Width - 18 - (border * 2), Me.Height - (border * 2))
			Dim fill As Color = If(Me.Enabled, Me.BackColor, DisabledColor(Me.BackColor))
			Using b As New SolidBrush(fill)
				g.FillRectangle(b, textRect)
			End Using

		End Sub
		Private Sub PaintDDArrowBackground(g As Graphics)
			Const border As Integer = 1
			Dim arrowRect As New Rectangle(Width - 18 - border, border, 18, Height - (border * 2))
			Dim bg As Color = If(_hovering, ControlPaint.Light(Me.BackColor, 0.25F), Me.BackColor)

			Using b As New SolidBrush(bg)
				g.FillRectangle(b, arrowRect)
			End Using

		End Sub
		Private Sub DrawUnifiedArrow(g As Graphics, rc As Rectangle)
			Dim arrowSize As Integer = 6
			Dim buttonW2 As Integer = SystemInformation.VerticalScrollBarWidth
			Dim arrowX As Integer = rc.Right - buttonW2 \ 2 - arrowSize \ 2 - 2
			Dim arrowY As Integer = (rc.Height - arrowSize) \ 2 + 1
			Dim arrowColor As Color
			If Not Me.Enabled Then
				arrowColor = DisabledColor(Me.ForeColor)
			Else
				arrowColor = If(_hovering, ControlPaint.Dark(Me.ForeColor, 0.1F), Me.ForeColor)
			End If
			Dim pts As Point() = {
				New Point(arrowX, arrowY),
				New Point(arrowX + arrowSize, arrowY),
				New Point(arrowX + arrowSize \ 2, arrowY + arrowSize)
			}

			Using b As New SolidBrush(arrowColor)
				g.FillPolygon(b, pts)
			End Using

		End Sub
		Private Sub RecalculateDropdownHeight()
			Dim maxLines As Integer = MaxDropDownItems
			If maxLines <= 0 Then maxLines = 8
			Dim itemHeight As Integer = Me.ItemHeight
			DropDownHeight = itemHeight * maxLines + 4
		End Sub
		Private Function DisabledColor(c As Color) As Color
			' Blend with white at 50%
			Return Color.FromArgb(
				c.A,
				(c.R + 255) \ 2,
				(c.G + 255) \ 2,
				(c.B + 255) \ 2)
		End Function

		Public Class ComboItem

			Public Property Text As String
			Public Property Image As Image
			Public Property Swatch As Color?

			Public Sub New(text As String,
						   Optional img As Image = Nothing,
						   Optional swatch As Color? = Nothing)
				Me.Text = text
				Me.Image = img
				Me.Swatch = swatch
			End Sub
			Public Overrides Function ToString() As String
				Return Text
			End Function

		End Class

	End Class

	''' <summary>
	''' Color Selector ComboBox, Using System.Drawing.Color Structure
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ColorComboBox
		Inherits ComboBox

		' Declarations
		Private _hovering As Boolean = False

		''' <summary>
		''' Gets/sets the selected color of ComboBox
		''' (Default color is Black)
		''' </summary>
		<System.ComponentModel.Category("Data")>
		<System.ComponentModel.ReadOnly(True)>
		<DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Color As Color
			Get
				If SelectedItem IsNot Nothing Then
					Return CType(SelectedItem, Color)
				End If
				Return Color.Black 'Default Setting
			End Get
			Set
				Dim ix As Integer = Items.IndexOf(Value)
				If ix >= 0 Then
					SelectedIndex = ix
				End If
			End Set
		End Property

		' Events
		Protected Overrides Sub WndProc(ByRef m As Message)

			' 0. Suppress native repaint on focus changes
			If m.Msg = WinAPI.WM_SETFOCUS OrElse m.Msg = WinAPI.WM_KILLFOCUS Then
				WinAPI.SendMessage(Me.Handle, WinAPI.WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)
				MyBase.WndProc(m)
				WinAPI.SendMessage(Me.Handle, WinAPI.WM_SETREDRAW, CType(1, IntPtr), IntPtr.Zero)
				Me.Invalidate()
				Return
			End If

			' 1. Full custom painting for DropDownList
			If Me.DropDownStyle = ComboBoxStyle.DropDownList AndAlso m.Msg = WinAPI.WM_PAINT Then
				Dim ps As New WinAPI.PAINTSTRUCT
				Dim hdc As IntPtr = WinAPI.BeginPaint(Me.Handle, ps)
				Try
					Using g As Graphics = Graphics.FromHdc(hdc)
						PaintDDLCombo(g)
					End Using
				Finally
					WinAPI.EndPaint(Me.Handle, ps)
				End Try
				Return
			End If

			' 2. Let Windows paint normally (textbox mode)
			MyBase.WndProc(m)

		End Sub
		Public Sub New()
			'Change DrawMode for custom drawing
			DrawMode = DrawMode.OwnerDrawFixed
			DropDownStyle = ComboBoxStyle.DropDownList
			'Add System.Drawing.Color Structure To Item List
			FillColors() 'May cause duplicates if used here for some reason, set manually by calling FillColors for each combobox in form constructor.
		End Sub
		Protected Overrides Sub OnMouseEnter(e As EventArgs)
			_hovering = True
			Invalidate()
		End Sub
		Protected Overrides Sub OnMouseLeave(e As EventArgs)
			_hovering = False
			Invalidate()
		End Sub
		Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
			Dim overNow As Boolean = Me.ClientRectangle.Contains(e.Location)
			If overNow <> _hovering Then
				_hovering = overNow
				Invalidate()
			End If
		End Sub
		Protected Overrides Sub OnDropDownClosed(e As EventArgs)
			MyBase.OnDropDownClosed(e)
			_hovering = False
			Invalidate()
		End Sub
		Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
			If e.Index < 0 Then Return
			Dim g = e.Graphics

			' ClearType + quality flags
			g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
			g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
			g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

			' Background
			Dim isSelected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected
			Dim bg As Color = If(isSelected, SystemColors.Highlight, Me.BackColor)
			Using b As New SolidBrush(bg)
				g.FillRectangle(b, e.Bounds)
			End Using

			' Draw color swatch 
			Dim color As Color = CType(Items(e.Index), Color)
			Dim swatchSize As Integer = e.Bounds.Height - 6
			Dim swatchRect As New Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 3, swatchSize, swatchSize)
			Dim radius As Integer = 4 ' adjust for softer or sharper corners
			Dim swatchPath = CreateRoundedRect(swatchRect, radius)
			Using b As New SolidBrush(color)
				g.FillPath(b, swatchPath)
			End Using
			Using p As New Pen(Color.FromArgb(180, 0, 0, 0)) ' softer black
				g.DrawPath(p, swatchPath)
			End Using

			' Draw text
			Dim textColor As Color = If(isSelected, SystemColors.HighlightText, Me.ForeColor)
			Dim textX As Integer = swatchRect.Right + 6
			Dim textRect As New Rectangle(textX, e.Bounds.Top, e.Bounds.Width - textX - 4, e.Bounds.Height)
			Dim sf As New StringFormat() With {
				.LineAlignment = StringAlignment.Center,
				.Alignment = StringAlignment.Near,
				.Trimming = StringTrimming.EllipsisCharacter,
				.FormatFlags = StringFormatFlags.NoWrap
			}
			Using b As New SolidBrush(textColor)
				g.DrawString(color.Name, Me.Font, b, textRect, sf)
			End Using

			e.DrawFocusRectangle()
		End Sub

		' Methods
		Private Sub PaintDDLCombo(g As Graphics)
			g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

			Dim rc As Rectangle = Me.ClientRectangle

			' 1. Background (hover aware)
			Dim bg As Color = If(_hovering, GetHoverBackground(Me.BackColor), Me.BackColor)
			Using b As New SolidBrush(bg)
				g.FillRectangle(b, rc)
			End Using

			' 2. Draw selected color swatch + text
			If SelectedIndex >= 0 Then

				Dim color As Color = CType(SelectedItem, Color)
				Dim swatchSize As Integer = rc.Height - 6
				Dim swatchRect As New Rectangle(4, 3, swatchSize, swatchSize)
				Dim radius As Integer = 4 ' adjust for softer or sharper corners
				Dim swatchPath = CreateRoundedRect(swatchRect, radius)
				Using b As New SolidBrush(color)
					g.FillPath(b, swatchPath)
				End Using
				Using p As New Pen(Color.FromArgb(90, 0, 0, 0)) ' softer black
					g.DrawPath(p, swatchPath)
				End Using

				Dim textX As Integer = swatchRect.Right + 6
				Dim textRect As New Rectangle(textX, rc.Top, rc.Width - textX - 20, rc.Height)
				Dim sf As New StringFormat With {.FormatFlags = StringFormatFlags.NoWrap, .Trimming = StringTrimming.EllipsisCharacter, .LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Near}
				Using b As New SolidBrush(Me.ForeColor)
					g.DrawString(color.Name, Me.Font, b, textRect, sf)
				End Using

			End If

			' 3. Arrow
			DrawUnifiedArrow(g, rc)

			' 4. Border
			Using p As New Pen(ControlPaint.Dark(Me.BackColor))
				g.DrawRectangle(p, 0, 0, rc.Width - 1, rc.Height - 1)
			End Using

		End Sub
		Private Sub DrawUnifiedArrow(g As Graphics, rc As Rectangle)
			Dim arrowSize As Integer = 6
			Dim buttonW2 As Integer = SystemInformation.VerticalScrollBarWidth
			Dim arrowX As Integer = rc.Right - buttonW2 \ 2 - arrowSize \ 2 - 2
			Dim arrowY As Integer = (rc.Height - arrowSize) \ 2 + 1
			Dim arrowColor As Color = If(_hovering, ControlPaint.Dark(Me.ForeColor, 0.1F), Me.ForeColor)

			Dim pts As Point() = {
		New Point(arrowX, arrowY),
		New Point(arrowX + arrowSize, arrowY),
		New Point(arrowX + arrowSize \ 2, arrowY + arrowSize)
	}

			Using b As New SolidBrush(arrowColor)
				g.FillPolygon(b, pts)
			End Using
		End Sub
		Private Sub FillColors() 'Must Import System.Linq
			If Not Me.DesignMode Then
				'Populate colors
				Items.Clear()
				'Fill Colors Using Reflection
				For Each color As Color In GetType(Color).GetProperties(Reflection.BindingFlags.[Static] Or Reflection.BindingFlags.[Public]).Where(Function(c) c.PropertyType Is GetType(Color)).[Select](Function(c) CType(c.GetValue(c, Nothing), Color))
					If Not color = Color.Transparent Then Items.Add(color)
				Next
			End If
		End Sub
		Private Function GetHoverBackground(baseColor As Color) As Color
			' If the background is very light, use a subtle gray hover
			If baseColor.R > 240 AndAlso baseColor.G > 240 AndAlso baseColor.B > 240 Then
				Return Color.FromArgb(245, 245, 245) ' light gray hover
			End If

			' Otherwise use the normal lightening
			Return ControlPaint.Light(baseColor, 0.25F)
		End Function
		Private Function CreateRoundedRect(rect As Rectangle, radius As Integer) As Drawing2D.GraphicsPath
			Dim path As New Drawing2D.GraphicsPath()
			Dim d As Integer = radius * 2

			path.AddArc(rect.X, rect.Y, d, d, 180, 90)
			path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
			path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
			path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
			path.CloseFigure()

			Return path
		End Function

	End Class

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
	''' Extended Listview control with Insertion Line for drag/drop operations. The Insertion Line is drawn in the specified color and appears between items to indicate where a dragged item would be dropped.
	''' Also supports inline editing of subitems on double-click, with events for BeforeEdit, AfterEdit, and SubItemEdited. Inline editing allows users to edit subitems directly within the ListView, with customizable editability per column and proper handling of edit lifecycle events.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ListViewEX
		Inherits ListView

		' DECLARATIONS
		' Insertion Line
		Private _LineBefore As Integer = -1
		Private _LineAfter As Integer = -1
		Private _InsertionLineColor As Color = Color.Teal
		' Inline Editing
		Private _lastClickTime As Integer = 0
		Private _lastClickSubItem As Integer = -1
		Private _editBox As TextBox = Nothing
		Private _editItem As ListViewItem = Nothing
		Private _editSubItem As Integer = -1
		Private _editableColumns As New List(Of Boolean)

		' PROPERTIES
		<DefaultValue(-1)>
		Public Property LineBefore As Integer
			Get
				Return _LineBefore
			End Get
			Set(ByVal value As Integer)
				_LineBefore = value
			End Set
		End Property
		<DefaultValue(-1)>
		Public Property LineAfter As Integer
			Get
				Return _LineAfter
			End Get
			Set(ByVal value As Integer)
				_LineAfter = value
			End Set
		End Property
		<Category("Appearance"), Description("Specify the color used to draw the Insertion Line"), DefaultValue(GetType(Color), "Color.Teal")>
		Public Property InsertionLineColor As Color
			Get
				Return _InsertionLineColor
			End Get
			Set(ByVal value As Color)
				_InsertionLineColor = value
				Me.Invalidate()
			End Set
		End Property
		<Category("Behavior"), Description("Specifies which ListView columns are editable."), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property EditableColumns As List(Of Boolean)
			Get
				Return _editableColumns
			End Get
			Set(value As List(Of Boolean))
				_editableColumns = value
			End Set
		End Property

		' EVENTS
		Public Event SubItemEdited(item As ListViewItem, subItemIndex As Integer, newValue As String)
		Public Event BeforeEdit(item As ListViewItem, subItemIndex As Integer, ByRef cancel As Boolean)
		Public Event AfterEdit(item As ListViewItem, subItemIndex As Integer, newValue As String)

		' Control Events
		Protected Overrides Sub WndProc(ByRef m As Message)
			MyBase.WndProc(m)

			If m.Msg = WinAPI.WM_PAINT AndAlso Me.IsHandleCreated AndAlso Not DesignMode Then
				Using g As Graphics = Graphics.FromHwnd(Me.Handle)
					If LineBefore >= 0 AndAlso LineBefore < Items.Count Then
						Dim rc As Rectangle = Items(LineBefore).GetBounds(ItemBoundsPortion.Label)
						DrawInsertionLine(g, rc.Left, rc.Right, rc.Top)
					End If
					If LineAfter >= 0 AndAlso LineAfter < Items.Count Then
						Dim rc As Rectangle = Items(LineAfter).GetBounds(ItemBoundsPortion.Label)
						DrawInsertionLine(g, rc.Left, rc.Right, rc.Bottom)
					End If
				End Using
			End If
		End Sub
		Public Sub New()
			SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
		End Sub
		Protected Overrides Sub OnCreateControl()
			MyBase.OnCreateControl()

			' Designer has not set anything
			If _editableColumns.Count = 0 Then
				For i = 0 To Me.Columns.Count - 1
					_editableColumns.Add(False)
				Next
			End If
		End Sub
		Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

			If keyData = Keys.Tab OrElse keyData = (Keys.Shift Or Keys.Tab) Then
				If _editBox IsNot Nothing Then
					' Simulate a KeyDown event for Tab
					EditBox_KeyDown(_editBox, New KeyEventArgs(keyData))
					Return True
				End If
			End If

			Return MyBase.ProcessCmdKey(msg, keyData)
		End Function
		Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
			MyBase.OnKeyDown(e)

			If e.KeyCode = Keys.F2 Then
				Dim item = FocusedItem
				If item IsNot Nothing Then
					' Edit the first editable column
					For i = 0 To EditableColumns.Count - 1
						If EditableColumns(i) Then
							EditSubItem(item, i)
							Exit For
						End If
					Next
				End If
				e.Handled = True
			End If

		End Sub
		Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
			MyBase.OnMouseDown(e)

			Dim info = HitTest(e.Location)
			If info.Item Is Nothing Then Return

			' Determine subitem index
			Dim subIndex As Integer = info.Item.SubItems.IndexOf(info.SubItem)

			If subIndex < 0 Then
				' Empty cell: determine column manually
				subIndex = GetColumnFromX(e.X)
			End If

			' Optional: per-column editability
			If subIndex < 0 Then Return
			If subIndex < EditableColumns.Count AndAlso Not EditableColumns(subIndex) Then
				Return
			End If

			' Double-click detection using system threshold
			Dim now = Environment.TickCount
			If subIndex = _lastClickSubItem AndAlso (now - _lastClickTime) <= SystemInformation.DoubleClickTime Then
				BeginEditSubItem(info.Item, subIndex)
			End If

			_lastClickTime = now
			_lastClickSubItem = subIndex
		End Sub
		Private Sub EditBox_KeyDown(sender As Object, e As KeyEventArgs)
			If e.KeyCode = Keys.Enter Then
				e.SuppressKeyPress = True
				EndEditSubItem(commit:=True)
				Return
			End If
			If e.KeyCode = Keys.Escape Then
				e.SuppressKeyPress = True
				EndEditSubItem(commit:=False)
				Return
			End If
			' ⭐ TAB NAVIGATION
			If e.KeyCode = Keys.Tab Then
				Debug.Print("Tab pressed in edit box. Shift: " & e.Shift.ToString())
				e.SuppressKeyPress = True
				e.Handled = True

				Dim shift As Boolean = e.Shift
				Dim currentCol = _editSubItem
				Dim item = _editItem

				EndEditSubItem(commit:=True)

				Dim nextCol = FindNextEditableColumn(currentCol, shift)
				If nextCol >= 0 Then
					BeginEditSubItem(item, nextCol)
				End If

				Return
			End If
		End Sub
		Private Sub EditBox_LostFocus(sender As Object, e As EventArgs)
			EndEditSubItem(commit:=True)
		End Sub

		' Handlers
		Private Sub BeginEditSubItem(item As ListViewItem, subIndex As Integer)
			If _editBox IsNot Nothing Then Return

			Dim cancel As Boolean = False
			RaiseEvent BeforeEdit(item, subIndex, cancel)
			If cancel Then Return

			' Ensure the subitem exists
			While item.SubItems.Count <= subIndex
				item.SubItems.Add(String.Empty)
			End While

			_editItem = item
			_editSubItem = subIndex

			Dim bounds = item.SubItems(subIndex).Bounds
			_editBox = New TextBox With {
				.Bounds = bounds,
				.Text = item.SubItems(subIndex).Text,
				.BorderStyle = BorderStyle.FixedSingle,
				.BackColor = Me.BackColor,
				.ForeColor = Me.ForeColor,
				.Font = Me.Font
			}
			If subIndex = 0 Then
				bounds.Width = Me.Columns(0).Width
				_editBox.Width = bounds.Width
			End If

			AddHandler _editBox.LostFocus, AddressOf EditBox_LostFocus
			AddHandler _editBox.KeyDown, AddressOf EditBox_KeyDown
			Me.Controls.Add(_editBox)
			_editBox.Focus()
			_editBox.SelectAll()

		End Sub
		Private Sub EndEditSubItem(commit As Boolean)
			If _editBox Is Nothing Then Return

			If commit AndAlso _editItem IsNot Nothing AndAlso _editSubItem >= 0 Then
				_editItem.SubItems(_editSubItem).Text = _editBox.Text
				RaiseEvent SubItemEdited(_editItem, _editSubItem, _editBox.Text)
				RaiseEvent AfterEdit(_editItem, _editSubItem, _editBox.Text)
			End If

			RemoveHandler _editBox.LostFocus, AddressOf EditBox_LostFocus
			RemoveHandler _editBox.KeyDown, AddressOf EditBox_KeyDown
			Me.Controls.Remove(_editBox)
			_editBox.Dispose()
			_editBox = Nothing
			_editItem = Nothing
			_editSubItem = -1

		End Sub

		' Methods
		Public Sub EditSubItem(item As ListViewItem, subIndex As Integer)
			If item Is Nothing Then Exit Sub
			If subIndex < 0 OrElse subIndex >= item.SubItems.Count Then Exit Sub
			If Not EditableColumns(subIndex) Then Exit Sub

			BeginEditSubItem(item, subIndex)
		End Sub
		Private Sub DrawInsertionLine(g As Graphics, X1 As Integer, X2 As Integer, Y As Integer)
			Using p As New Pen(_InsertionLineColor) With {.Width = 3}
				g.DrawLine(p, X1, Y, X2 - 1, Y)
				Dim leftTriangle As Point() = {New Point(X1, Y - 4), New Point(X1 + 7, Y), New Point(X1, Y + 4)}
				Dim rightTriangle As Point() = {New Point(X2, Y - 4), New Point(X2 - 8, Y), New Point(X2, Y + 4)}
				Using b As New SolidBrush(_InsertionLineColor)
					g.FillPolygon(b, leftTriangle)
					g.FillPolygon(b, rightTriangle)
				End Using
			End Using
		End Sub
		Private Function GetColumnFromX(x As Integer) As Integer
			Dim current As Integer = 0
			For i = 0 To Columns.Count - 1
				current += Columns(i).Width
				If x < current Then Return i
			Next
			Return Columns.Count - 1
		End Function
		Private Function FindNextEditableColumn(current As Integer, backwards As Boolean) As Integer
			If EditableColumns Is Nothing OrElse EditableColumns.Count = 0 Then
				Return -1
			End If

			If backwards Then
				For i = current - 1 To 0 Step -1
					If i < EditableColumns.Count AndAlso EditableColumns(i) Then
						Return i
					End If
				Next
			Else
				For i = current + 1 To Columns.Count - 1
					If i < EditableColumns.Count AndAlso EditableColumns(i) Then
						Return i
					End If
				Next
			End If

			Return -1
		End Function

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
