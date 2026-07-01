
Imports System.ComponentModel

Namespace Skye.UI

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

End Namespace
