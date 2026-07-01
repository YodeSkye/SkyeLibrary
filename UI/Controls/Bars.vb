
Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Namespace Skye.UI

	''' <summary>
	''' Extended Windows progress bar control.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	<DefaultProperty("PercentageMode")>
	Public Class ProgressEX
		Inherits UserControl

		' Declarations
		Public Enum PercentageDrawModes As Integer
			None = 0
			Center
			Movable
		End Enum
		Public Enum ColorDrawModes As Integer
			Gradient = 0
			Smooth
		End Enum
		Private maxValue As Integer
		Private minValue As Integer
		Private _value As Single
		Private stepValue As Integer
		Private percentageValue As Single
		Private m_drawingColor As Color
		Private ReadOnly gradientBlender As Drawing2D.ColorBlend
		Private percentageDrawMode As PercentageDrawModes
		Private colorDrawMode As ColorDrawModes
		Private _Brush As SolidBrush
		Private writingBrush As SolidBrush
		Private writingFont As Font
		Private _Drawer As Drawing2D.LinearGradientBrush

		' Properties
		<Category("Behavior"),
	 Description("Specify how to display the Percentage value"),
	 DefaultValue(PercentageDrawModes.Center)>
		Public Property PercentageMode As PercentageDrawModes
			Get
				Return percentageDrawMode
			End Get
			Set
				percentageDrawMode = Value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"),
	 Description("Specify how to display the Drawing Color"),
	 DefaultValue(ColorDrawModes.Gradient)>
		Public Property DrawingColorMode As ColorDrawModes
			Get
				Return colorDrawMode
			End Get
			Set
				colorDrawMode = Value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"),
	 Description("Specify the color used to draw the progress activities"),
	 DefaultValue(GetType(Color), "Red")>
		Public Property DrawingColor As Color
			Get
				Return m_drawingColor
			End Get
			Set
				m_drawingColor = Value

				If gradientBlender IsNot Nothing Then
					gradientBlender.Colors(0) = ControlPaint.Dark(Value)
					gradientBlender.Colors(1) = ControlPaint.Light(Value)
					gradientBlender.Colors(2) = ControlPaint.Dark(Value)
				End If

				_Brush?.Dispose()
				_Brush = New SolidBrush(Value)

				If _Drawer IsNot Nothing Then
					_Drawer.Dispose()
					_Drawer = New Drawing2D.LinearGradientBrush(
					New Rectangle(Point.Empty, If(ClientSize.IsEmpty, New Size(1, 1), ClientSize)),
					ControlPaint.Dark(Value),
					ControlPaint.Light(Value),
					Drawing2D.LinearGradientMode.Vertical) With {.InterpolationColors = gradientBlender}
				End If

				Invalidate()
			End Set
		End Property
		<Category("Layout"),
	 Description("Specify the maximum value the progress can increased to"),
	 DefaultValue(100)>
		Public Property Maximum As Integer
			Get
				Return maxValue
			End Get
			Set
				maxValue = Value
				If maxValue < minValue Then minValue = maxValue
				If _value > maxValue Then _value = maxValue
				Invalidate()
			End Set
		End Property
		<Category("Layout"),
	 Description("Specify the minimum value the progress can decreased to"),
	 DefaultValue(0)>
		Public Property Minimum As Integer
			Get
				Return minValue
			End Get
			Set
				minValue = Value
				If minValue > maxValue Then maxValue = minValue
				If _value < minValue Then _value = minValue
				Invalidate()
			End Set
		End Property
		<Category("Layout"),
	 Description("Specify the amount by which StepForward increases the current position"),
	 DefaultValue(5)>
		Public Property [Step] As Integer
			Get
				Return stepValue
			End Get
			Set
				stepValue = Value
			End Set
		End Property
		<Category("Layout"),
	 Description("Specify the current position of the progress bar"),
	 DefaultValue(0)>
		Public Property Value As Integer
			Get
				Return CInt(Math.Truncate(_value))
			End Get
			Set
				_value = Math.Max(minValue, Math.Min(maxValue, Value))
				Invalidate()
			End Set
		End Property
		Public ReadOnly Property Percent As Integer
			Get
				Return CInt(Math.Truncate(Math.Round(percentageValue)))
			End Get
		End Property
		<Category("Appearance"),
	 Description("Specify the font used to draw the Percentage value")>
		Public Overrides Property Font As Font
			Get
				Return writingFont
			End Get
			Set
				writingFont = Value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"),
	 Description("Specify the color used to draw the Percentage value")>
		Public Overrides Property ForeColor As Color
			Get
				Return writingBrush.Color
			End Get
			Set
				writingBrush.Color = Value
				Invalidate()
			End Set
		End Property

		' Constructor / Dispose
		Public Sub New()
			MyBase.New()

			Name = "ProgressEx"
			DoubleBuffered = True
			SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True)
			SetStyle(ControlStyles.SupportsTransparentBackColor, True)

			BackColor = Color.Transparent
			MinimumSize = New Size(50, 5)
			MaximumSize = New Size(Integer.MaxValue, 40)
			Size = New Size(96, 24)

			maxValue = 100
			minValue = 0
			stepValue = 5
			percentageDrawMode = PercentageDrawModes.Center
			colorDrawMode = ColorDrawModes.Gradient

			m_drawingColor = Color.Red

			gradientBlender = New Drawing2D.ColorBlend() With {
			.Positions = {0.0F, 0.5F, 1.0F},
			.Colors = {
				ControlPaint.Dark(m_drawingColor),
				ControlPaint.Light(m_drawingColor),
				ControlPaint.Dark(m_drawingColor)
			}
		}

			writingFont = New Font("Arial", 10, FontStyle.Bold)
			writingBrush = New SolidBrush(Color.Black)

		End Sub
		Protected Overrides Sub Dispose(disposing As Boolean)
			' If the Designer is disposing us, do NOT run any custom cleanup.
			' This avoids NullReferenceExceptions during design-time teardown.
			If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then
				MyBase.Dispose(disposing)
				Return
			End If

			If disposing Then
				If _Brush IsNot Nothing Then
					_Brush.Dispose()
					_Brush = Nothing
				End If
				If writingBrush IsNot Nothing Then
					writingBrush.Dispose()
					writingBrush = Nothing
				End If
				If writingFont IsNot Nothing Then
					writingFont.Dispose()
					writingFont = Nothing
				End If
				If _Drawer IsNot Nothing Then
					_Drawer.Dispose()
					_Drawer = Nothing
				End If
			End If

			MyBase.Dispose(disposing)
		End Sub

		' Painting / Layout
		Protected Overrides Sub OnResize(e As EventArgs)
			MyBase.OnResize(e)

			' Always safe, both at design time and runtime
			_Drawer?.Dispose()
			_Drawer = New LinearGradientBrush(
				New Rectangle(Point.Empty, If(ClientSize.IsEmpty, New Size(1, 1), ClientSize)),
				ControlPaint.Dark(m_drawingColor),
				ControlPaint.Light(m_drawingColor),
				LinearGradientMode.Vertical)

			If gradientBlender IsNot Nothing Then
				_Drawer.InterpolationColors = gradientBlender
			End If

			Me.Invalidate()
		End Sub
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			MyBase.OnPaint(e)
			EnsureBrushes()

			If _value <= 0 OrElse maxValue <= minValue Then
				Return
			End If

			percentageValue = (_value - minValue) / (maxValue - minValue) * 100.0F
			Dim w = CInt((Width - 1) * (_value - minValue) / (maxValue - minValue))

			Select Case colorDrawMode
				Case ColorDrawModes.Gradient
					e.Graphics.FillRectangle(_Drawer, 0, 0, w, Height)
				Case Else
					e.Graphics.FillRectangle(_Brush, 0, 0, w, Height)
			End Select

			If percentageDrawMode <> PercentageDrawModes.None Then
				Dim txt = CInt(Math.Truncate(percentageValue)).ToString() & "%"
				Dim sz = TextRenderer.MeasureText(txt, writingFont)

				Dim x As Integer
				If percentageDrawMode = PercentageDrawModes.Movable Then
					x = Math.Max(0, Math.Min(w - sz.Width, w - sz.Width \ 2))
				Else
					x = (Width - sz.Width) \ 2
				End If

				Dim y As Integer = (Height - sz.Height) \ 2

				TextRenderer.DrawText(e.Graphics, txt, writingFont, New Point(x, y), writingBrush.Color)
			End If
		End Sub

		' Methods
		Private Sub EnsureBrushes()
			If _Brush Is Nothing Then _Brush = New SolidBrush(m_drawingColor)
			If _Drawer Is Nothing Then
				Dim sz = If(ClientSize.IsEmpty, New Size(1, 1), ClientSize)
				_Drawer = New LinearGradientBrush(New Rectangle(Point.Empty, sz),
										  ControlPaint.Dark(m_drawingColor),
										  ControlPaint.Light(m_drawingColor),
										  LinearGradientMode.Vertical) With {.InterpolationColors = gradientBlender}
			End If
		End Sub
		Public Sub StepForward()
			If (_value + stepValue) < maxValue Then
				_value += stepValue
			Else
				_value = maxValue
			End If
			Invalidate()
		End Sub
		Public Sub StepBackward()
			If (_value - stepValue) > minValue Then
				_value -= stepValue
			Else
				_value = minValue
			End If
			Invalidate()
		End Sub

	End Class

	''' <summary>
	''' A versatile data bar control that supports horizontal and vertical orientations, solid or gradient fills, optional percentage text display, and customizable colors and corner radius.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class DataBarEX
		Inherits Control

		' DECLARATIONS
		Public Enum GradientMode
			Flat
			Micro
			Horizontal
			Vertical
		End Enum
		Public Enum OrientationMode
			Horizontal
			VerticalUp
			VerticalDown
		End Enum
		Public Enum TextPositionMode
			Centered
			InsideBar
		End Enum
		Public Enum TextDisplayFormats
			ValueOnly
			ValueAndMaximum
			Percentage
		End Enum

		' Fields
		Private _value As Integer
		Private _maximum As Integer = 100
		Private _orientation As OrientationMode = OrientationMode.Horizontal
		Private _text As String = ""
		Private _textDisplayFormat As TextDisplayFormats = TextDisplayFormats.ValueAndMaximum
		Private _barColor As Color = Color.IndianRed
		Private _barBackColor As Color = Color.FromArgb(40, 40, 40)
		Private _gradient As GradientMode = GradientMode.Flat
		Private _gradientStart As Color = Color.IndianRed
		Private _gradientEnd As Color = Color.PaleVioletRed
		Private _innerHighlight As Boolean = False
		Private _bevel As Boolean = False
		Private _trailingGlow As Boolean = False
		Private _showText As Boolean = False
		Private _textPosition As TextPositionMode = TextPositionMode.Centered
		Private _textColor As Color = Color.White
		Private _autoTextColor As Boolean = False

		' DESIGNER PROPERTIES
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Overrides Property BackColor As Color
			Get
				Return Color.Transparent
			End Get
			Set(value As Color)
				MyBase.BackColor = Color.Transparent
			End Set
		End Property
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Overrides Property ForeColor As Color
			Get
				Return Color.Transparent
			End Get
			Set(value As Color)
				MyBase.ForeColor = Color.Transparent
			End Set
		End Property
		<Category("Behavior"), Description("Current value of the bar"), DefaultValue(0)>
		Public Property Value As Integer
			Get
				Return _value
			End Get
			Set(v As Integer)
				_value = Math.Max(0, Math.Min(v, _maximum))
				Invalidate()
			End Set
		End Property
		<Category("Behavior"), Description("Maximum value of the bar"), DefaultValue(100)>
		Public Property Maximum As Integer
			Get
				Return _maximum
			End Get
			Set(v As Integer)
				_maximum = Math.Max(1, v)
				If _value > _maximum Then _value = _maximum
				Invalidate()
			End Set
		End Property
		<Category("Behavior"), Description("Specify the orientation of the bar"), DefaultValue(OrientationMode.Horizontal)>
		Public Property Orientation As OrientationMode
			Get
				Return _orientation
			End Get
			Set(value As OrientationMode)
				_orientation = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Text displayed on the bar, null to display current value"), DefaultValue("")>
		Public Overrides Property Text As String
			Get
				Return _text
			End Get
			Set(value As String)
				If value <> _text Then
					_text = value
					Invalidate()
				End If
			End Set
		End Property
		<Category("Appearance"), Description("Specify the display format of the text on the bar"), DefaultValue(TextDisplayFormats.ValueAndMaximum)>
		Public Property TextDisplayFormat As TextDisplayFormats
			Get
				Return _textDisplayFormat
			End Get
			Set(value As TextDisplayFormats)
				If value <> _textDisplayFormat Then
					_textDisplayFormat = value
					Invalidate()
				End If
			End Set
		End Property
		<Category("Appearance"), Description("Solid bar color when no gradient is used."), DefaultValue(GetType(Color), "IndianRed")>
		Public Property BarColor As Color
			Get
				Return _barColor
			End Get
			Set(value As Color)
				If value <> _barColor Then
					_barColor = value
					Invalidate()
				End If
			End Set
		End Property
		<Category("Appearance"), Description("Background color of the rounded bar shape."), DefaultValue(GetType(Color), "40, 40, 40")>
		Public Property BarBackColor As Color
			Get
				Return _barBackColor
			End Get
			Set(value As Color)
				_barBackColor = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Specify the gradient mode of the bar"), DefaultValue(GradientMode.Flat)>
		Public Property BarGradient As GradientMode
			Get
				Return _gradient
			End Get
			Set(value As GradientMode)
				_gradient = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Start color of the bar gradient."), DefaultValue(GetType(Color), "IndianRed")>
		Public Property GradientStart As Color
			Get
				Return _gradientStart
			End Get
			Set(value As Color)
				If value <> _gradientStart Then
					_gradientStart = value
					Invalidate()
				End If
			End Set
		End Property
		<Category("Appearance"), Description("End color of the bar gradient."), DefaultValue(GetType(Color), "PaleVioletRed")>
		Public Property GradientEnd As Color
			Get
				Return _gradientEnd
			End Get
			Set(value As Color)
				If value <> _gradientEnd Then
					_gradientEnd = value
					Invalidate()
				End If
			End Set
		End Property
		<Category("Appearance"), Description("Draw a subtle highlight line along the top edge of the bar"), DefaultValue(False)>
		Public Property InnerHighlight As Boolean
			Get
				Return _innerHighlight
			End Get
			Set(value As Boolean)
				_innerHighlight = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Draw a subtle bevel on the rounded trailing edge"), DefaultValue(False)>
		Public Property Bevel As Boolean
			Get
				Return _bevel
			End Get
			Set(value As Boolean)
				_bevel = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Draw a soft glow near the trailing edge of the bar"), DefaultValue(False)>
		Public Property TrailingGlow As Boolean
			Get
				Return _trailingGlow
			End Get
			Set(value As Boolean)
				_trailingGlow = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Specify whether to show the text overlay (e.g. '50/100')"), DefaultValue(False)>
		Public Property ShowText As Boolean
			Get
				Return _showText
			End Get
			Set(value As Boolean)
				_showText = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Specify the position of the text overlay"), DefaultValue(TextPositionMode.Centered)>
		Public Property TextPosition As TextPositionMode
			Get
				Return _textPosition
			End Get
			Set(value As TextPositionMode)
				_textPosition = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Specify the color of the text overlay"), DefaultValue(GetType(Color), "White")>
		Public Property TextColor As Color
			Get
				Return _textColor
			End Get
			Set(value As Color)
				_textColor = value
				Invalidate()
			End Set
		End Property
		<Category("Appearance"), Description("Specify whether to automatically adjust the text color for better contrast"), DefaultValue(False)>
		Public Property AutoTextColor As Boolean
			Get
				Return _autoTextColor
			End Get
			Set(value As Boolean)
				_autoTextColor = value
				Invalidate()
			End Set
		End Property

		' EVENTS
		Public Sub New()
			SetStyle(ControlStyles.UserPaint Or
				 ControlStyles.AllPaintingInWmPaint Or
				 ControlStyles.OptimizedDoubleBuffer Or
				 ControlStyles.ResizeRedraw Or
				 ControlStyles.SupportsTransparentBackColor, True)

			BackColor = Color.Transparent ' locked, don't allow change
			ForeColor = Color.Transparent ' locked, don't allow change

		End Sub
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			MyBase.OnPaint(e)

			Dim g = e.Graphics
			g.SmoothingMode = SmoothingMode.AntiAlias

			DrawBackgroundBar(g, Me.ClientRectangle)

			If _value <= 0 OrElse _maximum <= 0 Then
				If _showText Then DrawText(g, Rectangle.Empty)
				Return
			End If

			Dim pct As Double = _value / _maximum
			Dim barRect As Rectangle
			Select Case _orientation
				Case OrientationMode.Horizontal
					barRect = New Rectangle(0, 0, CInt(ClientSize.Width * pct), ClientSize.Height)
				Case OrientationMode.VerticalUp
					Dim h As Integer = CInt(ClientSize.Height * pct)
					barRect = New Rectangle(0, ClientSize.Height - h, ClientSize.Width, h)
				Case OrientationMode.VerticalDown
					Dim h As Integer = CInt(ClientSize.Height * pct)
					barRect = New Rectangle(0, 0, ClientSize.Width, h)
			End Select

			DrawForegroundBar(g, barRect)

			If _showText Then DrawText(g, barRect)

		End Sub

		' METHODS
		Private Sub DrawBackgroundBar(g As Graphics, rect As Rectangle)
			If rect.Width <= 0 OrElse rect.Height <= 0 Then Return

			Dim collapsedRect As Rectangle
			Dim radius As Integer
			If _orientation = OrientationMode.Horizontal Then
				' HORIZONTAL GEOMETRY
				Dim fullHeight As Integer = rect.Height
				Dim width As Integer = rect.Width

				Dim h As Integer = Math.Min(fullHeight, width)
				Dim y As Integer = rect.Top + (fullHeight - h) \ 2

				collapsedRect = New Rectangle(rect.Left, y, width, h)
				radius = h
			Else
				' VERTICAL GEOMETRY
				Dim fullWidth As Integer = rect.Width
				Dim height As Integer = rect.Height

				Dim w As Integer = Math.Min(fullWidth, height)
				Dim x As Integer = rect.Left + (fullWidth - w) \ 2

				collapsedRect = New Rectangle(x, rect.Top, w, height)
				radius = w
			End If
			collapsedRect = Rectangle.Inflate(collapsedRect, -1, -1)

			Using path As GraphicsPath = CreateRoundRect(collapsedRect, radius)
				Using br As New SolidBrush(_barBackColor)
					g.FillPath(br, path)
				End Using
			End Using
		End Sub
		Private Sub DrawForegroundBar(g As Graphics, rect As Rectangle)
			If rect.Width <= 0 OrElse rect.Height <= 0 Then Return

			Dim collapsedRect As Rectangle
			Dim radius As Integer
			If _orientation = OrientationMode.Horizontal Then
				' HORIZONTAL GEOMETRY
				Dim fullHeight As Integer = rect.Height
				Dim width As Integer = rect.Width

				Dim h As Integer = Math.Min(fullHeight, width)
				Dim y As Integer = rect.Top + (fullHeight - h) \ 2

				collapsedRect = New Rectangle(rect.Left, y, width, h)
				radius = h
			Else
				' VERTICAL GEOMETRY
				Dim fullWidth As Integer = rect.Width
				Dim height As Integer = rect.Height

				Dim w As Integer = Math.Min(fullWidth, height)
				Dim x As Integer = rect.Left + (fullWidth - w) \ 2

				collapsedRect = New Rectangle(x, rect.Top, w, height)
				radius = w
			End If
			collapsedRect = Rectangle.Inflate(collapsedRect, -1, -1)

			Using path As GraphicsPath = CreateRoundRect(collapsedRect, radius)
				If _gradient = GradientMode.Flat Then
					Using br As New SolidBrush(_barColor)
						g.FillPath(br, path)
					End Using
				ElseIf _gradient = GradientMode.Micro Then
					Using lg As New LinearGradientBrush(
						collapsedRect,
						ControlPaint.Light(_barColor, 0.25F),   ' lighter top
						ControlPaint.Dark(_barColor, 0.05F),    ' slightly darker bottom
						LinearGradientMode.Vertical
					)
						g.FillPath(lg, path)
					End Using
				Else
					Dim mode As LinearGradientMode = If(_gradient = GradientMode.Horizontal, LinearGradientMode.Horizontal, LinearGradientMode.Vertical)
					Using lg As New LinearGradientBrush(collapsedRect, _gradientStart, _gradientEnd, mode)
						g.FillPath(lg, path)
					End Using
				End If
			End Using

			' Inner Highlight (top edge)
			If _innerHighlight Then
				Using p As New Pen(Color.FromArgb(80, Color.White), 1)
					Dim y As Integer = collapsedRect.Top + 1
					Dim left As Integer = collapsedRect.Left + 1
					Dim right As Integer = collapsedRect.Right - 1
					' Only draw if there's enough height
					If collapsedRect.Height > 4 Then
						g.DrawLine(p, left, y, right, y)
					End If
				End Using
			End If

			' Micro-bevel (subtle shadow on leading edge)
			If _bevel Then
				Using bevelPen As New Pen(Color.FromArgb(60, ControlPaint.Dark(_barColor)), 1)
					Dim inset As Integer = 1
					' Top-right arc segment
					g.DrawArc(bevelPen,
					  collapsedRect.Right - radius + inset,
					  collapsedRect.Top + inset,
					  radius - inset * 2,
					  radius - inset * 2,
					  270,
					  90)
					' Bottom-right arc segment
					g.DrawArc(bevelPen,
					  collapsedRect.Right - radius + inset,
					  collapsedRect.Bottom - radius + inset,
					  radius - inset * 2,
					  radius - inset * 2,
					  0,
					  90)
				End Using
			End If

			' Trailing Glow (soft ellipse)
			If _trailingGlow Then
				' Glow size relative to bar geometry
				Dim glowHeight As Integer = CInt(collapsedRect.Height * 0.55)
				Dim glowWidth As Integer = CInt(radius * 0.9)
				' Position: locked to trailing edge, centered vertically
				Dim glowX As Integer = collapsedRect.Right - glowWidth - 1
				Dim glowY As Integer = collapsedRect.Top + (collapsedRect.Height - glowHeight) \ 2
				Dim glowRect As New Rectangle(glowX, glowY, glowWidth, glowHeight)
				Using glowBrush As New SolidBrush(Color.FromArgb(40, ControlPaint.Light(_barColor)))
					g.FillEllipse(glowBrush, glowRect)
				End Using
			End If

		End Sub
		Private Sub DrawText(g As Graphics, barRect As Rectangle)
			Dim txt As String = ""
			If String.IsNullOrWhiteSpace(Text) Then
				Select Case _textDisplayFormat
					Case TextDisplayFormats.ValueOnly
						txt = _value.ToString()
					Case TextDisplayFormats.ValueAndMaximum
						txt = $"{_value}/{_maximum}"
					Case TextDisplayFormats.Percentage
						txt = $"{CInt((_value / _maximum) * 100)}%"
				End Select
			Else
				txt = Text
			End If
			Dim colorToUse As Color = _textColor
			If _autoTextColor Then colorToUse = ComputeAutoContrastColor()

			Using br As New SolidBrush(colorToUse)
				Using sf As New StringFormat With {
					.Alignment = StringAlignment.Center,
					.LineAlignment = StringAlignment.Center
				}
					Dim rect As Rectangle = ClientRectangle
					Select Case _textPosition
						Case TextPositionMode.Centered
							g.DrawString(txt, Font, br, rect, sf)
						Case TextPositionMode.InsideBar
							Dim size As SizeF = g.MeasureString(txt, Font)
							' Can we center inside the bar?
							Dim canCenter As Boolean = barRect.Width >= size.Width + 8
							If canCenter Then
								' Center inside the bar using your existing StringFormat
								g.DrawString(txt, Font, br, barRect, sf)
							Else
								' Bar too small → draw text just to the right of the bar
								Dim x As Integer = barRect.Right + 2
								' Clamp so text never goes off the control
								x = Math.Min(x, Me.Width - CInt(size.Width))
								Dim y As Integer = (Height - CInt(size.Height)) \ 2
								g.DrawString(txt, Font, br, x, y)
							End If
					End Select
				End Using
			End Using
		End Sub

		' HELPERS
		Private Shared Function CreateRoundRect(rect As Rectangle, radius As Integer) As GraphicsPath
			Dim path As New GraphicsPath()
			Dim d As Integer = radius

			path.AddArc(rect.X, rect.Y, d, d, 180, 90)
			path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
			path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
			path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
			path.CloseFigure()

			Return path
		End Function
		Private Function ComputeAutoContrastColor() As Color
			Dim c As Color
			If _gradient = GradientMode.Flat Then
				c = BarColor
			Else
				' midpoint of gradient (safe integer math)
				c = Color.FromArgb(
					CInt((CInt(GradientStart.R) + CInt(GradientEnd.R)) \ 2),
					CInt((CInt(GradientStart.G) + CInt(GradientEnd.G)) \ 2),
					CInt((CInt(GradientStart.B) + CInt(GradientEnd.B)) \ 2)
				)
			End If
			Dim luminance As Double = (0.299 * c.R) + (0.587 * c.G) + (0.114 * c.B)

			Return If(luminance < 128, Color.White, Color.Black)
		End Function

	End Class

	''' <summary>
	''' A simple ProgressBar with an alternate theme applied for a cleaner, modern look, with rounded corners and highlights, and the ability to change the height.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class DataBar
		Inherits Control

		Private _value As Integer = 0
		Private _maximum As Integer = 100
		Private ReadOnly pulseTimer As Timer
		Private pulseProgress As Double = 0.0
		Private pulseActive As Boolean = False
		Private lastFillWidth As Integer = -1

		<Category("Behavior"), Description("The maximum value of the DataBar."), DefaultValue(100), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Maximum As Integer
			Get
				Return _maximum
			End Get
			Set(value As Integer)
				_maximum = Math.Max(1, value)
				_value = Math.Min(_value, _maximum)
				Invalidate()
			End Set
		End Property
		<Category("Behavior"), Description("The current value of the DataBar."), DefaultValue(0), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Value As Integer
			Get
				Return _value
			End Get
			Set(value As Integer)
				Dim newVal = Math.Max(0, Math.Min(value, _maximum))
				If newVal <> _value Then
					_value = newVal
					' Compute new width
					Dim pct As Double = CDbl(_value) / _maximum
					Dim newWidth As Integer = CInt(pct * Me.Width)
					' Only pulse if the bar actually changed visually
					If newWidth <> lastFillWidth Then
						StartPulse()
						lastFillWidth = newWidth
					End If
					Invalidate()
				End If
			End Set
		End Property

		Public Sub New()
			SetStyle(ControlStyles.AllPaintingInWmPaint Or
				 ControlStyles.OptimizedDoubleBuffer Or
				 ControlStyles.UserPaint Or
				 ControlStyles.ResizeRedraw, True)
			DoubleBuffered = True
			pulseTimer = New Timer With {
				.Interval = 15 ' ~60 FPS
				}
			AddHandler pulseTimer.Tick, AddressOf PulseTick
		End Sub

		Protected Overrides Sub OnPaintBackground(pevent As PaintEventArgs)
			' Only paint background if parent would overwrite us
			If Parent Is Nothing Then Return

			Using b As New SolidBrush(BackColor)
				pevent.Graphics.FillRectangle(b, ClientRectangle)
			End Using
		End Sub
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

			Dim pct As Double = CDbl(_value) / _maximum
			Dim fillWidth As Integer = CInt(pct * Me.Width)
			If fillWidth <= 0 Then Exit Sub

			Dim rect As New Rectangle(0, 0, fillWidth, Me.Height)
			If rect.Width <= 0 OrElse rect.Height <= 0 Then Exit Sub

			Dim radius As Integer = Me.Height
			Dim r As Integer = Math.Min(radius, rect.Height)

			' Build the bar path (flat left, rounded right)
			Dim path As New Drawing2D.GraphicsPath()
			Dim left As Integer = rect.X
			Dim top As Integer = rect.Y
			Dim right As Integer = rect.Right
			Dim bottom As Integer = rect.Bottom
			path.StartFigure()
			path.AddLine(left, top, right - r, top)
			path.AddArc(right - r, top, r, r, 270, 90)
			path.AddLine(right, top + r, right, bottom - r)
			path.AddArc(right - r, bottom - r, r, r, 0, 90)
			path.AddLine(right - r, bottom, left, bottom)
			path.CloseFigure()

			' -----------------------------------------
			' 1. STRONGER MICRO-GRADIENT (visible but modern)
			' -----------------------------------------
			Using lg As New LinearGradientBrush(
				rect,
				ControlPaint.Light(Me.ForeColor, 0.25),   ' brighter top
				ControlPaint.Dark(Me.ForeColor, 0.05),    ' slightly darker bottom
				90.0F
			)
				e.Graphics.FillPath(lg, path)
			End Using

			' -----------------------------------------
			' 2. SUBTLE INNER HIGHLIGHT (top edge)
			' -----------------------------------------
			Using p As New Pen(Color.FromArgb(80, Color.White), 1)
				Dim y As Integer = top + 1
				e.Graphics.DrawLine(p, left + 1, y, right - 1, y)
			End Using

			' -----------------------------------------
			' 3. SOFT GLOW ON TRAILING EDGE (tongue-proof)
			' -----------------------------------------
			Dim glowHeight As Integer = CInt(rect.Height * 0.55)   ' ~55% of bar height
			Dim glowWidth As Integer = CInt(r * 0.9)                ' narrower than radius
			Dim glowX As Integer = right - glowWidth - 1            ' NEVER crosses right edge
			Dim glowY As Integer = top + (rect.Height - glowHeight) \ 2
			Dim glowRect As New Rectangle(glowX, glowY, glowWidth, glowHeight)
			Using glowBrush As New SolidBrush(Color.FromArgb(40, ControlPaint.Light(Me.ForeColor)))
				e.Graphics.FillEllipse(glowBrush, glowRect)
			End Using

			' -----------------------------------------
			' 4. MICRO-BEVEL (original subtle version)
			' -----------------------------------------
			Using bevelPen As New Pen(Color.FromArgb(90, ControlPaint.Dark(Me.ForeColor)), 1)
				Dim inset As Integer = 1
				' Top-right arc segment
				e.Graphics.DrawArc(bevelPen,
					   right - r + inset,
					   top + inset,
					   r - inset * 2,
					   r - inset * 2,
					   270,
					   90)
				' Bottom-right arc segment
				e.Graphics.DrawArc(bevelPen,
					   right - r + inset,
					   bottom - r + inset,
					   r - inset * 2,
					   r - inset * 2,
					   0,
					   90)
			End Using

			' -----------------------------------------
			' 5. PULSE GLOW
			' -----------------------------------------
			If pulseActive Then

				' Very small expansion
				Dim pulseScaleW As Double = 1.0 + (pulseProgress * 0.1) ' 10% wider
				Dim pulseScaleH As Double = 1.0 + (pulseProgress * 0.25) ' 25% taller

				' Very faint alpha
				Dim pulseAlpha As Integer = CInt(40 * (1.0 - pulseProgress))

				' Slightly brighter than normal glow, but not LightLight
				Dim pulseColor As Color = ControlPaint.Light(Me.ForeColor)

				Dim pw As Integer = CInt(glowWidth * pulseScaleW)
				Dim ph As Integer = CInt(glowHeight * pulseScaleH)

				' Keep right edge locked
				Dim px As Integer = right - pw - 1
				Dim py As Integer = top + (rect.Height - ph) \ 2

				Dim pulseRect As New Rectangle(px, py, pw, ph)

				Using pulseBrush As New SolidBrush(Color.FromArgb(pulseAlpha, pulseColor))
					e.Graphics.FillEllipse(pulseBrush, pulseRect)
				End Using

			End If

		End Sub

		Private Sub PulseTick(sender As Object, e As EventArgs)
			pulseProgress += 0.08 ' controls speed

			If pulseProgress >= 1.0 Then
				pulseProgress = 1.0
				pulseActive = False
				pulseTimer.Stop()
			End If

			Invalidate()
		End Sub

		Private Sub StartPulse()
			pulseProgress = 0.0
			pulseActive = True
			pulseTimer.Start()
		End Sub

	End Class

	''' <summary>
	''' A simple ProgressBar with an alternate theme applied for a cleaner, modern look.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class DataBarBasic
		Inherits System.Windows.Forms.ProgressBar

		Public Sub New()
			Me.Style = ProgressBarStyle.Continuous
		End Sub
		Protected Overrides Sub CreateHandle()
			MyBase.CreateHandle()
			If Not Me.DesignMode Then
				Try
					Dim result As Integer = WinAPI.SetWindowTheme(Me.Handle, "", "")
					Debug.WriteLine($"SetWindowTheme result: {result}")
				Catch
				End Try
			End If
		End Sub

	End Class

End Namespace
