
Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Namespace UI

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
