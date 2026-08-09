
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Skye.UI

	' Public API
	Public NotInheritable Class Toast

		Private Sub New()
		End Sub

		Public Shared Sub ShowToast(options As ToastOptions)
			ToastManager.Show(options)
		End Sub

	End Class

	' Public Options Class
	Public Enum ToastLocation
		TopLeft
		TopCenter
		TopRight
		MiddleLeft
		MiddleCenter
		MiddleRight
		BottomLeft
		BottomCenter
		BottomRight
	End Enum
	Public Class ToastOptions

		' Basic Properties
		Public Property Title As String
		Public Property Message As String


		' Optional Properties
		Public Property PlaySound As Boolean = False
		Public Property Duration As Integer = 3000


		' Display Properties
		Public Property Width As Integer = 300
		Public Property Height As Integer = 80
		Public Property Margin As Integer = 10
		Public Property Image As Image = Nothing
		Public Property Icon As Icon = Nothing
		Public Property TitleFont As Font = New Font("Segoe UI", 10, FontStyle.Bold)
		Public Property MessageFont As Font = New Font("Segoe UI", 10, FontStyle.Regular)
		Public Property BackColor As Color = Color.FromArgb(40, 40, 40)
		Public Property BorderColor As Color = Color.White
		Public Property ForeColor As Color = Color.White
		Public Property Location As ToastLocation = ToastLocation.BottomRight
		Public Property BorderWidth As Integer = 2
		''' <summary>
		''' Sets the corner radius for rounded corners of the toast window. Zero value means no rounding (square corners).
		''' </summary>
		''' <returns></returns>
		Public Property CornerRadius As Integer = 16
		Public Property Shadow As Boolean = True

		' Click Action
		Public Property ClickAction As Action = Nothing

	End Class

	' Manager
	Friend Class ToastManager

		Private Shared _activeToast As ToastWindow
		Private Shared _margin As Integer

		Public Shared Sub Show(opts As ToastOptions)
			_margin = opts.Margin

			If opts.PlaySound Then
				System.Media.SystemSounds.Hand.Play()
			End If

			' Close previous toast immediately
			If _activeToast IsNot Nothing Then
				_activeToast.CloseImmediately()
				_activeToast = Nothing
			End If

			' Create new toast
			Dim toast As New ToastWindow(opts)
			_activeToast = toast

			' Initial anchor guess — final position is computed inside ToastWindow
			Dim anchor As Point = ComputePosition(opts.Width, opts.Height)
			toast.TargetPosition = anchor
			toast.ShowToastAt(anchor)

			AddHandler toast.ToastClosed,
			Sub()
				If _activeToast Is toast Then
					_activeToast = Nothing
				End If
			End Sub
		End Sub

		' Initial anchor guess only — final positioning happens in ToastWindow
		Private Shared Function ComputePosition(width As Integer, height As Integer) As Point
			Dim area = Screen.PrimaryScreen.WorkingArea
			Return New Point(area.Right - width - _margin, area.Bottom - height - _margin)
		End Function

	End Class

	' Win32 Layered Window
	Friend Class LayeredToastWindow
		Implements IDisposable

		Private _hwnd As IntPtr
		Private ReadOnly _className As String
		Private ReadOnly _wndProc As WinAPI.WndProcDelegate
		Private ReadOnly _opts As ToastOptions

		Public _width As Integer
		Public _height As Integer
		Private _titleHeight As Integer
		Private _messageHeight As Integer
		Private _cachedIconBmp As Bitmap = Nothing
		Private _iconSize As Integer = 0
		Private _opacity As Byte = 0

		' ---- Layout Constants ----
		Private Const TOAST_PADDING As Integer = 10
		Private Const TITLE_MESSAGE_GAP As Integer = 5
		Private Const ICON_SIZE As Integer = 48
		Private Const MAX_HEIGHT_PADDING As Integer = 40

		Private _lastPos As System.Drawing.Point
		Private _hasInitialPosition As Boolean = False

		Public Event ToastClosed()

		' ------------- WndProc ---------------------------
		Private Function WindowProc(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr

			Select Case msg
				Case WinAPI.WM_NCHITTEST
					Return CType(WinAPI.HTCAPTION, IntPtr)
				Case WinAPI.WM_NCLBUTTONDOWN
					OnToastClicked()
					CloseToast()
					Return IntPtr.Zero
				Case WinAPI.WM_NCRBUTTONUP
					OnToastClicked()
					CloseToast()
					Return IntPtr.Zero
				Case WinAPI.WM_MOUSEACTIVATE
					Return CType(WinAPI.MA_NOACTIVATE, IntPtr)
			End Select

			Return WinAPI.DefWindowProc(hWnd, msg, wParam, lParam)
		End Function

		' ------------- Constructor -------------------------------
		Public Sub New(opts As ToastOptions)
			_opts = opts

			_className = "LayeredToast_" & Guid.NewGuid().ToString("N")
			_wndProc = AddressOf Me.WindowProc

			RegisterWindowClass()
			CreateLayeredWindow()

			' Convert once on creation
			If _opts.Icon IsNot Nothing Then
				_cachedIconBmp = _opts.Icon.ToBitmap()
			End If

			' Now that window exists, measure text and compute final height
			AutoWidthToast()
			AutoSizeToast()

		End Sub
		Private Sub RegisterWindowClass()
			Dim wc As New WinAPI.WNDCLASSEX()
			wc.cbSize = CUInt(Marshal.SizeOf(wc))
			wc.style = 0UI
			wc.lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc)
			wc.hInstance = Marshal.GetHINSTANCE(GetType(LayeredToastWindow).Module)
			wc.lpszClassName = _className
			If WinAPI.RegisterClassEx(wc) = 0US Then
				Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
			End If
		End Sub
		Private Sub CreateLayeredWindow()
			Dim exStyle = WinAPI.WS_EX_TOPMOST Or WinAPI.WS_EX_TOOLWINDOW Or WinAPI.WS_EX_NOACTIVATE Or WinAPI.WS_EX_LAYERED
			Dim style = WinAPI.WS_POPUP

			_hwnd = WinAPI.CreateWindowEx(exStyle,
					   _className,
					   "",
					   style,
					   0, 0, _width, _height,
					   IntPtr.Zero, IntPtr.Zero,
					   Marshal.GetHINSTANCE(GetType(LayeredToastWindow).Module),
					   IntPtr.Zero)

			If _hwnd = IntPtr.Zero Then
				Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
			End If

			' Ensure WS_EX_LAYERED remains set
			Dim currentEx As IntPtr = WinAPI.GetWindowLongPtr(_hwnd, WinAPI.GWL_EXSTYLE)
			Dim newEx As Long = currentEx.ToInt64() Or WinAPI.WS_EX_LAYERED
			WinAPI.SetWindowLongPtr(_hwnd, WinAPI.GWL_EXSTYLE, New IntPtr(newEx))
		End Sub
		' ------------- IDisposable -----------------------
		Public Sub Dispose() Implements IDisposable.Dispose

			If _cachedIconBmp IsNot Nothing Then
				_cachedIconBmp.Dispose()
				_cachedIconBmp = Nothing
			End If

			If _hwnd <> IntPtr.Zero Then
				WinAPI.DestroyWindow(_hwnd)
				_hwnd = IntPtr.Zero
			End If
		End Sub

		Protected Overridable Sub OnToastClicked()
			' Base does nothing — derived class will handle it
		End Sub

		' ------------- Public API -------------------------
		Public Sub ShowAt(screenPos As System.Drawing.Point)
			_opacity = 0

			' 1. Move the window FIRST (no rendering yet)
			MoveTo(screenPos)

			' 2. Now safe to render at the correct position
			RefreshLayer()

		End Sub
		Public Sub MoveTo(screenPos As Point)
			_lastPos = screenPos
			_hasInitialPosition = True

			' Move without resizing — height is controlled by AutoSizeToast
			WinAPI.SetWindowPos(
				_hwnd,
				WinAPI.HWND_TOPMOST,
				screenPos.X,
				screenPos.Y,
				0,
				0,
				WinAPI.SWP_NOACTIVATE Or WinAPI.SWP_SHOWWINDOW Or WinAPI.SWP_NOSIZE
			)

			Debug.WriteLine($"[MoveTo] pos=({_lastPos.X},{_lastPos.Y}) size=({_width},{_height})")
		End Sub
		Public Sub SetOpacity(value As Double)
			If value < 0.0 Then value = 0.0
			If value > 1.0 Then value = 1.0

			Dim byteOpacity As Byte = CByte(value * 255)
			_opacity = byteOpacity

			RefreshLayer()
		End Sub
		Public Sub CloseToast()
			RaiseEvent ToastClosed()

			If _hwnd <> IntPtr.Zero Then
				WinAPI.DestroyWindow(_hwnd)
				_hwnd = IntPtr.Zero
			End If
		End Sub
		Public Sub Destroy()
			If _hwnd <> IntPtr.Zero Then
				WinAPI.DestroyWindow(_hwnd)
				_hwnd = IntPtr.Zero
			End If
		End Sub

		' ------------- Layered Render Pipeline -----------
		Private Sub RefreshLayer()
			' Use the last known position; window already moved via MoveTo
			UpdateBitmapAndApply(_lastPos)
		End Sub

		' ------------- Drawing ---------------------------
		Private Sub UpdateBitmapAndApply(screenPos As System.Drawing.Point)
			If _hwnd = IntPtr.Zero Then Return
			If Not _hasInitialPosition Then Return

			' Shadow padding (only if shadow enabled)
			Dim shadowPad As Integer = If(_opts.Shadow, 8, 0)

			' New bitmap size (toast + shadow padding)
			Dim bmpWidth As Integer = _width + shadowPad * 2
			Dim bmpHeight As Integer = _height + shadowPad * 2

			' Destination on screen must shift UP/LEFT by shadowPad
			Dim dstPoint As New WinAPI.POINT With {
		.X = screenPos.X - shadowPad,
		.Y = screenPos.Y - shadowPad
	}

			' Source always starts at 0,0 inside the bitmap
			Dim srcPoint As New WinAPI.POINT With {.X = 0, .Y = 0}

			' Size passed to UpdateLayeredWindow must match the NEW bitmap size
			Dim size As New WinAPI.SIZE With {.cx = bmpWidth, .cy = bmpHeight}

			Dim hdcScreen = WinAPI.GetDC(IntPtr.Zero)
			If hdcScreen = IntPtr.Zero Then Return

			Dim hresult As Integer
			Dim hdcMem = WinAPI.CreateCompatibleDC(hdcScreen)
			If hdcMem = IntPtr.Zero Then
				hresult = WinAPI.ReleaseDC(IntPtr.Zero, hdcScreen)
				Return
			End If

			' Create ARGB surface with padding
			Using bmp As New Bitmap(bmpWidth, bmpHeight, Imaging.PixelFormat.Format32bppArgb)
				Using g As Graphics = Graphics.FromImage(bmp)
					g.SmoothingMode = SmoothingMode.AntiAlias

					' Shift drawing by shadowPad so toast is centered inside padded bitmap
					g.TranslateTransform(shadowPad, shadowPad)

					' Render toast normally (shadow code inside RenderToast will now fit)
					RenderToast(g)

					' Reset transform
					g.ResetTransform()
				End Using

				Dim hBitmap As IntPtr = bmp.GetHbitmap(Color.FromArgb(0))
				Dim oldObj = WinAPI.SelectObject(hdcMem, hBitmap)

				Dim blend As New WinAPI.BLENDFUNCTION() With {
			.BlendOp = WinAPI.AC_SRC_OVER,
			.BlendFlags = 0,
			.SourceConstantAlpha = _opacity,
			.AlphaFormat = WinAPI.AC_SRC_ALPHA
		}

				Dim ok = WinAPI.UpdateLayeredWindow(
					_hwnd,
					hdcScreen,
					dstPoint,
					size,
					hdcMem,
					srcPoint,
					0,
					blend,
					WinAPI.ULW_ALPHA
				)

				WinAPI.SelectObject(hdcMem, oldObj)
				WinAPI.DeleteObject(hBitmap)
			End Using

			WinAPI.DeleteDC(hdcMem)
			hresult = WinAPI.ReleaseDC(IntPtr.Zero, hdcScreen)
		End Sub
		'Private Sub RenderToast(g As Graphics)
		'	Dim w = _width
		'	Dim h = _height

		'	Dim bgColor = _opts.BackColor
		'	Dim borderColor = _opts.BorderColor
		'	Dim foreColor = _opts.ForeColor
		'	Dim radius = _opts.CornerRadius

		'	g.Clear(Color.Transparent)

		'	' Outer rect
		'	Dim inset As Single = 0.5F
		'	Dim rect As New RectangleF(inset, inset, w - 1 - inset, h - 1 - inset)

		'	'-----------------------------------------
		'	' SHADOW
		'	'-----------------------------------------
		'	If _opts.Shadow Then
		'		Dim shadowOffset As Integer = 4
		'		Dim shadowRect As New RectangleF(rect.X + shadowOffset, rect.Y + shadowOffset, rect.Width, rect.Height)

		'		For i As Integer = 1 To 6
		'			Dim alpha As Integer = 30 - (i * 4)
		'			If alpha < 0 Then alpha = 0

		'			Using shadowBrush As New SolidBrush(Color.FromArgb(alpha, 0, 0, 0))
		'				Dim shadowRadius As Integer = radius + i

		'				If shadowRadius > 0 Then
		'					Using shadowPath As New GraphicsPath()
		'						shadowPath.AddArc(shadowRect.X, shadowRect.Y, shadowRadius, shadowRadius, 180, 90)
		'						shadowPath.AddArc(shadowRect.Right - shadowRadius, shadowRect.Y, shadowRadius, shadowRadius, 270, 90)
		'						shadowPath.AddArc(shadowRect.Right - shadowRadius, shadowRect.Bottom - shadowRadius, shadowRadius, shadowRadius, 0, 90)
		'						shadowPath.AddArc(shadowRect.X, shadowRect.Bottom - shadowRadius, shadowRadius, shadowRadius, 90, 90)
		'						shadowPath.CloseFigure()
		'						g.FillPath(shadowBrush, shadowPath)
		'					End Using
		'				Else
		'					g.FillRectangle(shadowBrush, shadowRect)
		'				End If
		'			End Using

		'			shadowRect.Inflate(1, 1)
		'		Next
		'	End If

		'	'-----------------------------------------
		'	' BACKGROUND
		'	'-----------------------------------------
		'	If radius <= 0 Then
		'		Using bgBrush As New SolidBrush(bgColor)
		'			g.FillRectangle(bgBrush, rect)
		'		End Using
		'		Using pen As New Pen(borderColor, _opts.BorderWidth)
		'			g.DrawRectangle(pen, Rectangle.Round(rect))
		'		End Using
		'	Else
		'		Using path As New GraphicsPath()
		'			path.AddArc(rect.X, rect.Y, radius, radius, 180, 90)
		'			path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90)
		'			path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90)
		'			path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90)
		'			path.CloseFigure()

		'			Using bgBrush As New SolidBrush(bgColor)
		'				g.FillPath(bgBrush, path)
		'			End Using
		'			Using pen As New Pen(borderColor, _opts.BorderWidth)
		'				g.DrawPath(pen, path)
		'			End Using
		'		End Using
		'	End If

		'	'-----------------------------------------
		'	' ICON / IMAGE
		'	'-----------------------------------------
		'	Dim padding As Integer = TOAST_PADDING
		'	Dim textX As Integer
		'	Dim iconRect As Rectangle

		'	If _opts.Image IsNot Nothing Then
		'		Dim size As Integer = h - padding * 2
		'		iconRect = New Rectangle(padding, padding, size, size)

		'		If radius > 0 Then
		'			Using path As New GraphicsPath()
		'				path.AddArc(iconRect.X, iconRect.Y, radius, radius, 180, 90)
		'				path.AddArc(iconRect.Right - radius, iconRect.Y, radius, radius, 270, 90)
		'				path.AddArc(iconRect.Right - radius, iconRect.Bottom - radius, radius, radius, 0, 90)
		'				path.AddArc(iconRect.X, iconRect.Bottom - radius, radius, radius, 90, 90)
		'				path.CloseFigure()

		'				g.SetClip(path)
		'				g.DrawImage(_opts.Image, iconRect)
		'				g.ResetClip()
		'			End Using
		'		Else
		'			g.DrawImage(_opts.Image, iconRect)
		'		End If

		'		textX = iconRect.Right + padding

		'	ElseIf _cachedIconBmp IsNot Nothing Then
		'		Dim size As Integer = ICON_SIZE
		'		Dim iconY As Integer = padding + ((h - (padding * 2) - size) \ 2)
		'		iconRect = New Rectangle(padding, iconY, size, size)

		'		Dim bmpW As Integer = _cachedIconBmp.Width
		'		Dim bmpH As Integer = _cachedIconBmp.Height
		'		Dim offsetX As Integer = iconRect.X + ((size - bmpW) \ 2)
		'		Dim offsetY As Integer = iconRect.Y + ((size - bmpH) \ 2)

		'		g.DrawImage(_cachedIconBmp, New Rectangle(offsetX, offsetY, bmpW, bmpH))
		'		textX = iconRect.Right + 2
		'	Else
		'		' No icon, no image → add horizontal inset
		'		textX = padding + 2
		'	End If

		'	Dim wAvail As Integer = w - textX - padding

		'	'-----------------------------------------
		'	' TITLE (TextRenderer)
		'	'-----------------------------------------
		'	Dim titleY As Integer = padding + TEXT_TOP_OFFSET

		'	If Not String.IsNullOrEmpty(_opts.Title) AndAlso _titleHeight > 0 Then
		'		Dim titleRectF As New RectangleF(textX, padding + 2, wAvail, _titleHeight)
		'		Using brush As New SolidBrush(foreColor)
		'			g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
		'			Using sf As New StringFormat(StringFormatFlags.LineLimit)
		'				sf.Trimming = StringTrimming.EllipsisCharacter
		'				g.DrawString(_opts.Title, _opts.TitleFont, brush, titleRectF, sf)
		'			End Using
		'		End Using
		'	End If

		'	'-----------------------------------------
		'	' MESSAGE (TextRenderer)
		'	'-----------------------------------------
		'	Dim messageY As Integer

		'	If String.IsNullOrEmpty(_opts.Title) Then
		'		' No title: message starts at the normal text top area
		'		messageY = padding + TEXT_TOP_OFFSET + 3
		'	Else
		'		' Normal layout when title exists: message under the title
		'		messageY = padding + TEXT_TOP_OFFSET + _titleHeight + TITLE_MESSAGE_GAP
		'	End If

		'	If Not String.IsNullOrEmpty(_opts.Message) AndAlso _messageHeight > 0 Then
		'		Dim messageRectF As New RectangleF(textX, padding + _titleHeight + TITLE_MESSAGE_GAP, wAvail, _messageHeight)
		'		Using brush As New SolidBrush(foreColor)
		'			g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
		'			Using sf As New StringFormat(StringFormatFlags.LineLimit)
		'				sf.Trimming = StringTrimming.EllipsisWord
		'				g.DrawString(_opts.Message, _opts.MessageFont, brush, messageRectF, sf)
		'			End Using
		'		End Using
		'	End If
		'	'Debug.WriteLine($"HEIGHT = {_height}")
		'	'Debug.WriteLine($"MESSAGEHEIGHT = {_messageHeight}")
		'	'Debug.WriteLine($"ICONSIZE = {_iconSize}")
		'	'Debug.WriteLine($"PADDING = {TOAST_PADDING}")
		'End Sub
		Private Sub RenderToast(g As Graphics)
			Dim w = _width
			Dim h = _height

			Dim bgColor = _opts.BackColor
			Dim borderColor = _opts.BorderColor
			Dim foreColor = _opts.ForeColor
			Dim radius = _opts.CornerRadius

			' ClearTypeGridFit works best on dark/solid backdrops, but on transparent/layered windows, AntiAliasGridFit or AntiAlias prevents fuzzy gray halo artifacts on the edges.
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
			g.Clear(Color.Transparent)

			' Outer rect
			Dim inset As Single = 0.5F
			Dim rect As New RectangleF(inset, inset, w - 1 - inset, h - 1 - inset)

			'-----------------------------------------
			' SHADOW
			'-----------------------------------------
			If _opts.Shadow Then
				Dim shadowOffset As Integer = 4
				Dim shadowRect As New RectangleF(rect.X + shadowOffset, rect.Y + shadowOffset, rect.Width, rect.Height)

				For i As Integer = 1 To 6
					Dim alpha As Integer = 30 - (i * 4)
					If alpha < 0 Then alpha = 0

					Using shadowBrush As New SolidBrush(Color.FromArgb(alpha, 0, 0, 0))
						Dim shadowRadius As Integer = radius + i

						If shadowRadius > 0 Then
							Using shadowPath As New GraphicsPath()
								shadowPath.AddArc(shadowRect.X, shadowRect.Y, shadowRadius, shadowRadius, 180, 90)
								shadowPath.AddArc(shadowRect.Right - shadowRadius, shadowRect.Y, shadowRadius, shadowRadius, 270, 90)
								shadowPath.AddArc(shadowRect.Right - shadowRadius, shadowRect.Bottom - shadowRadius, shadowRadius, shadowRadius, 0, 90)
								shadowPath.AddArc(shadowRect.X, shadowRect.Bottom - shadowRadius, shadowRadius, shadowRadius, 90, 90)
								shadowPath.CloseFigure()
								g.FillPath(shadowBrush, shadowPath)
							End Using
						Else
							g.FillRectangle(shadowBrush, shadowRect)
						End If
					End Using

					shadowRect.Inflate(1, 1)
				Next
			End If

			'-----------------------------------------
			' BACKGROUND
			'-----------------------------------------
			If radius <= 0 Then
				Using bgBrush As New SolidBrush(bgColor)
					g.FillRectangle(bgBrush, rect)
				End Using
				Using pen As New Pen(borderColor, _opts.BorderWidth)
					g.DrawRectangle(pen, Rectangle.Round(rect))
				End Using
			Else
				Using path As New GraphicsPath()
					path.AddArc(rect.X, rect.Y, radius, radius, 180, 90)
					path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90)
					path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90)
					path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90)
					path.CloseFigure()

					Using bgBrush As New SolidBrush(bgColor)
						g.FillPath(bgBrush, path)
					End Using
					Using pen As New Pen(borderColor, _opts.BorderWidth)
						g.DrawPath(pen, path)
					End Using
				End Using
			End If

			'-----------------------------------------
			' ICON / IMAGE
			'-----------------------------------------
			Dim padding As Integer = TOAST_PADDING
			Dim textX As Integer
			Dim iconRect As Rectangle

			If _opts.Image IsNot Nothing Then
				Dim maxContainerSize As Integer = h - (padding * 2)

				' Scale down ONLY if image dimensions exceed container slot
				Dim drawW As Integer = Math.Min(_opts.Image.Width, maxContainerSize)
				Dim drawH As Integer = Math.Min(_opts.Image.Height, maxContainerSize)

				' Preserve aspect ratio if image is non-square
				If _opts.Image.Width <> _opts.Image.Height AndAlso _opts.Image.Width > 0 AndAlso _opts.Image.Height > 0 Then
					Dim ratio As Double = Math.Min(maxContainerSize / CDbl(_opts.Image.Width), maxContainerSize / CDbl(_opts.Image.Height))
					If ratio < 1.0 Then
						drawW = CInt(_opts.Image.Width * ratio)
						drawH = CInt(_opts.Image.Height * ratio)
					Else
						drawW = _opts.Image.Width
						drawH = _opts.Image.Height
					End If
				End If

				' Center the image inside the left slot vertically and horizontally
				Dim imgX As Integer = padding + ((maxContainerSize - drawW) \ 2)
				Dim imgY As Integer = padding + ((maxContainerSize - drawH) \ 2)
				iconRect = New Rectangle(imgX, imgY, drawW, drawH)

				If radius > 0 Then
					' Clip rounded corners proportional to image dimensions
					Dim imgRadius As Integer = Math.Min(radius, Math.Min(drawW, drawH) \ 2)
					If imgRadius > 0 Then
						Using path As New GraphicsPath()
							path.AddArc(iconRect.X, iconRect.Y, imgRadius, imgRadius, 180, 90)
							path.AddArc(iconRect.Right - imgRadius, iconRect.Y, imgRadius, imgRadius, 270, 90)
							path.AddArc(iconRect.Right - imgRadius, iconRect.Bottom - imgRadius, imgRadius, imgRadius, 0, 90)
							path.AddArc(iconRect.X, iconRect.Bottom - imgRadius, imgRadius, imgRadius, 90, 90)
							path.CloseFigure()

							g.SetClip(path)
							g.DrawImage(_opts.Image, iconRect)
							g.ResetClip()
						End Using
					Else
						g.DrawImage(_opts.Image, iconRect)
					End If
				Else
					g.DrawImage(_opts.Image, iconRect)
				End If

				' Keep text boundary fixed relative to full image area so text alignment stays uniform
				textX = padding + maxContainerSize + padding

			ElseIf _cachedIconBmp IsNot Nothing Then
				Dim size As Integer = ICON_SIZE
				Dim iconY As Integer = padding + ((h - (padding * 2) - size) \ 2)
				iconRect = New Rectangle(padding, iconY, size, size)

				Dim bmpW As Integer = _cachedIconBmp.Width
				Dim bmpH As Integer = _cachedIconBmp.Height
				Dim offsetX As Integer = iconRect.X + ((size - bmpW) \ 2)
				Dim offsetY As Integer = iconRect.Y + ((size - bmpH) \ 2)

				g.DrawImage(_cachedIconBmp, New Rectangle(offsetX, offsetY, bmpW, bmpH))
				textX = iconRect.Right + 2
			Else
				' No icon, no image → add horizontal inset
				textX = padding + 2
			End If

			Dim wAvail As Integer = w - textX - padding

			'-----------------------------------------
			' TITLE & MESSAGE RENDERING
			'-----------------------------------------
			Dim hasTitle As Boolean = Not String.IsNullOrEmpty(_opts.Title) AndAlso _titleHeight > 0
			Dim hasMessage As Boolean = Not String.IsNullOrEmpty(_opts.Message) AndAlso _messageHeight > 0

			If hasTitle OrElse hasMessage Then
				Dim totalTextHeight As Integer = 0

				If hasTitle AndAlso hasMessage Then
					totalTextHeight = _titleHeight + TITLE_MESSAGE_GAP + _messageHeight
				ElseIf hasTitle Then
					totalTextHeight = _titleHeight
				ElseIf hasMessage Then
					totalTextHeight = _messageHeight
				End If

				' Vertically center the text (whether it's a single title, single message, or combined block)
				Dim startY As Integer = Math.Max(padding, (h - totalTextHeight) \ 2)

				' Title
				If hasTitle Then
					' Give titleRectF a slight width cushion (wAvail + 8)
					Dim titleRectF As New RectangleF(textX, startY, wAvail + 8, _titleHeight)
					Using brush As New SolidBrush(foreColor)
						Using sf As New StringFormat(StringFormatFlags.LineLimit)
							sf.Trimming = StringTrimming.EllipsisCharacter
							g.DrawString(_opts.Title, _opts.TitleFont, brush, titleRectF, sf)
						End Using
					End Using
				End If

				' Message
				If hasMessage Then
					Dim messageY As Integer = If(hasTitle, startY + _titleHeight + TITLE_MESSAGE_GAP, startY)
					' Give messageRectF a slight width cushion (wAvail + 8)
					Dim messageRectF As New RectangleF(textX, messageY, wAvail + 8, _messageHeight)

					Using brush As New SolidBrush(foreColor)
						Using sf As New StringFormat(StringFormatFlags.LineLimit)
							sf.Trimming = StringTrimming.EllipsisWord
							g.DrawString(_opts.Message, _opts.MessageFont, brush, messageRectF, sf)
						End Using
					End Using
				End If

			End If
		End Sub
		Private Sub AutoWidthToast()
            Dim padding As Integer = TOAST_PADDING
			Dim minWidth As Integer = _opts.Width 'defaulted to 300 in ToastOptions definition
			Dim maxWidth As Integer = 900

			Dim iconSpace As Integer = 0
			If _opts.Image IsNot Nothing OrElse _opts.Icon IsNot Nothing Then
				iconSpace = ICON_SIZE + padding
			End If

			Dim nonTextWidth As Integer = padding + iconSpace + padding
			Dim titleWidth As Single = 0
			Dim messageWidth As Single = 0

			Using bmp As New Bitmap(1, 1)
				Using g As Graphics = Graphics.FromImage(bmp)
					g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

					If Not String.IsNullOrEmpty(_opts.Title) Then
						titleWidth = g.MeasureString(_opts.Title, _opts.TitleFont).Width
					End If

					If Not String.IsNullOrEmpty(_opts.Message) Then
						messageWidth = g.MeasureString(_opts.Message, _opts.MessageFont).Width
					End If
				End Using
			End Using

			' ADD FUDGE FACTOR: +16px prevents GDI+ subpixel boundary clipping
			Dim requiredTextWidth As Integer = CInt(Math.Ceiling(Math.Max(titleWidth, messageWidth))) + 16
			Dim calculatedWidth As Integer = requiredTextWidth + nonTextWidth

			_width = Math.Max(minWidth, Math.Min(calculatedWidth, maxWidth))
		End Sub
		Private Sub AutoSizeToast()
			Dim padding As Integer = TOAST_PADDING

			_titleHeight = 0
			_messageHeight = 0

			' Determine icon/image slot width
			If _opts.Image IsNot Nothing OrElse _opts.Icon IsNot Nothing Then
				_iconSize = ICON_SIZE
			Else
				_iconSize = 0
			End If

			Dim textX As Integer = padding
			If _iconSize > 0 Then
				textX = padding + _iconSize + padding
			End If

			' Available width for text rendering
			Dim wAvail As Integer = _width - textX - padding

			Using bmp As New Bitmap(1, 1)
				Using g As Graphics = Graphics.FromImage(bmp)
					g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

					' Measure Title Height with typographic precision
					If Not String.IsNullOrEmpty(_opts.Title) Then
						Using sf As StringFormat = CType(StringFormat.GenericTypographic.Clone(), StringFormat)
							sf.FormatFlags = StringFormatFlags.LineLimit
							sf.Trimming = StringTrimming.EllipsisCharacter
							Dim sizeF As SizeF = g.MeasureString(_opts.Title, _opts.TitleFont, wAvail, sf)
							' Add +2px height safety buffer so DrawString never clips the bottom
							_titleHeight = CInt(Math.Ceiling(sizeF.Height)) + 2
						End Using
					End If

					' Measure Message Height with typographic precision
					If Not String.IsNullOrEmpty(_opts.Message) Then
						Using sf As StringFormat = CType(StringFormat.GenericTypographic.Clone(), StringFormat)
							sf.FormatFlags = StringFormatFlags.LineLimit
							sf.Trimming = StringTrimming.EllipsisWord
							Dim sizeF As SizeF = g.MeasureString(_opts.Message, _opts.MessageFont, wAvail, sf)
							' Add +2px height safety buffer so DrawString never clips the bottom
							_messageHeight = CInt(Math.Ceiling(sizeF.Height)) + 2
						End Using
					End If
				End Using
			End Using

			' Calculate Total Height
			Dim gap As Integer = If(Not String.IsNullOrEmpty(_opts.Title) AndAlso Not String.IsNullOrEmpty(_opts.Message), TITLE_MESSAGE_GAP, 0)
			Dim totalHeight As Integer = padding + _titleHeight + gap + _messageHeight + padding

			' Clamp to screen working area
			Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
			Dim maxHeight As Integer = wa.Height - MAX_HEIGHT_PADDING

			If totalHeight > maxHeight Then
				totalHeight = maxHeight
				Dim availableForMessage As Integer = totalHeight - padding - _titleHeight - gap - padding
				_messageHeight = Math.Max(availableForMessage, 0)
			End If

			_height = totalHeight

			' Assign final image size based on calculated height
			If _opts.Image IsNot Nothing Then
				_iconSize = _height - (TOAST_PADDING * 2)
			End If
		End Sub

	End Class

	' WinForms Window
	Friend Class ToastWindow
		Inherits LayeredToastWindow

		' Declarations
		Private ReadOnly _opts As ToastOptions
		Private ReadOnly FadeTimer As Timer
		Private ReadOnly LifeTimer As Timer
		Private _opacity As Double = 0.0
		Private _fadingOut As Boolean = False
		Public ReadOnly Property IsFadingOut As Boolean
			Get
				Return _fadingOut
			End Get
		End Property
		Public Property TargetPosition As System.Drawing.Point

		Public Shadows Event ToastClosed()

		' Constructor
		Public Sub New(opts As ToastOptions)
			MyBase.New(opts)
			_opts = opts
			FadeTimer = New Timer() With {.Interval = 15}
			AddHandler FadeTimer.Tick, AddressOf FadeTick
			LifeTimer = New Timer() With {.Interval = _opts.Duration}
			AddHandler LifeTimer.Tick, AddressOf BeginFadeOut
		End Sub

		Protected Overrides Sub OnToastClicked()
			If _opts.ClickAction IsNot Nothing Then _opts.ClickAction.Invoke()
		End Sub

		' Public API
		Public Sub ShowToastAt(anchor As Point)
			_opacity = 0.0
			_fadingOut = False

			' Compute final position using real measured height
			Dim finalPos = ComputeLocationPosition(_opts.Location, _width, _height, anchor)
			TargetPosition = finalPos

			MyBase.ShowAt(finalPos)

			FadeTimer.Start()
			LifeTimer.Start()
		End Sub
		Public Sub CloseToastWindow()
			FadeTimer.Stop()
			LifeTimer.Stop()
			_fadingOut = True
			FadeTimer.Start()
		End Sub
		Public Sub CloseImmediately()

			' Stop all timers
			FadeTimer.Stop()
			LifeTimer.Stop()

			' Skip fade-out, go straight to invisible
			_opacity = 0
			MyBase.SetOpacity(0)

			' Notify manager
			RaiseEvent ToastClosed()

			' Destroy window
			MyBase.Destroy()

		End Sub

		' Fade Logic
		Private Sub FadeTick(sender As Object, e As EventArgs)
			If Not _fadingOut Then
				' Fade in
				If _opacity < 1.0 Then
					_opacity += 0.05
					MyBase.SetOpacity(_opacity)
				Else
					FadeTimer.Stop()

				End If

			Else
				' Fade out
				If _opacity > 0.0 Then
					_opacity -= 0.05
					MyBase.SetOpacity(_opacity)
				Else
					FadeTimer.Stop()
					RaiseEvent ToastClosed()
					MyBase.Destroy()
				End If
			End If
		End Sub
		Private Sub BeginFadeOut(sender As Object, e As EventArgs)
			LifeTimer.Stop()
			_fadingOut = True
			FadeTimer.Start()
		End Sub

		' Methods
		Private Shared Function ComputeLocationPosition(loc As ToastLocation, toastWidth As Integer, toastHeight As Integer, anchor As Point) As Point
			Dim scr = Screen.FromPoint(anchor)
			Dim area = scr.WorkingArea

			Dim x As Integer
			Dim y As Integer
			Dim margin As Integer = 20

			Debug.WriteLine($"[ComputeLocationPosition] loc={loc} toastHeight={toastHeight} area.Top={area.Top} area.Bottom={area.Bottom}")

			Select Case loc
				Case ToastLocation.TopLeft
					x = area.Left + margin
					y = area.Top + margin

				Case ToastLocation.TopCenter
					x = area.Left + (area.Width - toastWidth) \ 2
					y = area.Top + margin

				Case ToastLocation.TopRight
					x = area.Right - toastWidth - margin
					y = area.Top + margin

				Case ToastLocation.MiddleLeft
					x = area.Left + margin
					y = area.Top + (area.Height - toastHeight) \ 2

				Case ToastLocation.MiddleCenter
					x = area.Left + (area.Width - toastWidth) \ 2
					y = area.Top + (area.Height - toastHeight) \ 2

				Case ToastLocation.MiddleRight
					x = area.Right - toastWidth - margin
					y = area.Top + (area.Height - toastHeight) \ 2

				Case ToastLocation.BottomLeft
					x = area.Left + margin
					y = area.Bottom - toastHeight - margin

				Case ToastLocation.BottomCenter
					x = area.Left + (area.Width - toastWidth) \ 2
					y = area.Bottom - toastHeight - margin

				Case ToastLocation.BottomRight
					x = area.Right - toastWidth - margin
					y = area.Bottom - toastHeight - margin
			End Select

			Debug.WriteLine($"[ComputeLocationPosition] result=({x},{y})")



			'-----------------------------------------
			' ⭐ Clamp Y so toast never goes off-screen
			'-----------------------------------------
			If y < area.Top Then
				y = area.Top
			End If

			If y + toastHeight > area.Bottom Then
				y = area.Bottom - toastHeight
			End If

			Debug.WriteLine($"[ComputeLocationPosition] clamped=({x},{y})")


			Return New Point(x, y)
		End Function

	End Class

End Namespace
