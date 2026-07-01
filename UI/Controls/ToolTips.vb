
Imports System.ComponentModel

Namespace Skye.UI

	''' <summary>
	''' ToolTip control with custom colors, fonts and image support.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	<ProvideProperty("ToolTipImage", GetType(Control))>
	Public Class ToolTip
		Inherits System.Windows.Forms.ToolTip

		'Declarations
		Public Enum ImageAlignmentType
			Left
			Right
		End Enum
		Private _font As New Font("Segoe UI", 10, FontStyle.Regular)
		Private _foreColor As Color = Color.Black
		Private _backColor As Color = Color.WhiteSmoke
		Private _borderColor As Color = Color.White
		Private _gradientBackColor As Boolean = False
		Private _gradientStartColor As Color = Color.White
		Private _gradientEndColor As Color = Color.LightGray
		Private _imageAlignment As ImageAlignmentType = ImageAlignmentType.Left
		Private _autoPopDelay As Integer = 5000
		Private _initialDelay As Integer = 1000
		Private _reShowDelay As Integer = 1000
		Private ReadOnly _imageMap As New Dictionary(Of Control, Image)
		Private ReadOnly padding As Integer = 7

		'Properties
		''' <summary>
		''' Specifies the Font to be used for the ToolTip.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Specifies the Font to be used for the ToolTip."), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), DefaultValue(GetType(Font), "Segoe UI, 10, FontStyle.Regular")>
		Public Shadows Property Font As Font
			Get
				Return _font
			End Get
			Set(value As Font)
				_font = value
			End Set
		End Property
		''' <summary>
		''' Resets the Font property to its default value.
		''' </summary>
		Public Sub ResetFont()
			Font = New Font("Segoe UI", 10, FontStyle.Regular)
		End Sub
		''' <summary>
		''' Specifies whether the Font property should be serialized by designer.
		''' </summary>
		''' <returns></returns>
		Public Function ShouldSerializeFont() As Boolean
			Return Not Font.Equals(New Font("Segoe UI", 10, FontStyle.Regular))
		End Function
		''' <summary>
		''' Specifies the Text Color for the ToolTip.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Specifies the Text Color for the ToolTip."), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), DefaultValue(GetType(Color), "Black")>
		Public Shadows Property ForeColor As Color
			Get
				Return _foreColor
			End Get
			Set(value As Color)
				_foreColor = value
			End Set
		End Property
		''' <summary>
		''' Specifies the Background Color of the ToolTip.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Specifies the Background Color of the ToolTip."), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), DefaultValue(GetType(Color), "WhiteSmoke")>
		Public Shadows Property BackColor As Color
			Get
				Return _backColor
			End Get
			Set(value As Color)
				_backColor = value
			End Set
		End Property
		''' <summary>
		''' Specifies the Border Color of the ToolTip.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Specifies the Border Color of the ToolTip."), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), DefaultValue(GetType(Color), "White")>
		Public Property BorderColor As Color
			Get
				Return _borderColor
			End Get
			Set(value As Color)
				_borderColor = value
			End Set
		End Property
		''' <summary>
		''' Enable horizontal gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Enable horizontal gradient background."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientBackColor As Boolean
			Get
				Return _gradientBackColor
			End Get
			Set(value As Boolean)
				_gradientBackColor = value
			End Set
		End Property
		''' <summary>
		''' Specifies the start color for the gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Start color for the gradient background."), DefaultValue(GetType(Color), "White"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientStartColor As Color
			Get
				Return _gradientStartColor
			End Get
			Set(value As Color)
				_gradientStartColor = value
			End Set
		End Property
		''' <summary>
		''' Specifies the end color for the gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("End color for the gradient background."), DefaultValue(GetType(Color), "LightGray"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientEndColor As Color
			Get
				Return _gradientEndColor
			End Get
			Set(value As Color)
				_gradientEndColor = value
			End Set
		End Property
		''' <summary>
		''' Specifies whether the image is aligned to the left or right.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Specifies whether the image Is aligned to the left Or right."), DefaultValue(ImageAlignmentType.Left), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ImageAlignment As ImageAlignmentType
			Get
				Return _imageAlignment
			End Get
			Set(value As ImageAlignmentType)
				_imageAlignment = value
			End Set
		End Property
		''' <summary>
		''' Specifies the time (in milliseconds) the tooltip remains visible.
		''' </summary>
		''' <returns></returns>
		<Category("Behavior"), Description("Specifies the time (in milliseconds) the tooltip remains visible."), DefaultValue(5000)>
		Public Shadows Property AutoPopDelay As Integer
			Get
				Return _autoPopDelay
			End Get
			Set(value As Integer)
				_autoPopDelay = value
				MyBase.AutoPopDelay = value
			End Set
		End Property
		''' <summary>
		''' Specifies the time (in milliseconds) before the tooltip appears.
		''' </summary>
		''' <returns></returns>
		<Category("Behavior"), Description("Specifies the time (in milliseconds) before the tooltip appears."), DefaultValue(1000)>
		Public Shadows Property InitialDelay As Integer
			Get
				Return _initialDelay
			End Get
			Set(value As Integer)
				_initialDelay = value
				MyBase.InitialDelay = value
			End Set
		End Property
		''' <summary>
		''' Specifies the time (in milliseconds) before another tooltip appears.
		''' </summary>
		''' <returns></returns>
		<Category("Behavior"), Description("Specifies the time (in milliseconds) before another tooltip appears."), DefaultValue(1000)>
		Public Shadows Property ReshowDelay As Integer
			Get
				Return _reShowDelay
			End Get
			Set(value As Integer)
				_reShowDelay = value
				MyBase.ReshowDelay = value
			End Set
		End Property
		''' <summary>
		''' Retrieves the Image associated with the specified Control.
		''' </summary>
		''' <param name="ctrl">The Control being specified.</param>
		''' <returns>An Image for the specified control.</returns>
		<Category("Misc"), Description("Specifies the Image to display in the ToolTip for this Control.")>
		Public Function GetToolTipImage(ctrl As Control) As Image
			Dim result As Image = Nothing
			_imageMap.TryGetValue(ctrl, result)
			Return result
		End Function

		''' <summary>
		''' Associates an Image with the specified Control.
		''' </summary>
		''' <param name="ctrl">The Control being specified.</param>
		''' <param name="value">The Image to display in the tooltip for this control</param>
		Public Sub SetToolTipImage(ctrl As Control, value As Image)
			If value Is Nothing Then
				_imageMap.Remove(ctrl)
			Else
				_imageMap(ctrl) = value
			End If
		End Sub
		''' <summary>
		''' NOT USED
		''' Hides the AutomaticDelay property from the designer property browser.
		''' </summary>
		''' <returns></returns>
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Shadows Property AutomaticDelay As Integer
			Get
				Return MyBase.AutomaticDelay
			End Get
			Set(value As Integer)
				MyBase.AutomaticDelay = value
			End Set
		End Property
		''' <summary>
		''' NOT USED
		''' Hides the IsBalloon property from the designer property browser.
		''' </summary>
		''' <returns></returns>
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Shadows Property IsBalloon As Boolean
			Get
				Return MyBase.IsBalloon
			End Get
			Set(value As Boolean)
				MyBase.IsBalloon = value
			End Set
		End Property
		''' <summary>
		''' NOT USED
		''' Hides the ToolTipIcon property from the designer property browser.
		''' </summary>
		''' <returns></returns>
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Shadows Property ToolTipIcon As ToolTipIcon
			Get
				Return MyBase.ToolTipIcon
			End Get
			Set(value As ToolTipIcon)
				MyBase.ToolTipIcon = value
			End Set
		End Property
		''' <summary>
		''' NOT USED
		''' Hides the ToolTipTitle property from the designer property browser.
		''' </summary>
		''' <returns></returns>
		<Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Shadows Property ToolTipTitle As String
			Get
				Return MyBase.ToolTipTitle
			End Get
			Set(value As String)
				MyBase.ToolTipTitle = value
			End Set
		End Property

		'Events
		Public Sub New()
			MyBase.New()
			Initialize()
		End Sub
		Public Sub New(container As IContainer)
			MyBase.New()
			container.Add(Me)
			Initialize()
		End Sub
		Protected Overrides Sub Dispose(disposing As Boolean)
			If disposing Then
				RemoveHandler Me.Popup, AddressOf OnPopup
				RemoveHandler Me.Draw, AddressOf OnDraw
			End If
			MyBase.Dispose(disposing)
		End Sub
		Private Sub OnPopup(sender As Object, e As PopupEventArgs)
			Static s As Size
			Static image As Image
			s = TextRenderer.MeasureText(GetToolTip(e.AssociatedControl), Font)
			_imageMap.TryGetValue(e.AssociatedControl, image)
			image = ResizeImage(image, s.Height)
			s.Width += padding * 2
			s.Height += padding * 2
			If image IsNot Nothing Then s.Width += image.Width
			e.ToolTipSize = s
			image = Nothing
		End Sub
		Private Sub OnDraw(sender As Object, e As DrawToolTipEventArgs)

			'Declarations
			Dim g As Graphics = e.Graphics
			Static image As Image
			Static borderadjustment As Integer
			Static verticaltextadjustment As Integer
			Select Case Font.Size
				Case Is < 11 : borderadjustment = 0
				Case Else : borderadjustment = 1
			End Select
			Select Case Font.Size
				Case Is < 10 : verticaltextadjustment = 0
				Case 10 To 16 : verticaltextadjustment = 1
				Case Is > 16 : verticaltextadjustment = 2
			End Select

			'Draw Background
			If GradientBackColor AndAlso GradientStartColor <> Nothing AndAlso GradientEndColor <> Nothing Then
				Using brush As New Drawing2D.LinearGradientBrush(e.Bounds, GradientStartColor, GradientEndColor, Drawing2D.LinearGradientMode.Horizontal)
					e.Graphics.FillRectangle(brush, e.Bounds)
				End Using
			Else
				e.Graphics.Clear(BackColor)
			End If

			'Draw Border
			Using p As New Pen(BorderColor, CInt(Font.Size / 4)) 'Scale border thickness with font
				g.DrawRectangle(p, 0, 0, e.Bounds.Width - borderadjustment, e.Bounds.Height - borderadjustment)
			End Using

			'Draw Icon
			_imageMap.TryGetValue(e.AssociatedControl, image)
			image = ResizeImage(image, e.Bounds.Height - padding * 2)
			If image IsNot Nothing Then
				Dim bounds As Rectangle = GetImageBounds(image, ImageAlignment, e.Bounds.Width, CInt(e.Bounds.Height / 2 - image.Height / 2))
				g.DrawImage(image, bounds)
				'g.DrawImage(image, New Rectangle(padding, CInt((e.Bounds.Height / 2 - image.Height / 2)), image.Width, image.Height))
			End If

			'Draw Text
			TextRenderer.DrawText(g, e.ToolTipText, Font, If(image Is Nothing, New Point(padding + CInt(Font.Size / 6 / 2), padding), If(_imageAlignment = ImageAlignmentType.Left, New Point(image.Width + padding + CInt(Font.Size / 6), padding - verticaltextadjustment), New Point(padding, padding - verticaltextadjustment))), ForeColor)

			'Finalize
			image = Nothing
			g.Dispose()

		End Sub

		'Procedures
		Private Sub Initialize()
			MyBase.AutoPopDelay = _autoPopDelay
			MyBase.InitialDelay = _initialDelay
			MyBase.ReshowDelay = _reShowDelay
			OwnerDraw = True
			AddHandler Me.Popup, AddressOf OnPopup
			AddHandler Me.Draw, AddressOf OnDraw
		End Sub
		''' <summary>
		''' Resizes an image to fit within the specified maximum height while maintaining aspect ratio.
		''' </summary>
		''' <param name="original">The original Image</param>
		''' <param name="maxHeight">The maximum height the Image can be.</param>
		''' <returns>An Image, either the original unmodified, or a new Image scaled appropriately.</returns>
		Private Shared Function ResizeImage(original As Image, maxHeight As Integer) As Image
			If original Is Nothing Then Return Nothing

			'Minimum size
			If maxHeight < 16 Then maxHeight = 16

			'If the image is already small enough, return as-is
			If original.Height <= maxHeight AndAlso original.Width <= maxHeight Then
				Return original
			End If

			'Calculate scale factor to fit within maxHeight
			Dim scale As Single = CSng(maxHeight / Math.Max(original.Width, original.Height))
			Dim newWidth As Integer = CInt(original.Width * scale)
			Dim newHeight As Integer = CInt(original.Height * scale)

			'Create resized image
			Dim resized As New Bitmap(newWidth, newHeight)
			Using g As Graphics = Graphics.FromImage(resized)
				g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
				g.DrawImage(original, New Rectangle(0, 0, newWidth, newHeight))
			End Using

			Return resized
		End Function
		''' <summary>
		''' Calculates the bounds for the image based on alignment and container size.
		''' </summary>
		''' <param name="img">The Image to be drawn.</param>
		''' <param name="alignment">Specifies ImageAlignmentType.</param>
		''' <param name="containerWidth">The width of the ToolTip</param>
		''' <param name="y">Vertical starting position.</param>
		''' <returns>A Rectangle the defines where to draw the Image.</returns>
		Private Function GetImageBounds(img As Image, alignment As ImageAlignmentType, containerWidth As Integer, y As Integer) As Rectangle
			Dim x As Integer
			Select Case alignment
				Case ImageAlignmentType.Left
					x = padding
				Case ImageAlignmentType.Right
					x = containerWidth - img.Width - padding
			End Select
			Return New Rectangle(x, y, img.Width, img.Height)
		End Function

	End Class

	''' <summary>
	''' Extended ToolTip control with custom colors, fonts and image support, and many more features. Does not inherit from the standard ToolTip control.
	''' </summary>
	<ToolboxItem(True)>
	<ProvideProperty("Text", GetType(Control))>
	<ProvideProperty("Image", GetType(Control))>
	Public Class ToolTipEX
		Inherits Component
		Implements IExtenderProvider

		'Declarations
		''' <summary>
		''' Specifies whether the image is aligned to the left or right of the text.
		''' </summary>
		Public Enum ImageAlignments
			Left
			Right
		End Enum
		''' <summary>
		''' Specifies whether the ToolTip is shown at the cursor, or relative to the control.
		''' </summary>
		Public Enum TooltipPositions
			Control
			Cursor
		End Enum
		Private Class TooltipRequest
			Public Property ID As Guid = New Guid()
			Public Property TargetControl As Control
			Public Property Text As String
			Public Property Image As Image
			Public Property ManualPosition As Point?
			Public Property ForceControlPosition As Boolean
			Public Sub New(targetControl As Control, text As String, image As Image, manualPosition As Point?, forceControlPosition As Boolean)
				Me.ID = Guid.NewGuid()
				Me.TargetControl = targetControl
				Me.Text = text
				Me.Image = image
				Me.ManualPosition = manualPosition
				Me.ForceControlPosition = forceControlPosition
			End Sub
		End Class
		Private popup As ToolTipPopup 'The actual popup form
		Private popupHandle As IntPtr 'Handle of the popup form, used to force handle creation
		Private ReadOnly tooltips As New Dictionary(Of Control, String) 'Map of controls to tooltip text
		Private ReadOnly tooltipImages As New Dictionary(Of Control, Image) 'Map of controls to tooltip images
		Private ShowDelayTimer As Timer 'Timer used to delay showing the tooltip
		Private WithEvents HideDelayTimer As New Timer 'Timer used to delay hiding the tooltip
		Private _hoveredcontrol As Control 'The control currently being hovered over
		Private _shadowalpha As Integer = 80 'Used only for get/set of ShadowAlpha property to control range.
		Private _fadeinrate As Integer = 50 'Used only for get/set of FadeInRate property to control range.
		Private _fadeoutrate As Integer = 50 'Used only for get/set of FadeOutRate property to control range.
		Private _manualTooltipActive As Boolean = False 'Used to temporarily disable mouse enter events when showing/hiding tooltips programmatically.

		'Properties
		''' <summary>
		''' Returns whether or not the tooltip is currently visible.
		''' </summary>
		''' <returns>True if the tooltip is currently visible, otherwise false.</returns>
		<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public ReadOnly Property IsVisible As Boolean
			Get
				Return popup IsNot Nothing AndAlso popup.Visible
			End Get
		End Property

		'Events
		Public Sub New()
			MyBase.New
			Initialize()
		End Sub
		Public Sub New(container As IContainer)
			MyBase.New
			container.Add(Me)
			Initialize()
		End Sub
		Protected Overrides Sub Dispose(disposing As Boolean)
			If disposing Then
				ShowDelayTimer?.Dispose()
				ShowDelayTimer = Nothing
				HideDelayTimer?.Dispose()
				HideDelayTimer = Nothing
				If popup IsNot Nothing Then
					popup.Dispose()
					popup = Nothing
				End If
				_hoveredcontrol = Nothing
				tooltips.Clear()
				tooltipImages.Clear()
			End If
			MyBase.Dispose(disposing)
		End Sub
		Private Sub OnMouseEnter(sender As Object, e As EventArgs)
			If _manualTooltipActive Then Exit Sub

			Dim ctrl As Control = CType(sender, Control)

			'Moving to a different control → hide current immediately
			If _hoveredcontrol IsNot Nothing AndAlso _hoveredcontrol IsNot ctrl AndAlso IsVisible Then
				HideDelayTimer.Stop()
				HideTooltip()
			End If

			_hoveredcontrol = ctrl

			'Start (or restart) the show delay
			If ShowDelayTimer Is Nothing Then
				ShowDelayTimer = New Timer()
				AddHandler ShowDelayTimer.Tick, AddressOf ShowTooltipDelayed
			Else
				ShowDelayTimer.Stop()
			End If

			ShowDelayTimer.Interval = If(ShowDelay < 1, 1, ShowDelay)
			ShowDelayTimer.Start()
		End Sub
		Private Sub OnMouseLeave(sender As Object, e As EventArgs)
			If _manualTooltipActive Then Exit Sub

			'Cancel a pending show for the control we’re leaving
			ShowDelayTimer?.Stop()

			'If we’re leaving the control that owns the current tooltip, start hide delay
			If _hoveredcontrol Is CType(sender, Control) Then StartHideDelayTimer()
		End Sub
		Private Sub OnMouseHover(sender As Object, e As EventArgs) 'Optional: keep this lightweight so it doesn’t control show timing anymore
			'Intentionally do nothing — enter drives show timing now
		End Sub
		Private Sub HideDelayTimer_Tick(sender As Object, e As EventArgs) Handles HideDelayTimer.Tick
			HideDelayTimer.Stop()
			HideTooltip()
			_hoveredcontrol = Nothing
			_manualTooltipActive = False
		End Sub

		'Procedures
		Private Sub Initialize()
			popup = New ToolTipPopup(Me)
			popupHandle = popup.Handle 'Force handle creation AFTER popup is initialized, required so showing the form does not steal focus.
		End Sub
		Private Sub ShowTooltipDelayed(sender As Object, e As EventArgs)
			ShowDelayTimer?.Stop()
			If _hoveredcontrol Is Nothing Then Exit Sub
			If Not tooltips.ContainsKey(_hoveredcontrol) Then Exit Sub

			Dim text As String = tooltips(_hoveredcontrol)
			If String.IsNullOrWhiteSpace(text) Then Exit Sub
			Dim image As Image = Nothing
			tooltipImages.TryGetValue(_hoveredcontrol, image)
			'If tooltipImages.ContainsKey(_hoveredcontrol) Then
			'	image = tooltipImages(_hoveredcontrol)
			'End If

			Dim request As New TooltipRequest(_hoveredcontrol, text, image, Nothing, False)
			ShowTooltip(request)
		End Sub
		''' <summary>
		''' Show a tooltip for the specified control using the text and image set via the extender properties.
		''' </summary>
		''' <param name="targetControl">The control to show the tooltip for.</param>
		Public Sub ShowTooltip(targetControl As Control)
			If targetControl Is Nothing OrElse Not tooltips.ContainsKey(targetControl) Then Exit Sub
			If String.IsNullOrWhiteSpace(tooltips(targetControl)) Then Exit Sub
			ShowDelayTimer?.Stop()
			Dim img As Image = Nothing
			tooltipImages.TryGetValue(targetControl, img)
			Dim request As New TooltipRequest(targetControl, tooltips(targetControl), img, Nothing, True)
			ShowTooltip(request)
			_manualTooltipActive = True
			StartHideDelayTimer()
		End Sub
		''' <summary>
		''' Show a tooltip for the specified control with the specified text and optional image, ignoring any text/image set via the extender properties.
		''' </summary>
		''' <param name="targetControl">The control to show the tooltip for.</param>
		''' <param name="text">The text to show.</param>
		''' <param name="image">Optionally, the image to show.</param>
		Public Sub ShowTooltip(targetControl As Control, text As String, Optional image As Image = Nothing)
			If targetControl Is Nothing Then Exit Sub
			If String.IsNullOrWhiteSpace(text) Then Exit Sub
			ShowDelayTimer?.Stop()
			Dim request As New TooltipRequest(targetControl, text, image, Nothing, True)
			ShowTooltip(request)
			_manualTooltipActive = True
			StartHideDelayTimer()
		End Sub
		''' <summary>
		''' Show a tooltip at the specified screen position with the specified text and optional image.
		''' </summary>
		''' <param name="position">The screen coordinates to show the tooltip</param>
		''' <param name="text">The text to show.</param>
		''' <param name="image">Optionally, the image to show.</param>
		Public Sub ShowTooltipAt(position As Point, text As String, Optional image As Image = Nothing)
			ShowDelayTimer?.Stop()
			Dim request As New TooltipRequest(Nothing, text, image, position, False)
			ShowTooltip(request)
			_manualTooltipActive = True
			StartHideDelayTimer()
		End Sub
		''' <summary>
		''' Show a tooltip at the current cursor position with the specified text and optional image.
		''' </summary>
		''' <param name="text">The text to show.</param>
		''' <param name="image">Optionally, the image to show.</param>
		Public Sub ShowTooltipAtCursor(text As String, Optional image As Image = Nothing)
			ShowDelayTimer?.Stop()
			Dim request As New TooltipRequest(Nothing, text, image, Nothing, False)
			ShowTooltip(request)
			_manualTooltipActive = True
			StartHideDelayTimer(True)
		End Sub
		Private Sub ShowTooltip(request As TooltipRequest)
			'Trace.WriteLine($"[ToolTipEX] ShowTooltip request for '{If(request.Text, "(null text)")}' " & $"on {If(request.TargetControl?.Name, "(no control)")}")
			'Common.Log.Write($"ToolTipEX SHOW → '{request.Text}' on {request.TargetControl?.Name}")

			'Ensure popup is valid
			If popup Is Nothing OrElse popup.IsDisposed OrElse Not popup.IsHandleCreated Then
				popup = New ToolTipPopup(Me)
				popupHandle = popup.Handle 'force handle creation
			End If

			HideDelayTimer?.Stop()
			HideTooltip()

			'Measure Text Size
			Dim textSize As Size = TextRenderer.MeasureText(request.Text, Me.Font, New Size(0, 0), TextFormatFlags.NoPrefix)
			Dim totalSize As New Size(textSize.Width + TextPadding * 2, textSize.Height + TextPadding * 2)
			request.Image = ResizeImage(request.Image, textSize.Height)
			If request.Image IsNot Nothing Then totalSize.Width += request.Image.Width + TextPadding

			'Set Location (& adjust to avoid clipping)
			Dim pos As Point
			If request.TargetControl IsNot Nothing Then
				If TooltipPosition = TooltipPositions.Control OrElse request.ForceControlPosition Then
					pos = request.TargetControl.PointToScreen(New Point(request.TargetControl.Width \ 2, request.TargetControl.Height))
				Else
					pos = GetCursorPositionWithOffset()
				End If
			ElseIf request.ManualPosition.HasValue Then
				pos = request.ManualPosition.Value
			Else
				pos = GetCursorPositionWithOffset()
			End If
			If pos.X + totalSize.Width > Screen.GetWorkingArea(pos).Right Then pos.X = Screen.GetWorkingArea(pos).Right - totalSize.Width - 2
			If pos.Y + totalSize.Height > Screen.GetWorkingArea(pos).Bottom Then pos.Y = Screen.GetWorkingArea(pos).Bottom - totalSize.Height - 2

			popup.Size = totalSize
			popup.TooltipText = request.Text
			popup.TooltipImage = request.Image
			popup.ShowTooltip(pos.X, pos.Y)
			WinAPI.SetWindowPos(popup.Handle, WinAPI.HWND_TOPMOST, 0, 0, 0, 0, WinAPI.SWP_NOMOVE Or WinAPI.SWP_NOSIZE Or WinAPI.SWP_NOACTIVATE)
		End Sub
		''' <summary>
		''' Hides the tooltip if it is currently visible.
		''' </summary>
		Public Sub HideTooltip()
			'Trace.WriteLine($"[ToolTipEX] HideTooltip called, IsVisible={IsVisible}")
			'Common.Log.Write("ToolTipEX HIDE")
			If IsVisible Then popup.HideTooltip()
		End Sub
		Private Sub StartHideDelayTimer(Optional PlusFadeInRate As Boolean = False)
			HideDelayTimer?.Stop()
			If HideDelay > 0 Then
				HideDelayTimer.Interval = HideDelay + If(PlusFadeInRate, FadeInRate * 10, 0) '*10 adjusts for the fact that the fade timer ticks ten times
				HideDelayTimer.Start()
			End If
		End Sub
		Private Shared Function ResizeImage(original As Image, maxHeight As Integer) As Image
			If original Is Nothing Then Return Nothing

			'Minimum size
			If maxHeight < 16 Then maxHeight = 16

			'If the image is already small enough, return as-is
			If original.Height <= maxHeight AndAlso original.Width <= maxHeight Then
				Return original
			End If

			'Calculate scale factor to fit within maxHeight
			Dim scale As Single = CSng(maxHeight / Math.Max(original.Width, original.Height))
			Dim newWidth As Integer = CInt(original.Width * scale)
			Dim newHeight As Integer = CInt(original.Height * scale)

			'Create resized image
			Dim resized As New Bitmap(newWidth, newHeight)
			Using g As Graphics = Graphics.FromImage(resized)
				g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
				g.DrawImage(original, New Rectangle(0, 0, newWidth, newHeight))
			End Using

			Return resized
		End Function
		Private Shared Function GetCursorPositionWithOffset() As Point
			Dim pos As Point
			pos = Cursor.Position
			pos = New Point(pos.X + 16, pos.Y + 32) 'Offset for cursor size
			Return pos
		End Function

		'Desginer Properties
		''' <summary>
		''' Gets or sets the font of the tooltip.
		''' </summary>
		''' <value>The <see cref="System.Drawing.Font"/> used for the tooltip background.</value>

		<Category("Appearance"), Description("The font used for tooltip text."), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Font As Font = New Font("Segoe UI", 10, FontStyle.Regular)
		''' <summary>
		''' Gets or sets the background color of the tooltip.
		''' The transparency key of the tooltip is set to SystemColors.Control(R240,G240,B240),
		''' so setting the BackColor to this value will make the tooltip background invisible.
		''' </summary>
		''' <value>The <see cref="Color"/> used for the tooltip background.</value>
		<Category("Appearance"), Description("The background color of the tooltip."), DefaultValue(GetType(Color), "WhiteSmoke"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property BackColor As Color = Color.WhiteSmoke
		''' <summary>
		''' Specifies whether to enable a horizontal gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Enable horizontal gradient background."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientBackColor As Boolean = False
		''' <summary>
		''' Specifies the start color for the gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("Start color for the gradient background."), DefaultValue(GetType(Color), "White"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientStartColor As Color = Color.White
		''' <summary>
		''' Specifies the end color for the gradient background.
		''' </summary>
		''' <returns></returns>
		<Category("Appearance"), Description("End color for the gradient background."), DefaultValue(GetType(Color), "LightGray"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property GradientEndColor As Color = Color.LightGray
		''' <summary>
		''' Gets or sets the text color of the tooltip.
		''' </summary>
		''' <value>The <see cref="Color"/> used for the tooltip text.</value>
		<Category("Appearance"), Description("The text color of the tooltip."), DefaultValue(GetType(Color), "Black"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ForeColor As Color = Color.Black
		''' <summary>
		''' Gets or sets the border color of the tooltip.
		''' </summary>
		''' <value>The <see cref="Color"/> used for the tooltip border.</value>
		<Category("Appearance"), Description("The border color of the tooltip."), DefaultValue(GetType(Color), "White"), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property BorderColor As Color = Color.White
		''' <summary>
		''' Gets or sets a value indicating whether or not to show the tooltip border.
		''' </summary>
		''' <value>A <see cref="Boolean"/></value>
		<Category("Appearance"), Description("Whether or not to show a border."), DefaultValue(True), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ShowBorder As Boolean = True
		''' <summary>
		''' Gets or sets a value indicating the thickness of the border.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the border thickness in pixels.</value>
		<Category("Appearance"), Description("Thickness of the border. 0 = Auto."), DefaultValue(0), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property BorderThickness As Integer = 0
		''' <summary>
		''' Gets or sets a value indicating the padding around the tooltip text.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the padding around the tooltip text.</value>
		<Category("Layout"), Description("Padding around the tooltip text."), DefaultValue(8), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property TextPadding As Integer = 8
		''' <summary>
		''' Gets or sets a value indicating the thickness of the tooltip shadow.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the shadow thickness in pixels.</value>
		<Category("Appearance"), Description("Thickness of the tooltip shadow."), DefaultValue(2), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ShadowThickness As Integer = 2
		''' <summary>
		''' Gets or sets a value indicating the alpha value of the tooltip shadow color.
		''' </summary>
		''' <returns>An <see cref="Integer"/> representing the Alpha value of the tooltip's shadow.</returns>
		<Category("Appearance"), Description("Alpha Value of the tooltip's Shadow RGB Color, 0 - 255, 0 = Transparent"), DefaultValue(80), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ShadowAlpha As Integer
			Get
				Return _shadowalpha
			End Get
			Set(value As Integer)
				_shadowalpha = Math.Max(0, Math.Min(255, value))
			End Set
		End Property
		''' <summary>
		''' Gets or sets a value for the fade-in rate of the tooltip.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the fade-in rate in milliseconds. 0 means show immediately.</value>
		<Category("Behavior"), Description("Fade In Rate of the ToolTip, 0 - 200, 0 = Immediate"), DefaultValue(50), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property FadeInRate As Integer
			Get
				Return _fadeinrate
			End Get
			Set(value As Integer)
				_fadeinrate = Math.Max(0, Math.Min(200, value))
			End Set
		End Property
		''' <summary>
		''' Gets or sets a value for the fade-out rate of the tooltip.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the fade-out rate in milliseconds. 0 means hide immediately.</value>
		<Category("Behavior"), Description("Fade Out Rate of the ToolTip, 0 - 200, 0 = Immediate"), DefaultValue(50), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property FadeOutRate As Integer
			Get
				Return _fadeoutrate
			End Get
			Set(value As Integer)
				_fadeoutrate = Math.Max(0, Math.Min(200, value))
			End Set
		End Property
		''' <summary>
		''' Gets or sets a value for the show delay of the tooltip.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the milliseconds before the tooltip is shown. 0 means show immediately on hover.</value>
		<Category("Behavior"), Description("The milliseconds before the tooltip is shown."), DefaultValue(500), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ShowDelay As Integer = 500
		''' <summary>
		''' Gets or sets a value for the hide delay of the tooltip.
		''' </summary>
		''' <value>An <see cref="Integer"/> representing the milliseconds before the tooltip is hidden. 0 = don't hide until hidden intentionally.</value>
		<Category("Behavior"), Description("The milliseconds before the tooltip is hidden. 0 = disable auto-hide"), DefaultValue(500), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property HideDelay As Integer = 500
		''' <summary>
		''' Gets or sets a value indicating how the image is to be aligned relative to the text.
		''' </summary>
		''' <value>An <see cref="ImageAlignments"/> representing the desired image alignment.</value>
		<Category("Layout"), Description("Specifies whether the image is aligned to the left or right."), DefaultValue(ImageAlignments.Left), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property ImageAlignment As ImageAlignments = ImageAlignments.Left
		''' <summary>
		''' Gets or sets a value indicating where the tooltip should be automatically located on hover.
		''' </summary>
		''' <value>An <see cref="TooltipPositions"/> representing the desired location for the tooltip.</value>
		<Category("Layout"), Description("Specifies whether the ToolTip is shown at the cursor, or relative to the control."), DefaultValue(TooltipPositions.Cursor), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property TooltipPosition As TooltipPositions = TooltipPositions.Cursor
		''' <summary>
		''' Gets or sets a value indicating whether to copy the tooltip text when right-clicking the tooltip itself.
		''' </summary>
		''' <value>An <see cref="Boolean"/> specifying whether to copy or not.</value>
		<Category("Behavior"), Description("Whether or not to copy the tooltip text when the user right-clicks the tooltip."), DefaultValue(False), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property CopyOnRightClick As Boolean = False

		'Designer Functions
		Public Function CanExtend(extendee As Object) As Boolean Implements IExtenderProvider.CanExtend
			Return TypeOf extendee Is Control
		End Function
		''' <summary>
		''' Gets the Text to display in the ToolTip for the specified control.
		''' </summary>
		''' <param name="ctrl">A reference to a control.</param>
		''' <returns>A string representing the text that will be displayed in the tooltip for this control.</returns>
		<Editor(GetType(System.ComponentModel.Design.MultilineStringEditor), GetType(System.Drawing.Design.UITypeEditor))>
		<Category("ToolTipEX"), Description("Specifies the Text to display in the ToolTip for this control."), Browsable(True), DefaultValue("")>
		Public Function GetText(ctrl As Control) As String
			Dim result As String = String.Empty
			tooltips.TryGetValue(ctrl, result)
			Return result
		End Function

		''' <summary>
		''' Sets the Text to display in the ToolTip for the specified control.
		''' </summary>
		''' <param name="ctrl">A reference to a control.</param>
		''' <param name="value">A String representing the text to be displayed in the tooltip for this control.</param>
		Public Sub SetText(ctrl As Control, value As String)
			Dim isNew As Boolean = Not tooltips.ContainsKey(ctrl)
			tooltips(ctrl) = value
			If isNew AndAlso Not String.IsNullOrWhiteSpace(value) Then
				AddHandler ctrl.MouseEnter, AddressOf OnMouseEnter
				AddHandler ctrl.MouseHover, AddressOf OnMouseHover
				AddHandler ctrl.MouseLeave, AddressOf OnMouseLeave
			End If
		End Sub

		''' <summary>
		''' Gets the Image to display in the ToolTip for the specified control.
		''' </summary>
		''' <param name="ctrl">A reference to a control.</param>
		''' <returns>The Image that will be displayed in the tooltip for this control.</returns>
		<Category("ToolTipEX"), Description("Specifies the Image to display in the ToolTip for this Control."), Browsable(True)>
		Public Function GetImage(ctrl As Control) As Image
			Dim result As Image = Nothing
			tooltipImages.TryGetValue(ctrl, result)
			Return result
		End Function

		''' <summary>
		''' Sets the Image to display in the ToolTip for the specified control.
		''' </summary>
		''' <param name="ctrl">A reference to a control.</param>
		''' <param name="value">An Image to be displayed in the tooltip for this control.</param>
		Public Sub SetImage(ctrl As Control, value As Image)
			If value Is Nothing Then
				tooltipImages.Remove(ctrl)
			Else
				tooltipImages(ctrl) = value
			End If
		End Sub

		Private Class ToolTipPopup
			Inherits Form

			'Declarations
			Private _owner As ToolTipEX
			Public TooltipText As String = String.Empty
			Public TooltipImage As Image = Nothing
			Private FadeInTimer As Timer
			Private FadeOutTimer As Timer

			'Events
			Protected Overrides Sub WndProc(ByRef m As Message)
				Select Case m.Msg
					Case WinAPI.WM_MOUSEACTIVATE
						m.Result = CType(WinAPI.MA_NOACTIVATE, IntPtr)
						Exit Sub
					Case WinAPI.WM_RBUTTONUP
						If _owner.CopyOnRightClick AndAlso Not String.IsNullOrWhiteSpace(TooltipText) Then
							Try
								Clipboard.SetText(TooltipText)
							Catch ex As Exception 'Ignore clipboard errors
							End Try
						End If
						HideTooltip()
						'Case WinAPI.WM_THEMECHANGED Or WinAPI.WM_SYSCOLORCHANGE
						'	''Force Reinitialization
						'	'Trace.WriteLine($"[ToolTipPopup] {m.Msg} received — calling RecreateHandle at {DateTime.Now:HH:mm:ss.fff}")
						'	'If Not Me.IsDisposed Then Me.RecreateHandle()
						'	'Trace.WriteLine($"[ToolTipPopup] {m.Msg} received — refreshing visuals")
						'	Me.Invalidate() 'forces repaint with new theme colors
					Case WinAPI.WM_THEMECHANGED, WinAPI.WM_SYSCOLORCHANGE, WinAPI.WM_DPICHANGED
						'Force handle recreation if still alive
						If Not Me.IsDisposed Then
							Me.RecreateHandle()
						End If
						'Optional: also refresh visuals
						Me.Invalidate()
				End Select
				MyBase.WndProc(m)
			End Sub
			Public Sub New(owner As ToolTipEX)
				_owner = owner
				Initialize()
			End Sub
			Protected Overrides ReadOnly Property CreateParams As CreateParams
				Get
					Dim cp As CreateParams = MyBase.CreateParams
					cp.Style = WinAPI.WS_POPUP
					cp.ExStyle = cp.ExStyle Or WinAPI.WS_EX_TOOLWINDOW Or WinAPI.WS_EX_TOPMOST Or WinAPI.WS_EX_NOACTIVATE 'Or WinAPI.WS_EX_TRANSPARENT
					Return cp
				End Get
			End Property
			Protected Overrides Sub OnHandleCreated(e As EventArgs)
				MyBase.OnHandleCreated(e)
				'Trace.WriteLine($"[ToolTipPopup] Handle created")
				'Reapply any visual/theme settings that depend on the handle
				Initialize()
			End Sub
			Protected Overrides Sub Dispose(disposing As Boolean)
				If disposing Then
					_owner = Nothing
					FadeInTimer?.Dispose()
					FadeInTimer = Nothing
					FadeOutTimer?.Dispose()
					FadeOutTimer = Nothing
				End If
				MyBase.Dispose(disposing)
			End Sub
			Protected Overrides Sub OnPaint(e As PaintEventArgs)
				Dim g As Graphics = e.Graphics
				g.SmoothingMode = Drawing2D.SmoothingMode.None
				g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

				'Draw Shadow
				If _owner.ShadowThickness > 0 Then
					Dim shadowRect As New Rectangle(_owner.ShadowThickness * 2, _owner.ShadowThickness * 2, Me.Width, Me.Height)
					Using shadowBrush As New SolidBrush(Color.FromArgb(_owner.ShadowAlpha, Color.Black)) '60
						g.FillRectangle(shadowBrush, shadowRect)
					End Using
				End If

				'Draw Background
				Dim rect As New Rectangle(0, 0, Me.Width - _owner.ShadowThickness, Me.Height - _owner.ShadowThickness)
				If _owner.GradientBackColor AndAlso _owner.GradientStartColor <> Nothing AndAlso _owner.GradientEndColor <> Nothing Then
					Using bgBrush As New Drawing2D.LinearGradientBrush(rect, _owner.GradientStartColor, _owner.GradientEndColor, Drawing2D.LinearGradientMode.Horizontal)
						g.FillRectangle(bgBrush, rect)
					End Using
				Else
					Using bgBrush As New SolidBrush(_owner.BackColor)
						g.FillRectangle(bgBrush, rect)
					End Using
				End If

				'Draw Border
				If _owner.ShowBorder Then
					Dim thickness As Integer
					If _owner.BorderThickness = 0 Then
						thickness = CInt(_owner.Font.Size / 4) 'Scale border thickness with font
					Else
						thickness = _owner.BorderThickness
					End If
					Using p As New Pen(_owner.BorderColor, thickness)
						p.Alignment = Drawing2D.PenAlignment.Inset
						g.DrawRectangle(p, 0, 0, Me.Width - _owner.ShadowThickness, Me.Height - _owner.ShadowThickness)
					End Using
				End If


				'Draw Text
				Dim textRect As Rectangle
				If TooltipImage Is Nothing Then
					textRect = New Rectangle(_owner.TextPadding, _owner.TextPadding - _owner.ShadowThickness, Me.Width - _owner.TextPadding, Me.Height - _owner.TextPadding)
				Else
					Select Case _owner.ImageAlignment
						Case ImageAlignments.Left
							textRect = New Rectangle(TooltipImage.Width + _owner.TextPadding * 2, _owner.TextPadding - _owner.ShadowThickness, Me.Width - _owner.TextPadding, Me.Height - _owner.TextPadding)
						Case ImageAlignments.Right
							textRect = New Rectangle(_owner.TextPadding, _owner.TextPadding - _owner.ShadowThickness, Me.Width - TooltipImage.Width - _owner.TextPadding * 2, Me.Height - _owner.TextPadding)
					End Select
				End If
				' See if multiline text for image layout
				Dim measured As Size = TextRenderer.MeasureText(TooltipText, _owner.Font, New Size(textRect.Width, Integer.MaxValue), TextFormatFlags.WordBreak Or TextFormatFlags.NoPrefix)
				Dim isMultiLine As Boolean = measured.Height > _owner.Font.Height
				TextRenderer.DrawText(g, TooltipText, _owner.Font, textRect, _owner.ForeColor, TextFormatFlags.Left Or TextFormatFlags.Top Or TextFormatFlags.WordBreak Or TextFormatFlags.NoPrefix)

				'Draw Image
				If TooltipImage IsNot Nothing Then
					Dim contentHeight As Integer = Me.Height - _owner.ShadowThickness
					Dim iconY As Integer = (contentHeight - TooltipImage.Height) \ 2
					Dim bounds As Rectangle = GetImageBounds(TooltipImage, _owner.ImageAlignment, Me.Width, iconY)
					'Dim bounds As Rectangle = GetImageBounds(TooltipImage, _owner.ImageAlignment, Me.Width, CInt(Me.Height / 2 - TooltipImage.Height / 2) + If(isMultiLine, 0, 1))
					g.DrawImage(TooltipImage, bounds)
				End If

			End Sub
			Private Sub Form_Click(sender As Object, e As EventArgs) Handles MyBase.Click
				HideTooltip()
			End Sub

			'Form Functions
			<System.Diagnostics.CodeAnalysis.SuppressMessage("Performance",
				 "CA1822:Mark members as static",
				 Justification:="Kept as instance for override/API consistency")>
			Protected Shadows Function ShowWithoutActivation() As Boolean
				Return True
			End Function

			'Handlers
			' Class-level handlers
			Private Sub FadeInHandler(sender As Object, e As EventArgs)
				Me.Opacity = Math.Min(1, Me.Opacity + 0.1)
				'Trace.WriteLine($"[ToolTipPopup] FadeIn tick, Opacity={Me.Opacity}")
				If Me.Opacity >= 1 Then
					FadeInTimer.Stop()
				End If
			End Sub
			Private Sub FadeOutHandler(sender As Object, e As EventArgs)
				If Not Me.IsHandleCreated Then
					FadeOutTimer.Stop()
					Return
				End If

				Me.Opacity = Math.Max(0, Me.Opacity - 0.1)
				If Me.Opacity <= 0 Then
					FadeOutTimer.Stop()
					Me.Opacity = 0
					Me.Hide()
				End If
			End Sub

			'Procedures
			Private Sub Initialize()
				Me.FormBorderStyle = FormBorderStyle.None
				Me.ShowInTaskbar = False
				Me.StartPosition = FormStartPosition.Manual
				Me.TransparencyKey = Me.BackColor
				Me.Opacity = 1
			End Sub
			Public Sub ShowTooltip(x As Integer, y As Integer)

				'Trace.WriteLine($"[ToolTipPopup] ShowTooltip at {x},{y}, Visible={Me.Visible}, HandleCreated={Me.IsHandleCreated}, Opacity={Me.Opacity}")

				'Stop any existing animations
				StopTimers()

				'Reset if already visible
				If Me.Visible Then Me.Hide()

				'Ensure handle exists
				If Not Me.IsHandleCreated Then Me.CreateHandle()

				'Position the popup
				Me.Location = New Point(x, y)

				If Me.Visible Then
					'Already visible → just move it
					WinAPI.SetWindowPos(Me.Handle, IntPtr.Zero, x, y, 0, 0, WinAPI.SWP_NOACTIVATE Or WinAPI.SWP_NOSIZE Or WinAPI.SWP_NOZORDER)
				Else
					'Show without stealing focus
					WinAPI.ShowWindow(Me.Handle, WinAPI.SW_SHOWNOACTIVATE)
					WinAPI.SetWindowPos(Me.Handle, WinAPI.HWND_TOPMOST, x, y, Me.Width, Me.Height, WinAPI.SWP_NOACTIVATE Or WinAPI.SWP_SHOWWINDOW Or WinAPI.SWP_NOZORDER)
				End If
				'Trace.WriteLine($"[ToolTipPopup] SetWindowPos applied, TopMost={Me.TopMost}")

				'Fade-in effect
				If _owner.FadeInRate > 0 Then
					Me.Opacity = 0
					FadeInTimer = New Timer() With {.Interval = _owner.FadeInRate}
					AddHandler FadeInTimer.Tick, AddressOf FadeInHandler
					FadeInTimer.Start()
				Else
					Me.Opacity = 1
				End If

			End Sub
			Public Sub HideTooltip()

				'If already invisible, skip redundant work
				If Not Me.Visible AndAlso Me.Opacity <= 0 Then Return

				'Stop any animation in progress
				StopTimers()

				If _owner.FadeOutRate > 0 AndAlso Me.Visible Then
					'Trace.WriteLine($"[ToolTipPopup] HideTooltip starting fade-out, Visible={Me.Visible}, HandleCreated={Me.IsHandleCreated}")
					FadeOutTimer = New Timer() With {.Interval = _owner.FadeOutRate}
					AddHandler FadeOutTimer.Tick, AddressOf FadeOutHandler
					FadeOutTimer.Start()
				Else
					Me.Opacity = 0
					If Me.IsHandleCreated Then Me.Hide()
				End If

			End Sub
			Private Sub StopTimers()
				If FadeInTimer IsNot Nothing Then
					RemoveHandler FadeInTimer.Tick, AddressOf FadeInHandler
					FadeInTimer.Stop()
					FadeInTimer.Dispose()
					FadeInTimer = Nothing
				End If
				If FadeOutTimer IsNot Nothing Then
					RemoveHandler FadeOutTimer.Tick, AddressOf FadeOutHandler
					FadeOutTimer.Stop()
					FadeOutTimer.Dispose()
					FadeOutTimer = Nothing
				End If
			End Sub
			Private Function GetImageBounds(img As Image, alignment As ImageAlignments, containerWidth As Integer, y As Integer) As Rectangle
				Dim x As Integer
				Select Case alignment
					Case ImageAlignments.Left
						x = _owner.TextPadding
					Case ImageAlignments.Right
						x = containerWidth - img.Width - _owner.TextPadding - _owner.ShadowThickness
				End Select
				Return New Rectangle(x, y, img.Width, img.Height)
			End Function

		End Class

	End Class

End Namespace
