
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Text

Namespace UI

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
		'Private Sub OnMouseEnter(sender As Object, e As EventArgs)

		'	If Not _manualTooltipActive Then
		'		HideTooltip()
		'	End If

		'	'If _manualTooltipActive Then Exit Sub 'Ignore if we're showing/hiding a tooltip programmatically

		'	'Dim ctrl As Control = CType(sender, Control)

		'	''Cancel any pending hide
		'	'HideDelayTimer.Stop()

		'	''If we're already showing for this control, do nothing
		'	'If _hoveredcontrol Is ctrl AndAlso IsVisible Then Exit Sub

		'	''Update hovered control
		'	'_hoveredcontrol = ctrl

		'	''Check if tooltip content exists
		'	'If tooltips.ContainsKey(ctrl) AndAlso Not String.IsNullOrWhiteSpace(tooltips(ctrl)) Then
		'	'	' Reset and start show delay
		'	'	If ShowDelayTimer Is Nothing Then
		'	'		ShowDelayTimer = New Timer()
		'	'		AddHandler ShowDelayTimer.Tick, AddressOf ShowTooltipDelayed
		'	'	Else
		'	'		ShowDelayTimer.Stop()
		'	'	End If
		'	'	If ShowDelay < 1 Then
		'	'		ShowDelayTimer.Interval = 1
		'	'	Else
		'	'		ShowDelayTimer.Interval = ShowDelay
		'	'	End If
		'	'	ShowDelayTimer.Start()
		'	'End If

		'End Sub
		'      Private Sub OnMouseHover(sender As Object, e As EventArgs)
		'	If _manualTooltipActive Then Exit Sub 'Ignore if we're showing/hiding a tooltip programmatically

		'	Dim ctrl As Control = CType(sender, Control)

		'	'If we're already showing for this control, do nothing
		'	If _hoveredcontrol Is ctrl AndAlso IsVisible Then Exit Sub

		'	'Cancel any pending hide
		'	HideDelayTimer.Stop()

		'	'Update hovered control
		'	_hoveredcontrol = ctrl

		'	'Check if tooltip content exists
		'	If tooltips.ContainsKey(ctrl) AndAlso Not String.IsNullOrWhiteSpace(tooltips(ctrl)) Then
		'		' Reset and start show delay
		'		If ShowDelayTimer Is Nothing Then
		'			ShowDelayTimer = New Timer()
		'			AddHandler ShowDelayTimer.Tick, AddressOf ShowTooltipDelayed
		'		Else
		'			ShowDelayTimer.Stop()
		'		End If
		'		If ShowDelay < 1 Then
		'			ShowDelayTimer.Interval = 1
		'		Else
		'			ShowDelayTimer.Interval = ShowDelay
		'		End If
		'		ShowDelayTimer.Start()
		'	End If
		'End Sub
		'Private Sub OnMouseLeave(sender As Object, e As EventArgs)
		'	ShowDelayTimer?.Stop()
		'	If _hoveredcontrol Is CType(sender, Control) Then StartHideDelayTimer()
		'End Sub
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
		End Sub
		''' <summary>
		''' Hides the tooltip if it is currently visible.
		''' </summary>
		Public Sub HideTooltip()
			'Trace.WriteLine($"[ToolTipEX] HideTooltip called, IsVisible={IsVisible}")
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
					Dim bounds As Rectangle = GetImageBounds(TooltipImage, _owner.ImageAlignment, Me.Width, CInt(Me.Height / 2 - TooltipImage.Height / 2) + If(isMultiLine, 0, 1))
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
	''' Extended Windows progress bar control.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	<DefaultProperty("PercentageMode")>
	Public Class ProgressEX
		Inherits System.Windows.Forms.UserControl

		'Declarations
		''' <summary>
		''' Specifies how the percentage value should be drawn
		''' </summary>
		Public Enum PercentageDrawModes As Integer
			None = 0 'No Percentage shown
			Center 'Percentage alwayes centered
			Movable 'Percentage moved with the progress activities
		End Enum
		Public Enum ColorDrawModes As Integer
			Gradient = 0
			Smooth
		End Enum
		Private maxValue As Integer 'Maximum value
		Private minValue As Integer 'Minimum value
		Private _value As Single 'Value property value
		Private stepValue As Integer 'Step value
		Private percentageValue As Single 'Percent value
		'Private drawingWidth As Integer 'Drawing width according to the logical Value property
		Private m_drawingColor As Color 'Color used for drawing activities
		Private ReadOnly gradientBlender As Drawing2D.ColorBlend 'Color mixer object
		Private percentageDrawMode As PercentageDrawModes 'Percent Drawing type
		Private colorDrawMode As ColorDrawModes
		Private _Brush As SolidBrush
		Private writingBrush As SolidBrush 'Percent writing brush
		Private writingFont As Font 'Font to write Percent with
		Private _Drawer As Drawing2D.LinearGradientBrush

		'Properties
		''' <summary>
		''' Gets or Sets a value determine how to display Percentage value
		''' </summary>
		<Category("Behavior"), Description("Specify how to display the Percentage value"), DefaultValue(PercentageDrawModes.Center)>
		Public Property PercentageMode As PercentageDrawModes
			Get
				Return percentageDrawMode
			End Get
			Set
				percentageDrawMode = Value
				'Me.Refresh()
			End Set
		End Property
		''' <summary>
		''' Gets or Sets a value to determine use of a color gradient
		''' </summary>
		<Category("Appearance"), Description("Specify how to display the Drawing Color"), DefaultValue(ColorDrawModes.Gradient)>
		Public Property DrawingColorMode As ColorDrawModes
			Get
				Return colorDrawMode
			End Get
			Set
				colorDrawMode = Value
				Me.Refresh()
			End Set
		End Property
		''' <summary>
		''' Gets or Sets the color used to draw the Progress activities
		''' </summary>
		<Category("Appearance"), Description("Specify the color used to draw the progress activities"), DefaultValue(GetType(Color), "Red")>
		Public Property DrawingColor As Color
			Get
				Return m_drawingColor
			End Get
			Set
				'If assigned then remix the colors used for gradient display
				m_drawingColor = Value
				gradientBlender.Colors(0) = ControlPaint.Dark(Value)
				gradientBlender.Colors(1) = ControlPaint.Light(Value)
				gradientBlender.Colors(2) = ControlPaint.Dark(Value)
				_Brush?.Dispose()
				_Brush = New SolidBrush(Value)
				Me.Invalidate(False)
			End Set
		End Property
		''' <summary>
		'''  Gets or sets the maximum value of the range of the control. 
		''' </summary>
		<Category("Layout"), Description("Specify the maximum value the progress can increased to"), DefaultValue(100)>
		Public Property Maximum As Integer
			Get
				Return maxValue
			End Get
			Set
				maxValue = Value
				Me.Refresh()
			End Set
		End Property
		''' <summary>
		''' Gets or sets the minimum value of the range of the control.
		''' </summary>
		<Category("Layout"), Description("Specify the minimum value the progress can decreased to"), DefaultValue(0)>
		Public Property Minimum As Integer
			Get
				Return minValue
			End Get
			Set
				minValue = Value
				Me.Refresh()
			End Set
		End Property
		''' <summary>
		'''  Gets or sets the amount by which a call to the System.Windows.Forms.ProgressBar.
		'''  StepForword method increases the current position of the progress bar.
		''' </summary>
		<Category("Layout"), Description("Specify the amount by which a call to the System.Windows.Forms.ProgressBar.StepForword method increases the current position of the progress bar"), DefaultValue(5)>
		Public Property [Step] As Integer
			Get
				Return stepValue
			End Get
			Set
				stepValue = Value
				Me.Refresh()
			End Set
		End Property
		''' <summary>
		''' Gets or sets the current position of the progress bar. 
		''' </summary>
		''' <exception cref="System.ArgumentException">The value specified is greater than the value of
		''' the System.Windows.Forms.ProgressBar.Maximum property.  -or- The value specified is less
		''' than the value of the System.Windows.Forms.ProgressBar.Minimum property</exception>
		<Category("Layout"), Description("Specify the current position of the progress bar"), DefaultValue(0)>
		Public Property Value As Integer
			Get
				Return CInt(Math.Truncate(_value))
			End Get
			Set
				'Protect the value and refuse any invalid values
				'Here we may just handle invalid values and dont bother the client by exceptions
				If Value > maxValue Or Value < minValue Then
					Throw New ArgumentException("Invalid value used")
				End If
				_value = Value
				Me.Refresh()
			End Set
		End Property
		''' <summary>
		''' Gets the Percent value the Progress activities reached
		''' </summary>
		Public ReadOnly Property Percent As Integer
			Get
				Return CInt(Math.Truncate(Math.Round(percentageValue))) 'Its float value, so to be accurate round it then return
			End Get
		End Property
		'This property exist in the parent, overide it for our own good
		''' <summary>
		''' Gets or Sets the color used to draw the Precentage value
		''' </summary>
		<Category("Appearance"), Description("Specify the font used to draw the Percentage value")>
		Public Overrides Property Font As Font
			Get
				Return writingFont
			End Get
			Set
				writingFont = Value
				Me.Invalidate(False)
			End Set
		End Property
		'This property exist in the parent, overide it for our own good
		''' <summary>
		''' Gets or Sets the color used to draw the Precentage value
		''' </summary>
		<Category("Appearance"), Description("Specify the color used to draw the Percentage value")>
		Public Overrides Property ForeColor As Color
			Get
				Return writingBrush.Color
			End Get
			Set
				writingBrush.Color = Value
				Me.Invalidate(False)
			End Set
		End Property

		'Events
		''' <summary>
		''' Initialize new instance of the ProgressEx control
		''' </summary>
		Public Sub New()

			'Initialize
			MyBase.New()
			Me.Name = "ProgressEx"
			Me.DoubleBuffered = True
			Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True) 'Cancel Reflection while drawing
			Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True) 'Allow Transparent backcolor
			Me.BackColor = Color.Transparent
			Me.MinimumSize = New Size(50, 5)
			Me.MaximumSize = New Size(Integer.MaxValue, 40)
			Me.Size = New System.Drawing.Size(96, 24)

			'Designer Stuff
			maxValue = 100
			minValue = 0
			stepValue = 5
			percentageDrawMode = PercentageDrawModes.Center
			colorDrawMode = ColorDrawModes.Gradient

			'ProgressEx Stuff
			m_drawingColor = Color.Red
			gradientBlender = New Drawing2D.ColorBlend() With {
					.Positions = {0.0F, 0.5F, 1.0F},
					.Colors = {ControlPaint.Dark(m_drawingColor), ControlPaint.Light(m_drawingColor), ControlPaint.Dark(m_drawingColor)}}
			writingFont = New Font("Arial", 10, FontStyle.Bold)
			writingBrush = New SolidBrush(Color.Black)

		End Sub
		''' <summary> 
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Sub Dispose(disposing As Boolean)
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
		Protected Overrides Sub OnPaint(e As PaintEventArgs)
			MyBase.OnPaint(e)
			If LicenseManager.UsageMode = LicenseUsageMode.Runtime AndAlso _value > 0 Then
				percentageValue = (_value - minValue) / (maxValue - minValue) * 100
				Dim w = CInt((Me.Width - 1) * (_value - minValue) / (maxValue - minValue))
				Select Case colorDrawMode
					Case ColorDrawModes.Gradient
						_Drawer.InterpolationColors = gradientBlender
						e.Graphics.FillRectangle(_Drawer, 0, 0, w, Me.Height)
					Case Else
						e.Graphics.FillRectangle(_Brush, 0, 0, w, Me.Height)
				End Select

				If percentageDrawMode <> PercentageDrawModes.None Then
					Dim txt = CInt(Math.Truncate(percentageValue)).ToString() & "%"
					Dim sz = e.Graphics.MeasureString(txt, writingFont)
					Debug.Print(percentageDrawMode.ToString)
					Dim x = If(percentageDrawMode = PercentageDrawModes.Movable, w, (Me.Width - sz.Width) / 2)
					Dim y = (Me.Height - sz.Height) / 2
					e.Graphics.DrawString(txt, writingFont, writingBrush, x, y)
				End If
			End If
		End Sub
		Protected Overrides Sub OnResize(e As EventArgs)
			MyBase.OnResize(e)
			If LicenseManager.UsageMode = LicenseUsageMode.Runtime Then
				_Drawer?.Dispose()
				_Drawer = New Drawing2D.LinearGradientBrush(New Rectangle(Point.Empty, Me.ClientSize), Color.Black, Color.White, Drawing2D.LinearGradientMode.Vertical)
				If gradientBlender IsNot Nothing Then _Drawer.InterpolationColors = gradientBlender
				Me.Invalidate()
			End If
		End Sub

		'Procedures
		''' <summary>
		''' Increment the progress one step
		''' </summary>
		Public Sub StepForward()
			If (_value + stepValue) < maxValue Then
				'If valid increment the value by step size
				_value += stepValue
				Me.Refresh()
			Else
				'If not dont exceed the maximum allowed
				_value = maxValue
				Me.Refresh()
			End If
		End Sub
		''' <summary>
		''' Decrement the progress one step
		''' </summary>
		Public Sub StepBackward()
			If (_value - stepValue) > minValue Then
				'If valid decrement the value by step size
				_value -= stepValue
				Me.Refresh()
			Else
				'If not dont exceed the minimum allowed
				_value = minValue
				Me.Refresh()
			End If
		End Sub

	End Class

	''' <summary>
	''' A simple ProgressBar with an alternate theme applied for a cleaner, modern look.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class DataBar
		Inherits System.Windows.Forms.ProgressBar

		'Events
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

	''' <summary>
	''' Color Selector ComboBox, Using System.Drawing.Color Structure
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ColorComboBox
		Inherits ComboBox

		'Declarations
		''' <summary>
		''' Gets/sets the selected color of ComboBox
		''' (Default color is Black)
		''' </summary>
		<System.ComponentModel.Category("Data")>
		<System.ComponentModel.ReadOnly(True)>
		<DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
		Public Property Color As Color
			Get
				If Me.SelectedItem IsNot Nothing Then
					Return CType(Me.SelectedItem, Color)
				End If
				Return Color.Black 'Default Setting
			End Get
			Set
				Dim ix As Integer = Me.Items.IndexOf(Value)
				If ix >= 0 Then
					Me.SelectedIndex = ix
				End If
			End Set
		End Property

		'Events
		Public Sub New()
			'Change DrawMode for custom drawing
			Me.DrawMode = DrawMode.OwnerDrawFixed
			Me.DropDownStyle = ComboBoxStyle.DropDownList
			'Add System.Drawing.Color Structure To Item List
			FillColors() 'May cause duplicates if used here for some reason, set manually by calling FillColors for each combobox in form constructor.
		End Sub
		Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
			If e.Index >= 0 Then
				Dim color As Color = CType(Me.Items(e.Index), Color)
				Dim brush As New SolidBrush(e.BackColor)
				Dim nextX As Integer = 0
				e.Graphics.FillRectangle(brush, e.Bounds)
				DrawColor(e, color, nextX)
				DrawText(e, color, nextX)
				brush.Dispose()
			Else
				MyBase.OnDrawItem(e)
			End If
		End Sub

		'Procedures
		Private Sub FillColors() 'Must Import System.Linq
			If Not Me.DesignMode Then
				'Populate colors
				Me.Items.Clear()
				'Fill Colors Using Reflection
				For Each color As Color In GetType(Color).GetProperties(Reflection.BindingFlags.[Static] Or Reflection.BindingFlags.[Public]).Where(Function(c) c.PropertyType Is GetType(Color)).[Select](Function(c) CType(c.GetValue(c, Nothing), Color))
					If Not color = Color.Transparent Then Me.Items.Add(color)
				Next
			End If
		End Sub
		Private Shared Sub DrawColor(e As DrawItemEventArgs, color As Color, ByRef nextX As Integer)
			Dim width As Integer = e.Bounds.Height * 2 - 8
			Dim rectangle As New Rectangle(e.Bounds.X + 3, e.Bounds.Y + 3, width, e.Bounds.Height - 6)
			Dim brush As New SolidBrush(color)
			e.Graphics.FillRectangle(brush, rectangle)
			nextX = width + 8
			brush.Dispose()
		End Sub
		Private Shared Sub DrawText(e As DrawItemEventArgs, color As Color, nextX As Integer)
			Dim brush As New SolidBrush(e.ForeColor)
			e.Graphics.DrawString(color.Name, e.Font, brush, New PointF(nextX, e.Bounds.Y + (e.Bounds.Height - e.Font.Height) \ 2))
			brush.Dispose()
		End Sub

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
				e.Handled = True
				e.SuppressKeyPress = True
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
	''' A ComboBox that supports custom items with text, images, and color swatches. Fixes the stock combobox dropdownlist draw background issues, making it theme-ready. Also includes hover effects for improved user experience.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ComboBox
		Inherits System.Windows.Forms.ComboBox

		' Declarations
		Private _hovering As Boolean = False

		' Control Events
		Protected Overrides Sub WndProc(ByRef m As Message)
			MyBase.WndProc(m)
			If m.Msg = WinAPI.WM_PAINT OrElse m.Msg = WinAPI.WM_PRINTCLIENT OrElse m.Msg = WinAPI.WM_ERASEBKGND Then
				Using g As Graphics = CreateGraphics()
					PaintCombo(g)
				End Using
			End If
		End Sub
		Public Sub New()
			MyBase.New()
			DrawMode = DrawMode.OwnerDrawFixed
			DropDownStyle = ComboBoxStyle.DropDownList
			DoubleBuffered = True
			SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
			UpdateStyles()
		End Sub
		Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
			MyBase.OnDrawItem(e)
			If e.Index < 0 Then Return

			Dim item = TryCast(Me.Items(e.Index), ComboItem)
			If item Is Nothing Then
				' Fallback for non-ComboItem entries
				e.DrawBackground()
				Using b As New SolidBrush(Me.ForeColor)
					e.Graphics.DrawString(Me.Items(e.Index).ToString(),
								  Me.Font,
								  b,
								  e.Bounds.Left + 4,
								  e.Bounds.Top + (e.Bounds.Height - Me.Font.Height) \ 2)
				End Using
				e.DrawFocusRectangle()
				Return
			End If

			Dim g = e.Graphics
			g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

			' Background (selection aware)
			Dim bg As Color = If((e.State And DrawItemState.Selected) = DrawItemState.Selected,
						 SystemColors.Highlight,
						 Me.BackColor)

			Using b As New SolidBrush(bg)
				g.FillRectangle(b, e.Bounds)
			End Using

			' Text color (selection aware)
			Dim textColor As Color = If((e.State And DrawItemState.Selected) = DrawItemState.Selected,
								SystemColors.HighlightText,
								Me.ForeColor)

			Dim iconSize As Integer = e.Bounds.Height - 4
			Dim iconX As Integer = e.Bounds.Left + 4
			Dim iconY As Integer = e.Bounds.Top + 2
			Dim textX As Integer = iconX

			' Image And Swatch
			If item.Image IsNot Nothing Then ' 1) Image
				g.DrawImage(item.Image, New Rectangle(iconX, iconY, iconSize, iconSize))
				textX += iconSize + 6
			ElseIf item.Swatch.HasValue Then ' 2) Color swatch
				Dim swatchRect As New Rectangle(iconX, iconY, iconSize, iconSize)
				Using b As New SolidBrush(item.Swatch.Value)
					g.FillRectangle(b, swatchRect)
				End Using
				Using p As New Pen(Color.Black)
					g.DrawRectangle(p, swatchRect)
				End Using
				textX += iconSize + 6
			Else ' 3) Neither – text starts near left
				textX = iconX
			End If

			' Text
			Using b As New SolidBrush(textColor)
				g.DrawString(item.Text,
					 Me.Font,
					 b,
					 textX,
					 e.Bounds.Top + (e.Bounds.Height - Me.Font.Height) \ 2)
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
		Protected Overrides Sub OnDropDownClosed(e As EventArgs)
			MyBase.OnDropDownClosed(e)
			_hovering = False
			Invalidate()
		End Sub

		' Methods
		Private Sub PaintCombo(g As Graphics)
			g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
			Dim rc As Rectangle = ClientRectangle

			' Background
			Dim bg As Color = If(_hovering, ControlPaint.Light(BackColor, 0.25F), BackColor)
			Using b As New SolidBrush(bg)
				g.FillRectangle(b, rc)
			End Using

			' Determine selected item
			Dim text As String = ""
			Dim selectedItem = TryCast(Me.SelectedItem, ComboItem)
			Dim textX As Integer
			If selectedItem IsNot Nothing Then
				text = selectedItem.Text
				Dim iconSize As Integer = rc.Height - 6
				Dim iconX As Integer = 4
				Dim iconY As Integer = (rc.Height - iconSize) \ 2
				If selectedItem.Image IsNot Nothing Then ' Image
					g.DrawImage(selectedItem.Image,
						New Rectangle(iconX, iconY, iconSize, iconSize))
					textX = iconX + iconSize + 6
				ElseIf selectedItem.Swatch.HasValue Then ' Color swatch
					Dim swatchRect As New Rectangle(iconX, iconY, iconSize, iconSize)
					Using b As New SolidBrush(selectedItem.Swatch.Value)
						g.FillRectangle(b, swatchRect)
					End Using
					Using p As New Pen(Color.Black)
						g.DrawRectangle(p, swatchRect)
					End Using
					textX = iconX + iconSize + 6
				Else
					textX = 4
				End If
			Else ' No selected ComboItem – fallback
				If Me.SelectedItem IsNot Nothing Then
					text = Me.SelectedItem.ToString()
				End If
				textX = 4
			End If

			' Text
			Dim arrowWidth As Integer = 6
			Dim textRect As New Rectangle(textX, rc.Top, rc.Width - textX - arrowWidth - 4, rc.Height)
			Dim sf As New StringFormat With {
				.FormatFlags = StringFormatFlags.NoWrap,
				.Trimming = StringTrimming.EllipsisCharacter,
				.LineAlignment = StringAlignment.Center,   ' vertical centering
				.Alignment = StringAlignment.Near         ' left align
				}
			Using b As New SolidBrush(ForeColor)
				g.DrawString(text, Me.Font, b, textRect, sf)
			End Using

			' Arrow
			Dim arrowSize As Integer = 6
			Dim arrowX As Integer = rc.Right - arrowSize - 6
			Dim arrowY As Integer = (rc.Height - arrowSize) \ 2
			Dim pts As Point() = {
				New Point(arrowX, arrowY),
				New Point(arrowX + arrowSize, arrowY),
				New Point(arrowX + arrowSize \ 2, arrowY + arrowSize)
			}
			Dim arrowColor As Color = If(_hovering, ControlPaint.Dark(ForeColor, 0.1F), ForeColor)
			Using b As New SolidBrush(arrowColor)
				g.FillPolygon(b, pts)
			End Using

		End Sub

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
	''' Extended Listview control with Insertion Line for drag/drop operations.
	''' </summary>
	<ToolboxItem(True)>
	<DesignerCategory("Code")>
	Public Class ListViewEX
		Inherits ListView

		'Declarations
		Private _LineBefore As Integer = -1
		Private _LineAfter As Integer = -1
		Private _InsertionLineColor As Color = Color.Teal

		'Properties
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

		'Events
		Public Sub New()
			SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
		End Sub
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

		'Procedures
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

#Region "Toast System"

	' Public API
	Public Module Toast

		Public Sub ShowToast(options As ToastOptions)
			ToastManager.Show(options)
		End Sub

	End Module

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

		Public ReadOnly _width As Integer
		Public _height As Integer
		Private _titleHeight As Integer
		Private _messageHeight As Integer
		Private _iconSize As Integer = 0
		Private _opacity As Byte = 0

		' ---- Layout Constants ----
		Private Const TOAST_PADDING As Integer = 10
		Private Const TITLE_MESSAGE_GAP As Integer = 5
		Private Const ICON_SIZE As Integer = 48
		Private Const MAX_HEIGHT_PADDING As Integer = 40
		Private Const TEXT_TOP_OFFSET As Integer = 7


		Private _lastPos As System.Drawing.Point
		Private _hasInitialPosition As Boolean = False

		Public Event ToastClosed()

		' ------------- WndProc ---------------------------
		Private Function WindowProc(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr

			Select Case msg
				Case WinAPI.WM_NCHITTEST
					Return CType(WinAPI.HTCAPTION, IntPtr)
				Case WinAPI.WM_NCLBUTTONDOWN
					CloseToast()
					Return IntPtr.Zero
				Case WinAPI.WM_NCRBUTTONUP
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
			_width = opts.Width

			_className = "LayeredToast_" & Guid.NewGuid().ToString("N")
			_wndProc = AddressOf Me.WindowProc

			RegisterWindowClass()
			CreateLayeredWindow()

			' Now that window exists, measure text and compute final height
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
			If _hwnd <> IntPtr.Zero Then
				WinAPI.DestroyWindow(_hwnd)
				_hwnd = IntPtr.Zero
			End If
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
		Private Sub RenderToast(g As Graphics)
			Dim w = _width
			Dim h = _height

			Dim bgColor = _opts.BackColor
			Dim borderColor = _opts.BorderColor
			Dim foreColor = _opts.ForeColor
			Dim radius = _opts.CornerRadius

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
				Dim size As Integer = h - padding * 2
				iconRect = New Rectangle(padding, padding, size, size)

				If radius > 0 Then
					Using path As New GraphicsPath()
						path.AddArc(iconRect.X, iconRect.Y, radius, radius, 180, 90)
						path.AddArc(iconRect.Right - radius, iconRect.Y, radius, radius, 270, 90)
						path.AddArc(iconRect.Right - radius, iconRect.Bottom - radius, radius, radius, 0, 90)
						path.AddArc(iconRect.X, iconRect.Bottom - radius, radius, radius, 90, 90)
						path.CloseFigure()

						g.SetClip(path)
						g.DrawImage(_opts.Image, iconRect)
						g.ResetClip()
					End Using
				Else
					g.DrawImage(_opts.Image, iconRect)
				End If

				textX = iconRect.Right + padding

			ElseIf _opts.Icon IsNot Nothing Then
				Dim size As Integer = ICON_SIZE

				' Center the icon square vertically
				Dim iconY As Integer = padding + (h - padding * 2 - size) \ 2
				iconRect = New Rectangle(padding, iconY, size, size)

				Using bmp As Bitmap = _opts.Icon.ToBitmap()
					Dim bmpW As Integer = bmp.Width
					Dim bmpH As Integer = bmp.Height

					' Center the bitmap INSIDE iconRect
					Dim offsetX As Integer = iconRect.X + (size - bmpW) \ 2
					Dim offsetY As Integer = iconRect.Y + (size - bmpH) \ 2

					Dim centeredRect As New Rectangle(offsetX, offsetY, bmpW, bmpH)

					g.DrawImage(bmp, centeredRect)
				End Using

				textX = iconRect.Right + 2
			Else
				' No icon, no image → add horizontal inset
				textX = padding + 2
			End If

			Dim wAvail As Integer = w - textX - padding

			'-----------------------------------------
			' TITLE (TextRenderer)
			'-----------------------------------------
			Dim titleY As Integer = padding + TEXT_TOP_OFFSET

			If Not String.IsNullOrEmpty(_opts.Title) AndAlso _titleHeight > 0 Then
				Dim titleRectF As New RectangleF(textX, padding + 2, wAvail, _titleHeight)
				Using brush As New SolidBrush(foreColor)
					g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
					Using sf As New StringFormat(StringFormatFlags.LineLimit)
						sf.Trimming = StringTrimming.EllipsisCharacter
						g.DrawString(_opts.Title, _opts.TitleFont, brush, titleRectF, sf)
					End Using
				End Using
			End If

			'-----------------------------------------
			' MESSAGE (TextRenderer)
			'-----------------------------------------
			Dim messageY As Integer

			If String.IsNullOrEmpty(_opts.Title) Then
				' No title: message starts at the normal text top area
				messageY = padding + TEXT_TOP_OFFSET + 3
			Else
				' Normal layout when title exists: message under the title
				messageY = padding + TEXT_TOP_OFFSET + _titleHeight + TITLE_MESSAGE_GAP
			End If

			If Not String.IsNullOrEmpty(_opts.Message) AndAlso _messageHeight > 0 Then
				Dim messageRectF As New RectangleF(textX, padding + _titleHeight + TITLE_MESSAGE_GAP, wAvail, _messageHeight)
				Using brush As New SolidBrush(foreColor)
					g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
					Using sf As New StringFormat(StringFormatFlags.LineLimit)
						sf.Trimming = StringTrimming.EllipsisWord
						g.DrawString(_opts.Message, _opts.MessageFont, brush, messageRectF, sf)
					End Using
				End Using
			End If
			'Debug.WriteLine($"HEIGHT = {_height}")
			'Debug.WriteLine($"MESSAGEHEIGHT = {_messageHeight}")
			'Debug.WriteLine($"ICONSIZE = {_iconSize}")
			'Debug.WriteLine($"PADDING = {TOAST_PADDING}")
		End Sub
		Private Sub AutoSizeToast()
			Dim padding As Integer = TOAST_PADDING

			_titleHeight = 0
			_messageHeight = 0

			' Compute icon size FIRST
			If _opts.Image IsNot Nothing Then
				_iconSize = Math.Max(24, Math.Min(_messageHeight, 96))
			ElseIf _opts.Icon IsNot Nothing Then
				_iconSize = ICON_SIZE
			Else
				_iconSize = 0
			End If

			' Compute textX
			Dim textX As Integer = padding
			If _iconSize > 0 Then
				textX = padding + _iconSize + padding
			End If

			Dim wAvail As Integer = _width - textX - padding

			' Measure using DrawString-compatible metrics
			Using bmp As New Bitmap(1, 1)
				Using g As Graphics = Graphics.FromImage(bmp)
					g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

					' Title
					If Not String.IsNullOrEmpty(_opts.Title) Then
						Using sf As New StringFormat(StringFormatFlags.LineLimit)
							sf.Trimming = StringTrimming.EllipsisCharacter
							Dim sizeF As SizeF = g.MeasureString(_opts.Title, _opts.TitleFont, wAvail, sf)
							_titleHeight = CInt(Math.Ceiling(sizeF.Height))
						End Using
					End If

					' Message
					If Not String.IsNullOrEmpty(_opts.Message) Then
						Using sf As New StringFormat(StringFormatFlags.LineLimit)
							sf.Trimming = StringTrimming.EllipsisWord
							Dim sizeF As SizeF = g.MeasureString(_opts.Message, _opts.MessageFont, wAvail, sf)
							_messageHeight = CInt(Math.Ceiling(sizeF.Height))
						End Using
					End If
				End Using
			End Using

			' NEW: Clean DrawString-based height math
			Dim totalHeight As Integer =
				padding +
				_titleHeight +
				TITLE_MESSAGE_GAP +
				_messageHeight +
				padding

			' Clamp
			Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
			Dim maxHeight As Integer = wa.Height - MAX_HEIGHT_PADDING

			If totalHeight > maxHeight Then
				totalHeight = maxHeight
				Dim availableForMessage As Integer =
					totalHeight - padding - _titleHeight - TITLE_MESSAGE_GAP - padding
				_messageHeight = Math.Max(availableForMessage, 0)
			End If

			_height = totalHeight

			' Final icon size
			If _opts.Image IsNot Nothing Then
				_iconSize = _height - TOAST_PADDING * 2
			ElseIf _opts.Icon IsNot Nothing Then
				_iconSize = ICON_SIZE
			Else
				_iconSize = 0
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

#End Region

#Region "Theme System"

	Public Class SkyeTheme

		' Identity
		Public Property Name As String

		' Core
		Public Property BackColor As Color
		Public Property ForeColor As Color
		Public Property AccentColor As Color
		Public Property BorderColor As Color

		' Buttons
		Public Property ButtonBack As Color
		Public Property ButtonFore As Color

		' Text boxes & RTB
		Public Property TextBack As Color
		Public Property TextFore As Color

		' GroupBox, labels, etc.
		Public Property GroupBoxFore As Color

		' DataGridView
		Public Property GridBack As Color
		Public Property GridFore As Color
		Public Property GridHeaderBack As Color
		Public Property GridHeaderFore As Color
		Public Property GridBorder As Color
		Public Property GridAlternateRowBack As Color

		' Tooltips
		Public Property TooltipBack As Color
		Public Property TooltipFore As Color
		Public Property TooltipBorder As Color

        ' Menus
        Public Property MenuBack As Color
		Public Property MenuFore As Color
		Public Property MenuHover As Color
		Public Property MenuBorder As Color
		Public Property MenuSeparator As Color

	End Class

	Public Module SkyeThemes

		Public ReadOnly Light As New SkyeTheme With {
			.Name = "Light",
			.BackColor = Color.WhiteSmoke,
			.ForeColor = Color.Black,
			.AccentColor = Color.DeepSkyBlue,
			.BorderColor = Color.LightGray,
			.ButtonBack = Color.Gainsboro,
			.ButtonFore = Color.Black,
			.TextBack = Color.WhiteSmoke,
			.TextFore = Color.Black,
			.GroupBoxFore = Color.Black,
			.GridBack = Color.WhiteSmoke,
			.GridFore = Color.Black,
			.GridHeaderBack = Color.Gainsboro,
			.GridHeaderFore = Color.Black,
			.GridBorder = Color.LightGray,
			.GridAlternateRowBack = Color.FromArgb(245, 245, 245),
			.TooltipBack = Color.WhiteSmoke,
			.TooltipFore = Color.Black,
			.TooltipBorder = Color.LightGray,
			.MenuBack = Color.WhiteSmoke,
			.MenuFore = Color.Black,
			.MenuHover = Color.FromArgb(230, 230, 230),
			.MenuBorder = Color.LightGray,
			.MenuSeparator = Color.LightGray
		}
		Public ReadOnly Dark As New SkyeTheme With {
			.Name = "Dark",
			.BackColor = Color.FromArgb(32, 32, 32),
			.ForeColor = Color.White,
			.AccentColor = Color.DeepSkyBlue,
			.BorderColor = Color.FromArgb(64, 64, 64),
			.ButtonBack = Color.FromArgb(45, 45, 45),
			.ButtonFore = Color.White,
			.TextBack = Color.FromArgb(40, 40, 40),
			.TextFore = Color.White,
			.GroupBoxFore = Color.White,
			.GridBack = Color.FromArgb(32, 32, 32),
			.GridFore = Color.White,
			.GridHeaderBack = Color.FromArgb(45, 45, 45),
			.GridHeaderFore = Color.White,
			.GridBorder = Color.FromArgb(70, 70, 70),
			.GridAlternateRowBack = Color.FromArgb(40, 40, 40),
			.TooltipBack = Color.FromArgb(50, 50, 50),
			.TooltipFore = Color.White,
			.TooltipBorder = Color.FromArgb(80, 80, 80),
			.MenuBack = Color.FromArgb(40, 40, 40),
			.MenuFore = Color.White,
			.MenuHover = Color.FromArgb(60, 60, 60),
			.MenuBorder = Color.FromArgb(80, 80, 80),
			.MenuSeparator = Color.FromArgb(90, 90, 90)
		}
		Public ReadOnly Blossom As New SkyeTheme With {
			.Name = "Blossom",
			.BackColor = Color.Pink,
			.ForeColor = Color.DeepPink,
			.AccentColor = Color.White,
			.BorderColor = ControlPaint.Light(Color.DeepPink, 0.75F),
			.ButtonBack = Color.HotPink,
			.ButtonFore = Color.White,
			.TextBack = Color.Pink,
			.TextFore = Color.DeepPink,
			.GroupBoxFore = Color.DeepPink,
			.GridBack = Color.Pink,
			.GridFore = Color.DeepPink,
			.GridHeaderBack = Color.Pink,
			.GridHeaderFore = Color.DeepPink,
			.GridBorder = ControlPaint.Light(Color.DeepPink, 0.75F),
			.GridAlternateRowBack = ControlPaint.Light(Color.Pink, 0.25F),
			.TooltipBack = Color.Pink,
			.TooltipFore = Color.DeepPink,
			.TooltipBorder = ControlPaint.Light(Color.DeepPink, 0.75F),
			.MenuBack = Color.Pink,
			.MenuFore = Color.DeepPink,
			.MenuHover = Color.LightPink,
			.MenuBorder = Color.DeepPink,
			.MenuSeparator = Color.LightGray
		}
		Public ReadOnly CrimsonNight As New SkyeTheme With {
			.Name = "Crimson Night",
			.BackColor = Color.FromArgb(255, 35, 35, 35),
			.ForeColor = Color.DeepPink,
			.AccentColor = Color.White,
			.BorderColor = ControlPaint.Dark(Color.DeepPink, 0.25F),
			.ButtonBack = Color.DeepPink,
			.ButtonFore = Color.White,
			.TextBack = Color.FromArgb(255, 35, 35, 35),
			.TextFore = Color.DeepPink,
			.GroupBoxFore = Color.White,
			.GridBack = Color.FromArgb(255, 35, 35, 35),
			.GridFore = Color.DeepPink,
			.GridHeaderBack = Color.FromArgb(255, 35, 35, 35),
			.GridHeaderFore = Color.White,
			.GridBorder = ControlPaint.Dark(Color.DeepPink, 0.25F),
			.GridAlternateRowBack = Color.FromArgb(255, 40, 40, 40),
			.TooltipBack = Color.FromArgb(255, 35, 35, 35),
			.TooltipFore = Color.DeepPink,
			.TooltipBorder = ControlPaint.Dark(Color.DeepPink, 0.25F),
			.MenuBack = Color.DeepPink,
			.MenuFore = Color.White,
			.MenuHover = Color.HotPink,
			.MenuBorder = Color.FromArgb(255, 80, 80, 80),
			.MenuSeparator = Color.FromArgb(255, 90, 90, 90)
		}
		Private ReadOnly _themes As New List(Of SkyeTheme) From {Light, Dark, Blossom, CrimsonNight}
		Public ReadOnly Property AllThemes As List(Of SkyeTheme)
			Get
				Return _themes
			End Get
		End Property

		Public Sub AddTheme(t As SkyeTheme)
			_themes.Add(t)
		End Sub
		Public Sub RemoveTheme(t As SkyeTheme)
			If t Is Light OrElse t Is Dark Then Exit Sub
			_themes.Remove(t)
		End Sub
		Public Sub RemoveTheme(name As String)
			Dim t = _themes.FirstOrDefault(Function(x) String.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase))
			If t IsNot Nothing AndAlso Not (t Is Light OrElse t Is Dark) Then _themes.Remove(t)
		End Sub
		Public Function GetTheme(name As String) As SkyeTheme
			Dim t = _themes.FirstOrDefault(Function(x) String.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase))
			If t IsNot Nothing Then Return t
			Return Light
		End Function

	End Module

	Public Module ThemeManager

		Public Property CurrentTheme As SkyeTheme = SkyeThemes.Dark
		Public Event ThemeChanged As EventHandler

		Public Sub SetTheme(theme As SkyeTheme)
			CurrentTheme = theme
			RaiseEvent ThemeChanged(Nothing, EventArgs.Empty)
		End Sub

		' Apply theme to a single form
		Public Sub ApplyTheme(target As Form)
			If target Is Nothing Then Return

			target.BackColor = CurrentTheme.BackColor
			target.ForeColor = CurrentTheme.ForeColor

			ApplyToControls(target.Controls)

			' If the form itself has a menu
			If target.ContextMenuStrip IsNot Nothing Then
				ApplyToMenu(target.ContextMenuStrip)
			End If
		End Sub
		' Apply theme to all forms in the application
		Public Sub ApplyThemeToAllOpenForms()
			For Each f As Form In Application.OpenForms
				ApplyTheme(f)
				f.Invalidate()
			Next
		End Sub

		' Recursively theme all controls
		Private Sub ApplyToControls(controls As Control.ControlCollection)
			For Each c As Control In controls

				' Theme any context menu attached to the control
				If c.ContextMenuStrip IsNot Nothing Then
					ApplyToMenu(c.ContextMenuStrip)
				End If

				Select Case True
					Case TypeOf c Is Button
						Dim b = DirectCast(c, Button)
						b.BackColor = CurrentTheme.ButtonBack
						b.ForeColor = CurrentTheme.ButtonFore
					Case TypeOf c Is TextBoxBase
						Dim t = DirectCast(c, TextBoxBase)
						t.BackColor = CurrentTheme.TextBack
						t.ForeColor = CurrentTheme.TextFore
					Case TypeOf c Is System.Windows.Forms.RichTextBox
						Dim r = DirectCast(c, System.Windows.Forms.RichTextBox)
						r.BackColor = CurrentTheme.TextBack
						r.ForeColor = CurrentTheme.TextFore
					Case TypeOf c Is System.Windows.Forms.ComboBox
						Dim cb = DirectCast(c, System.Windows.Forms.ComboBox)
						cb.BackColor = CurrentTheme.ButtonBack
						cb.ForeColor = CurrentTheme.ButtonFore
					Case TypeOf c Is GroupBox
						c.ForeColor = CurrentTheme.GroupBoxFore
					Case TypeOf c Is Panel OrElse TypeOf c Is SplitContainer
						c.BackColor = CurrentTheme.BackColor
					Case TypeOf c Is DataGridView
						ApplyToDataGridView(DirectCast(c, DataGridView))
					Case TypeOf c Is ListView
						ApplyToListView(DirectCast(c, ListView))
					Case TypeOf c Is StatusStrip
						c.BackColor = CurrentTheme.BackColor
						c.ForeColor = CurrentTheme.ForeColor
				End Select

				If c.HasChildren Then
					ApplyToControls(c.Controls)
				End If

			Next
		End Sub

		' DataGridView theming (generic, no app-specific assumptions)
		Private Sub ApplyToDataGridView(dgv As DataGridView)
			dgv.BackgroundColor = CurrentTheme.GridBack
			dgv.BorderStyle = BorderStyle.None

			dgv.EnableHeadersVisualStyles = False
			dgv.GridColor = CurrentTheme.GridBorder

			Dim header = dgv.ColumnHeadersDefaultCellStyle
			header.BackColor = CurrentTheme.GridHeaderBack
			header.ForeColor = CurrentTheme.GridHeaderFore
			header.SelectionBackColor = CurrentTheme.GridHeaderBack
			header.SelectionForeColor = CurrentTheme.GridHeaderFore

			With dgv
				.EnableHeadersVisualStyles = False
				.RowHeadersDefaultCellStyle.BackColor = CurrentTheme.GridBack
				.RowHeadersDefaultCellStyle.ForeColor = CurrentTheme.GridFore
				.RowHeadersDefaultCellStyle.SelectionBackColor = CurrentTheme.AccentColor
				.RowHeadersDefaultCellStyle.SelectionForeColor = CurrentTheme.GridFore
				'.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
				.RowHeadersWidth = 24
			End With

			Dim cell = dgv.DefaultCellStyle
			cell.BackColor = CurrentTheme.GridBack
			cell.ForeColor = CurrentTheme.GridFore
			cell.SelectionBackColor = CurrentTheme.AccentColor
			cell.SelectionForeColor = CurrentTheme.ForeColor

			Dim alt = dgv.AlternatingRowsDefaultCellStyle
			alt.BackColor = CurrentTheme.GridAlternateRowBack
			alt.ForeColor = CurrentTheme.GridFore
		End Sub

		' Theme a tooltip
		Public Sub ApplyToTooltip(tip As System.Windows.Forms.ToolTip)
			If tip Is Nothing Then Return
			tip.BackColor = CurrentTheme.TooltipBack
			tip.ForeColor = CurrentTheme.TooltipFore
		End Sub
		Public Sub ApplyToTooltip(tip As Skye.UI.ToolTip)
			If tip Is Nothing Then Return
			tip.BackColor = CurrentTheme.TooltipBack
			tip.ForeColor = CurrentTheme.TooltipFore
			tip.BorderColor = CurrentTheme.TooltipBorder
		End Sub
        Public Sub ApplyToTooltip(tip As ToolTipEX)
            If tip Is Nothing Then Return
            tip.BackColor = CurrentTheme.TooltipBack
            tip.ForeColor = CurrentTheme.TooltipFore
            tip.BorderColor = CurrentTheme.TooltipBorder
        End Sub

        ' Theme a ListView (generic)
        Private Sub ApplyToListView(lv As ListView)
			' Only apply basic theming when OwnerDraw is False
			If lv.OwnerDraw = False Then
				lv.BackColor = CurrentTheme.BackColor
				lv.ForeColor = CurrentTheme.ForeColor
				lv.BorderStyle = BorderStyle.FixedSingle

				' Optional: makes the control feel more "modern"
				lv.HideSelection = False
			End If
		End Sub

		' Theme a menu (generic)
		Public Sub ApplyToMenu(menu As ContextMenuStrip)
			If menu Is Nothing Then Return

			menu.Renderer = New SkyeMenuRenderer()
			menu.BackColor = CurrentTheme.MenuBack
			menu.ForeColor = CurrentTheme.MenuFore

			For Each item As ToolStripItem In menu.Items
				item.ForeColor = CurrentTheme.MenuFore
			Next
		End Sub

	End Module
	Public Class SkyeMenuRenderer
		Inherits ToolStripProfessionalRenderer

		Protected Overrides Sub OnRenderToolStripBackground(e As ToolStripRenderEventArgs)
			Using b As New SolidBrush(ThemeManager.CurrentTheme.MenuBack)
				e.Graphics.FillRectangle(b, e.AffectedBounds)
			End Using
		End Sub
		Protected Overrides Sub OnRenderMenuItemBackground(e As ToolStripItemRenderEventArgs)
			Dim t = ThemeManager.CurrentTheme
			Dim g = e.Graphics
			Dim rect = New Rectangle(Point.Empty, e.Item.Size)

			Dim backColor As Color = If(e.Item.Selected OrElse e.Item.Pressed,
										t.MenuHover,
										t.MenuBack)

			Using b As New SolidBrush(backColor)
				g.FillRectangle(b, rect)
			End Using
		End Sub
		Protected Overrides Sub OnRenderToolStripBorder(e As ToolStripRenderEventArgs)
			Dim t = ThemeManager.CurrentTheme
			Dim g = e.Graphics
			Dim rect = New Rectangle(Point.Empty, e.ToolStrip.Size - New Size(1, 1))

			Using p As New Pen(t.MenuBorder)
				g.DrawRectangle(p, rect)
			End Using
		End Sub
		Protected Overrides Sub OnRenderImageMargin(e As ToolStripRenderEventArgs)
			Using b As New SolidBrush(ThemeManager.CurrentTheme.MenuBack)
				e.Graphics.FillRectangle(b, e.AffectedBounds)
			End Using
		End Sub
		'Protected Overrides Sub OnRenderSeparator(e As ToolStripSeparatorRenderEventArgs)
		'	Dim t = ThemeManager.CurrentTheme
		'	Dim g = e.Graphics
		'	Dim rect = New Rectangle(0, 0, e.Item.Width, e.Item.Height)
		'	Dim y = rect.Top + rect.Height \ 2

		'	Using p As New Pen(t.MenuSeparator)
		'		g.DrawLine(p, rect.Left + 4, y, rect.Right - 4, y)
		'	End Using
		'End Sub
		Protected Overrides Sub OnRenderSeparator(e As ToolStripSeparatorRenderEventArgs)
			Dim c = ThemeManager.CurrentTheme.MenuSeparator
			Using p As New Pen(c)
				Dim y = e.Item.Height \ 2
				e.Graphics.DrawLine(p, 2, y, e.Item.Width - 2, y)
			End Using
		End Sub
		Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
			Dim t = ThemeManager.CurrentTheme
			Dim g = e.Graphics
			Dim item = e.Item

			Dim textColor As Color = If(item.Enabled,
										t.MenuFore,
										ControlPaint.LightLight(t.MenuFore))

			TextRenderer.DrawText(
				g,
				item.Text,
				e.TextFont,
				e.TextRectangle,
				textColor,
				TextFormatFlags.Left Or TextFormatFlags.VerticalCenter
			)
		End Sub

	End Class

#End Region

End Namespace
