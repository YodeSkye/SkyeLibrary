
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms

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

				'Draw Image
				If TooltipImage IsNot Nothing Then
					Dim bounds As Rectangle = GetImageBounds(TooltipImage, _owner.ImageAlignment, Me.Width, CInt(Me.Height / 2 - TooltipImage.Height / 2) - 1)
					g.DrawImage(TooltipImage, bounds)
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
				TextRenderer.DrawText(g, TooltipText, _owner.Font, textRect, _owner.ForeColor, TextFormatFlags.Left Or TextFormatFlags.Top Or TextFormatFlags.WordBreak Or TextFormatFlags.NoPrefix)

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
	''' Improved And Extended RichTextBox control that prevents the cursor from changing to the I-beam when hovering over the control. It can cause cursor "blinking". This is especially useful in scenarios where the RichTextBox is used for display purposes only and should not allow text selection or editing.
	''' Also includes a SetAlignment method to easily set text alignment.
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

		Public Sub SetAlignment(align As HorizontalAlignment)
			Me.SelectAll()
			Me.SelectionAlignment = align
			Me.DeselectAll()
		End Sub

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
		Public Property Icon As Icon = Nothing
		Public Property TitleFont As Font = New Font("Segoe UI", 11, FontStyle.Bold)
		Public Property MessageFont As Font = New Font("Segoe UI", 9, FontStyle.Regular)
		Public Property BackColor As Color = Color.FromArgb(40, 40, 40)
		Public Property BorderColor As Color = Color.White
		Public Property ForeColor As Color = Color.White
		Public Property CornerRadius As Integer = 16

	End Class

	' Manager
	Friend Class ToastManager

        ' Declarations
        Private Shared ReadOnly ActiveToasts As New List(Of ToastWindow)
		Private Shared ReadOnly ToastWidth As Integer = 300 'REMOVE THESE
        Private Shared ReadOnly ToastHeight As Integer = 80
        Private Shared ReadOnly Margin As Integer = 10
        Private Shared toast As ToastWindow

		' Events
		Public Shared Sub Show(opts As ToastOptions)

            ' Play sound if requested
            If opts.PlaySound Then
                System.Media.SystemSounds.Hand.Play()
            End If

            ' Create toast window
            toast = New ToastWindow(opts)

            ' Track active toasts BEFORE positioning
            ActiveToasts.Add(toast)

            ' Position bottom-right, stacking upward
            RearrangeToasts()

            ' Show toast at its assigned position
            toast.ShowToastAt(toast.TargetPosition)

			AddHandler toast.ToastClosed,
                Sub()
                    ActiveToasts.Remove(toast)
				End Sub

		End Sub

		' Methods
		Public Shared Sub RearrangeToasts()
            If ActiveToasts.Count = 0 Then Exit Sub

            Dim first = ActiveToasts(0)
            Dim area = Screen.FromPoint(first.TargetPosition).WorkingArea

            ' Start at the TOP of the stack
            Dim y = area.Bottom - ((ToastHeight + Margin) * ActiveToasts.Count)

            For Each toast In ActiveToasts
                toast.TargetPosition = New Point(area.Right - ToastWidth - Margin, y)
                toast.MoveTo(toast.TargetPosition)
                y += (ToastHeight + Margin) ' stack downward
            Next
        End Sub

    End Class

	' Win32 Layered Window
	Public Class LayeredToastWindow
		Implements IDisposable

		Private _hwnd As IntPtr
		Private ReadOnly _className As String
		Private ReadOnly _wndProc As WinAPI.WndProcDelegate
		Private ReadOnly _opts As ToastOptions

		Private ReadOnly _width As Integer
		Private ReadOnly _height As Integer
		Private _opacity As Byte = 0

		Private _lastPos As System.Drawing.Point
		Private _hasInitialPosition As Boolean = False

		Public Event ToastClosed()

		' ------------- WndProc ---------------------------
		Private Function WindowProc(hWnd As IntPtr,
								msg As UInteger,
								wParam As IntPtr,
								lParam As IntPtr) As IntPtr

			Select Case CInt(msg)
				Case WinAPI.WM_MOUSEACTIVATE
					Return New IntPtr(WinAPI.MA_NOACTIVATE)

				Case WinAPI.WM_NCHITTEST
					Return New IntPtr(WinAPI.HTCLIENT)

				Case WinAPI.WM_DESTROY
					' no special cleanup needed here
			End Select

			Return WinAPI.DefWindowProc(hWnd, msg, wParam, lParam)
		End Function

		' ------------- Constructor -------------------------------
		Public Sub New(opts As ToastOptions)
			_opts = opts
			_width = opts.Width
			_height = opts.Height

			_className = "LayeredToast_" & Guid.NewGuid().ToString("N")
			_wndProc = AddressOf Me.WindowProc

			RegisterWindowClass()
			CreateLayeredWindow()

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
		Public Sub MoveTo(screenPos As System.Drawing.Point)
			' Store the new position
			_lastPos = screenPos
			_hasInitialPosition = True

			' Move the window WITHOUT redrawing the bitmap
			WinAPI.SetWindowPos(_hwnd, WinAPI.HWND_TOPMOST,
				 screenPos.X, screenPos.Y,
				 _width, _height,
				 WinAPI.SWP_NOACTIVATE Or WinAPI.SWP_SHOWWINDOW)

			' Now update the bitmap at the new position
			'RefreshLayer()
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
		Private Sub UpdateBitmapAndApply(screenPos As System.Drawing.Point)
			If _hwnd = IntPtr.Zero Then Return

			' ⭐ DO NOT RENDER UNTIL WE HAVE A REAL POSITION
			If Not _hasInitialPosition Then Return

			Dim size As New WinAPI.SIZE With {.cx = _width, .cy = _height}
			Dim dstPoint As New WinAPI.POINT With {.X = screenPos.X, .Y = screenPos.Y}
			Dim srcPoint As New WinAPI.POINT With {.X = 0, .Y = 0}

			Dim hdcScreen = WinAPI.GetDC(IntPtr.Zero)
			If hdcScreen = IntPtr.Zero Then Return

			Dim hdcMem = WinAPI.CreateCompatibleDC(hdcScreen)
			If hdcMem = IntPtr.Zero Then
				WinAPI.ReleaseDC(IntPtr.Zero, hdcScreen)
				Return
			End If

			' Create ARGB surface for GDI+
			Using bmp As New Bitmap(_width, _height, Imaging.PixelFormat.Format32bppArgb)
				Using g As Graphics = Graphics.FromImage(bmp)
					g.SmoothingMode = SmoothingMode.AntiAlias
					RenderToast(g)
				End Using

				Dim hBitmap As IntPtr = bmp.GetHbitmap(Color.FromArgb(0)) ' preserve alpha
				Dim oldObj = WinAPI.SelectObject(hdcMem, hBitmap)

				Dim blend As New WinAPI.BLENDFUNCTION() With {
			.BlendOp = WinAPI.AC_SRC_OVER,
			.BlendFlags = 0,
			.SourceConstantAlpha = _opacity,
			.AlphaFormat = WinAPI.AC_SRC_ALPHA
		}

				' Position

				Dim ok = WinAPI.UpdateLayeredWindow(_hwnd,
									 hdcScreen,
									 dstPoint,
									 size,
									 hdcMem,
									 srcPoint,
									 0,
									 blend,
									 WinAPI.ULW_ALPHA)

				' Cleanup
				WinAPI.SelectObject(hdcMem, oldObj)
				WinAPI.DeleteObject(hBitmap)
			End Using

			WinAPI.DeleteDC(hdcMem)
			WinAPI.ReleaseDC(IntPtr.Zero, hdcScreen)

		End Sub

		' ------------- Drawing ---------------------------
		Private Sub RenderToast(g As Graphics)
			Dim w = _width
			Dim h = _height

			Dim bgColor = _opts.BackColor
			Dim borderColor = _opts.BorderColor
			Dim foreColor = _opts.ForeColor
			Dim radius = _opts.CornerRadius

			' Clear entire bitmap to transparent
			g.Clear(Color.Transparent)

			' Rounded rect background
			Dim inset As Single = 0.5F
			Dim rect As New RectangleF(inset, inset, w - 1 - inset, h - 1 - inset)

			Using path As New GraphicsPath()
				path.AddArc(rect.X, rect.Y, radius, radius, 180, 90)
				path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90)
				path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90)
				path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90)
				path.CloseFigure()

				Using bgBrush As New SolidBrush(bgColor)
					g.FillPath(bgBrush, path)
				End Using

				Using pen As New Pen(borderColor, 1)
					g.DrawPath(pen, path)
				End Using
			End Using

			Dim textX As Integer = 10

			' Icon
			If _opts.Icon IsNot Nothing Then
				Dim maxSize As Integer = h - 20
				Dim bestIcon As Icon = _opts.Icon
				Try
					bestIcon = New Icon(_opts.Icon, New System.Drawing.Size(maxSize, maxSize))
				Catch
				End Try

				g.DrawIcon(bestIcon, New Rectangle(10, 10, maxSize, maxSize))
				textX = 20 + maxSize
			End If

			' Title
			Using brush As New SolidBrush(foreColor)
				Dim wAvail As Single = w - textX - 10
				If wAvail > 0 AndAlso Not String.IsNullOrEmpty(_opts.Title) Then
					Dim titleRect As New RectangleF(textX, 10, wAvail, 20)
					Dim fmt As New StringFormat With {.Trimming = StringTrimming.EllipsisCharacter}
					g.DrawString(_opts.Title, _opts.TitleFont, brush, titleRect, fmt)
				End If
			End Using

			' Message
			Using brush As New SolidBrush(foreColor)
				Dim messageRect As New RectangleF(
				textX,
				35,
				w - textX - 10,
				h - 45
			)

				Dim fmt As New StringFormat With {
				.Trimming = StringTrimming.EllipsisWord,
				.FormatFlags = StringFormatFlags.LineLimit,
				.Alignment = StringAlignment.Near,
				.LineAlignment = StringAlignment.Near
			}

				If Not String.IsNullOrEmpty(_opts.Message) Then
					g.DrawString(_opts.Message, _opts.MessageFont, brush, messageRect, fmt)
				End If
			End Using
		End Sub

	End Class

	' WinForms Window
	Public Class ToastWindow
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
		Public Sub ShowToastAt(p As System.Drawing.Point)
			_opacity = 0.0
			_fadingOut = False
			MyBase.ShowAt(p)
			FadeTimer.Start()
			LifeTimer.Start()
		End Sub
		Public Sub CloseToastWindow()
			FadeTimer.Stop()
			LifeTimer.Stop()
			_fadingOut = True
			FadeTimer.Start()
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
	End Class

#End Region

End Namespace
