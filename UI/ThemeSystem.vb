
Namespace UI

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

		' Textboxes & RTB
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

		' Menus
		Public Property MenuBack As Color
		Public Property MenuFore As Color
		Public Property MenuHover As Color
		Public Property MenuBorder As Color
		Public Property MenuSeparator As Color

		' Tooltips
		Public Property TooltipBack As Color
		Public Property TooltipFore As Color
		Public Property TooltipBorder As Color

	End Class

	Public Module SkyeThemes

		Public ReadOnly Light As New SkyeTheme With {
			.Name = "Light",
			.BackColor = Color.White,
			.ForeColor = Color.Black,
			.AccentColor = Color.DeepSkyBlue,
			.BorderColor = Color.LightGray,
			.ButtonBack = Color.Gainsboro,
			.ButtonFore = Color.Black,
			.TextBack = Color.White,
			.TextFore = Color.Black,
			.GroupBoxFore = Color.Black,
			.GridBack = Color.White,
			.GridFore = Color.Black,
			.GridHeaderBack = Color.Gainsboro,
			.GridHeaderFore = Color.Black,
			.GridBorder = Color.LightGray,
			.GridAlternateRowBack = Color.FromArgb(245, 245, 245),
			.TooltipBack = Color.White,
			.TooltipFore = Color.Black,
			.TooltipBorder = Color.LightGray,
			.MenuBack = Color.White,
			.MenuFore = Color.Black,
			.MenuHover = Color.FromArgb(230, 230, 230),
			.MenuBorder = Color.LightGray,
			.MenuSeparator = Color.LightGray
		}
		Public ReadOnly Dark As New SkyeTheme With {
			.Name = "Dark",
			.BackColor = Color.Black,
			.ForeColor = Color.White,
			.AccentColor = Color.DeepSkyBlue,
			.BorderColor = Color.FromArgb(64, 64, 64),
			.ButtonBack = Color.FromArgb(45, 45, 45),
			.ButtonFore = Color.White,
			.TextBack = Color.FromArgb(40, 40, 40),
			.TextFore = Color.White,
			.GroupBoxFore = Color.White,
			.GridBack = Color.Black,
			.GridFore = Color.White,
			.GridHeaderBack = Color.FromArgb(32, 32, 32),
			.GridHeaderFore = Color.White,
			.GridBorder = Color.FromArgb(70, 70, 70),
			.GridAlternateRowBack = Color.FromArgb(40, 40, 40),
			.TooltipBack = Color.Black,
			.TooltipFore = Color.White,
			.TooltipBorder = Color.FromArgb(80, 80, 80),
			.MenuBack = Color.Black,
			.MenuFore = Color.White,
			.MenuHover = Color.FromArgb(60, 60, 60),
			.MenuBorder = Color.FromArgb(80, 80, 80),
			.MenuSeparator = Color.FromArgb(90, 90, 90)
		}
		Public ReadOnly Slate As New SkyeTheme With {
			.Name = "Slate",
			.BackColor = Color.FromArgb(245, 245, 245),
			.ForeColor = Color.FromArgb(40, 40, 40),
			.AccentColor = Color.FromArgb(100, 140, 180),
			.BorderColor = Color.FromArgb(200, 200, 200),
			.ButtonBack = Color.FromArgb(230, 230, 230),
			.ButtonFore = Color.FromArgb(40, 40, 40),
			.TextBack = Color.White,
			.TextFore = Color.FromArgb(40, 40, 40),
			.GroupBoxFore = Color.FromArgb(40, 40, 40),
			.GridBack = Color.White,
			.GridFore = Color.FromArgb(40, 40, 40),
			.GridHeaderBack = Color.FromArgb(225, 225, 225),
			.GridHeaderFore = Color.FromArgb(40, 40, 40),
			.GridBorder = Color.FromArgb(200, 200, 200),
			.GridAlternateRowBack = Color.FromArgb(240, 240, 240),
			.TooltipBack = Color.FromArgb(245, 245, 245),
			.TooltipFore = Color.FromArgb(40, 40, 40),
			.TooltipBorder = Color.FromArgb(200, 200, 200),
			.MenuBack = Color.FromArgb(245, 245, 245),
			.MenuFore = Color.FromArgb(40, 40, 40),
			.MenuHover = Color.FromArgb(230, 230, 230),
			.MenuBorder = Color.FromArgb(200, 200, 200),
			.MenuSeparator = Color.FromArgb(200, 200, 200)
		}
		Public ReadOnly Graphite As New SkyeTheme With {
			.Name = "Graphite",
			.BackColor = Color.FromArgb(25, 25, 25),
			.ForeColor = Color.FromArgb(230, 230, 230),
			.AccentColor = Color.FromArgb(120, 160, 200),
			.BorderColor = Color.FromArgb(70, 70, 70),
			.ButtonBack = Color.FromArgb(40, 40, 40),
			.ButtonFore = Color.FromArgb(230, 230, 230),
			.TextBack = Color.FromArgb(40, 40, 40),
			.TextFore = Color.FromArgb(230, 230, 230),
			.GroupBoxFore = Color.FromArgb(230, 230, 230),
			.GridBack = Color.FromArgb(30, 30, 30),
			.GridFore = Color.FromArgb(230, 230, 230),
			.GridHeaderBack = Color.FromArgb(50, 50, 50),
			.GridHeaderFore = Color.FromArgb(230, 230, 230),
			.GridBorder = Color.FromArgb(70, 70, 70),
			.GridAlternateRowBack = Color.FromArgb(35, 35, 35),
			.TooltipBack = Color.FromArgb(40, 40, 40),
			.TooltipFore = Color.FromArgb(230, 230, 230),
			.TooltipBorder = Color.FromArgb(70, 70, 70),
			.MenuBack = Color.FromArgb(30, 30, 30),
			.MenuFore = Color.FromArgb(230, 230, 230),
			.MenuHover = Color.FromArgb(50, 50, 50),
			.MenuBorder = Color.FromArgb(70, 70, 70),
			.MenuSeparator = Color.FromArgb(70, 70, 70)
		}
		Public ReadOnly HighContrast As New SkyeTheme With {
			.Name = "High Contrast",
			.BackColor = Color.Black,
			.ForeColor = Color.White,
			.AccentColor = Color.Yellow,
			.BorderColor = Color.White,
			.ButtonBack = Color.Black,
			.ButtonFore = Color.White,
			.TextBack = Color.Black,
			.TextFore = Color.White,
			.GroupBoxFore = Color.White,
			.GridBack = Color.Black,
			.GridFore = Color.White,
			.GridHeaderBack = Color.White,
			.GridHeaderFore = Color.Black,
			.GridBorder = Color.White,
			.GridAlternateRowBack = Color.FromArgb(32, 32, 32),
			.TooltipBack = Color.Black,
			.TooltipFore = Color.White,
			.TooltipBorder = Color.White,
			.MenuBack = Color.Black,
			.MenuFore = Color.White,
			.MenuHover = Color.Yellow,
			.MenuBorder = Color.White,
			.MenuSeparator = Color.White
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
			.MenuBack = Color.FromArgb(255, 35, 35, 35),
			.MenuFore = Color.DeepPink,
			.MenuHover = Color.HotPink,
			.MenuBorder = Color.DeepPink,
			.MenuSeparator = Color.FromArgb(255, 90, 90, 90)
		}
		Public ReadOnly Sky As New SkyeTheme With {
			.Name = "Sky",
			.BackColor = Color.FromArgb(235, 245, 255),
			.ForeColor = Color.FromArgb(20, 40, 60),
			.AccentColor = Color.FromArgb(70, 140, 220),
			.BorderColor = Color.FromArgb(180, 200, 220),
			.ButtonBack = Color.FromArgb(220, 235, 250),
			.ButtonFore = Color.FromArgb(20, 40, 60),
			.TextBack = Color.FromArgb(235, 245, 255),
			.TextFore = Color.FromArgb(20, 40, 60),
			.GroupBoxFore = Color.FromArgb(20, 40, 60),
			.GridBack = Color.White,
			.GridFore = Color.FromArgb(20, 40, 60),
			.GridHeaderBack = Color.FromArgb(210, 225, 245),
			.GridHeaderFore = Color.FromArgb(20, 40, 60),
			.GridBorder = Color.FromArgb(180, 200, 220),
			.GridAlternateRowBack = Color.FromArgb(245, 250, 255),
			.TooltipBack = Color.FromArgb(235, 245, 255),
			.TooltipFore = Color.FromArgb(20, 40, 60),
			.TooltipBorder = Color.FromArgb(180, 200, 220),
			.MenuBack = Color.FromArgb(235, 245, 255),
			.MenuFore = Color.FromArgb(20, 40, 60),
			.MenuHover = Color.FromArgb(220, 235, 250),
			.MenuBorder = Color.FromArgb(180, 200, 220),
			.MenuSeparator = Color.FromArgb(180, 200, 220)
		}
		Public ReadOnly MidnightBlue As New SkyeTheme With {
			.Name = "Midnight Blue",
			.BackColor = Color.FromArgb(15, 25, 40),
			.ForeColor = Color.FromArgb(220, 235, 255),
			.AccentColor = Color.FromArgb(90, 160, 255),
			.BorderColor = Color.FromArgb(60, 80, 110),
			.ButtonBack = Color.FromArgb(25, 40, 60),
			.ButtonFore = Color.FromArgb(220, 235, 255),
			.TextBack = Color.FromArgb(25, 40, 60),
			.TextFore = Color.FromArgb(220, 235, 255),
			.GroupBoxFore = Color.FromArgb(220, 235, 255),
			.GridBack = Color.FromArgb(20, 30, 45),
			.GridFore = Color.FromArgb(220, 235, 255),
			.GridHeaderBack = Color.FromArgb(35, 50, 75),
			.GridHeaderFore = Color.FromArgb(220, 235, 255),
			.GridBorder = Color.FromArgb(60, 80, 110),
			.GridAlternateRowBack = Color.FromArgb(25, 35, 55),
			.TooltipBack = Color.FromArgb(25, 40, 60),
			.TooltipFore = Color.FromArgb(220, 235, 255),
			.TooltipBorder = Color.FromArgb(60, 80, 110),
			.MenuBack = Color.FromArgb(20, 30, 45),
			.MenuFore = Color.FromArgb(220, 235, 255),
			.MenuHover = Color.FromArgb(35, 50, 75),
			.MenuBorder = Color.FromArgb(60, 80, 110),
			.MenuSeparator = Color.FromArgb(60, 80, 110)
		}
		Public ReadOnly Mint As New SkyeTheme With {
			.Name = "Mint",
			.BackColor = Color.FromArgb(235, 250, 240),
			.ForeColor = Color.FromArgb(25, 60, 45),
			.AccentColor = Color.FromArgb(80, 180, 140),
			.BorderColor = Color.FromArgb(180, 210, 195),
			.ButtonBack = Color.FromArgb(220, 240, 230),
			.ButtonFore = Color.FromArgb(25, 60, 45),
			.TextBack = Color.FromArgb(235, 250, 240),
			.TextFore = Color.FromArgb(25, 60, 45),
			.GroupBoxFore = Color.FromArgb(25, 60, 45),
			.GridBack = Color.White,
			.GridFore = Color.FromArgb(25, 60, 45),
			.GridHeaderBack = Color.FromArgb(210, 235, 225),
			.GridHeaderFore = Color.FromArgb(25, 60, 45),
			.GridBorder = Color.FromArgb(180, 210, 195),
			.GridAlternateRowBack = Color.FromArgb(240, 250, 245),
			.TooltipBack = Color.FromArgb(235, 250, 240),
			.TooltipFore = Color.FromArgb(25, 60, 45),
			.TooltipBorder = Color.FromArgb(180, 210, 195),
			.MenuBack = Color.FromArgb(235, 250, 240),
			.MenuFore = Color.FromArgb(25, 60, 45),
			.MenuHover = Color.FromArgb(220, 240, 230),
			.MenuBorder = Color.FromArgb(180, 210, 195),
			.MenuSeparator = Color.FromArgb(180, 210, 195)
		}
		Public ReadOnly Evergreen As New SkyeTheme With {
			.Name = "Evergreen",
			.BackColor = Color.FromArgb(15, 35, 25),
			.ForeColor = Color.FromArgb(220, 245, 230),
			.AccentColor = Color.FromArgb(90, 200, 150),
			.BorderColor = Color.FromArgb(60, 100, 80),
			.ButtonBack = Color.FromArgb(25, 55, 40),
			.ButtonFore = Color.FromArgb(220, 245, 230),
			.TextBack = Color.FromArgb(25, 55, 40),
			.TextFore = Color.FromArgb(220, 245, 230),
			.GroupBoxFore = Color.FromArgb(220, 245, 230),
			.GridBack = Color.FromArgb(20, 45, 30),
			.GridFore = Color.FromArgb(220, 245, 230),
			.GridHeaderBack = Color.FromArgb(35, 70, 50),
			.GridHeaderFore = Color.FromArgb(220, 245, 230),
			.GridBorder = Color.FromArgb(60, 100, 80),
			.GridAlternateRowBack = Color.FromArgb(25, 60, 40),
			.TooltipBack = Color.FromArgb(25, 55, 40),
			.TooltipFore = Color.FromArgb(220, 245, 230),
			.TooltipBorder = Color.FromArgb(60, 100, 80),
			.MenuBack = Color.FromArgb(20, 45, 30),
			.MenuFore = Color.FromArgb(220, 245, 230),
			.MenuHover = Color.FromArgb(35, 70, 50),
			.MenuBorder = Color.FromArgb(60, 100, 80),
			.MenuSeparator = Color.FromArgb(60, 100, 80)
		}
		Private ReadOnly _themes As New List(Of SkyeTheme) From {Light, Dark, HighContrast, Slate, Graphite, Blossom, CrimsonNight, Sky, MidnightBlue, Mint, Evergreen}
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

		Private ReadOnly _registeredComponents As New List(Of Object)

		Public Property CurrentTheme As SkyeTheme = SkyeThemes.Dark
		Public Event ThemeChanged As EventHandler

		Public Sub RegisterComponent(comp As Object)
			If comp Is Nothing Then Return
			_registeredComponents.Add(comp)
		End Sub
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

			' Theme registered components (ToolTip, ToolTipEX, etc.)
			For Each comp In _registeredComponents
				Select Case True
					Case TypeOf comp Is System.Windows.Forms.ToolTip
						ApplyToTooltip(DirectCast(comp, System.Windows.Forms.ToolTip))
					Case TypeOf comp Is ToolTip
						ApplyToTooltip(DirectCast(comp, ToolTip))
					Case TypeOf comp Is ToolTipEX
						ApplyToTooltip(DirectCast(comp, ToolTipEX))
					Case Else
						' Future: NotifyIcon, ContextMenuStrip, ImageList, etc.
				End Select
			Next

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
		' Theme a tooltip
		Public Sub ApplyToTooltip(tip As System.Windows.Forms.ToolTip)
			If tip Is Nothing Then Return
			tip.BackColor = CurrentTheme.TooltipBack
			tip.ForeColor = CurrentTheme.TooltipFore
		End Sub
		Public Sub ApplyToTooltip(tip As ToolTip)
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

End Namespace
