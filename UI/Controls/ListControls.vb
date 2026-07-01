
Imports System.ComponentModel

Namespace Skye.UI

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

End Namespace
