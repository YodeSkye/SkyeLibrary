Namespace Skye.UI.Log

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class LogViewerControl
        Inherits System.Windows.Forms.UserControl

        'UserControl overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            components = New ComponentModel.Container()
            BtnPrevious = New Button()
            PanelTop = New Panel()
            BtnHighlightAll = New Button()
            BtnRefresh = New Button()
            BtnNext = New Button()
            TxtBoxSearch = New TextBox()
            RTBLog = New RichTextBox()
            Tip = New ToolTipEX(components)
            PanelTop.SuspendLayout()
            SuspendLayout()
            ' 
            ' BtnPrevious
            ' 
            Tip.SetImage(BtnPrevious, Resources.Resources.ImagePrevious32)
            BtnPrevious.Image = Resources.Resources.ImagePrevious32
            BtnPrevious.Location = New Point(199, 4)
            BtnPrevious.Margin = New Padding(4)
            BtnPrevious.Name = "BtnPrevious"
            BtnPrevious.Size = New Size(48, 48)
            BtnPrevious.TabIndex = 20
            Tip.SetText(BtnPrevious, "Previous")
            BtnPrevious.UseVisualStyleBackColor = True
            ' 
            ' PanelTop
            ' 
            PanelTop.Controls.Add(BtnHighlightAll)
            PanelTop.Controls.Add(BtnRefresh)
            PanelTop.Controls.Add(BtnNext)
            PanelTop.Controls.Add(TxtBoxSearch)
            PanelTop.Controls.Add(BtnPrevious)
            PanelTop.Dock = DockStyle.Top
            Tip.SetImage(PanelTop, Nothing)
            PanelTop.Location = New Point(0, 0)
            PanelTop.Margin = New Padding(4)
            PanelTop.Name = "PanelTop"
            PanelTop.Size = New Size(637, 57)
            PanelTop.TabIndex = 1
            Tip.SetText(PanelTop, Nothing)
            ' 
            ' BtnHighlightAll
            ' 
            Tip.SetImage(BtnHighlightAll, Resources.Resources.ImageHighlightAll32)
            BtnHighlightAll.Image = Resources.Resources.ImageHighlightAll32
            BtnHighlightAll.Location = New Point(311, 4)
            BtnHighlightAll.Margin = New Padding(4)
            BtnHighlightAll.Name = "BtnHighlightAll"
            BtnHighlightAll.Size = New Size(48, 48)
            BtnHighlightAll.TabIndex = 40
            Tip.SetText(BtnHighlightAll, "Highlight All Occurrences")
            BtnHighlightAll.UseVisualStyleBackColor = True
            ' 
            ' BtnRefresh
            ' 
            BtnRefresh.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Tip.SetImage(BtnRefresh, Resources.Resources.ImageRefreshLog32)
            BtnRefresh.Image = Resources.Resources.ImageRefreshLog32
            BtnRefresh.Location = New Point(585, 4)
            BtnRefresh.Margin = New Padding(4)
            BtnRefresh.Name = "BtnRefresh"
            BtnRefresh.Size = New Size(48, 48)
            BtnRefresh.TabIndex = 50
            Tip.SetText(BtnRefresh, "Refresh")
            BtnRefresh.UseVisualStyleBackColor = True
            ' 
            ' BtnNext
            ' 
            Tip.SetImage(BtnNext, Resources.Resources.ImageNext32)
            BtnNext.Image = Resources.Resources.ImageNext32
            BtnNext.Location = New Point(255, 4)
            BtnNext.Margin = New Padding(4)
            BtnNext.Name = "BtnNext"
            BtnNext.Size = New Size(48, 48)
            BtnNext.TabIndex = 30
            Tip.SetText(BtnNext, "Next")
            BtnNext.UseVisualStyleBackColor = True
            ' 
            ' TxtBoxSearch
            ' 
            Tip.SetImage(TxtBoxSearch, Nothing)
            TxtBoxSearch.Location = New Point(4, 14)
            TxtBoxSearch.Margin = New Padding(4)
            TxtBoxSearch.Name = "TxtBoxSearch"
            TxtBoxSearch.PlaceholderText = "Search"
            TxtBoxSearch.Size = New Size(187, 29)
            TxtBoxSearch.TabIndex = 10
            Tip.SetText(TxtBoxSearch, Nothing)
            ' 
            ' RTBLog
            ' 
            RTBLog.BorderStyle = BorderStyle.None
            RTBLog.Dock = DockStyle.Fill
            RTBLog.HideSelection = False
            Tip.SetImage(RTBLog, Nothing)
            RTBLog.Location = New Point(0, 57)
            RTBLog.Name = "RTBLog"
            RTBLog.ReadOnly = True
            RTBLog.Size = New Size(637, 372)
            RTBLog.TabIndex = 0
            RTBLog.TabStop = False
            Tip.SetText(RTBLog, Nothing)
            RTBLog.Text = ""
            RTBLog.WordWrap = False
            ' 
            ' Tip
            ' 
            Tip.Font = New Font("Segoe UI", 12.0F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
            Tip.ShadowAlpha = 0
            Tip.ShadowThickness = 0
            ' 
            ' LogViewerControl
            ' 
            AutoScaleDimensions = New SizeF(9.0F, 21.0F)
            AutoScaleMode = AutoScaleMode.Font
            Controls.Add(RTBLog)
            Controls.Add(PanelTop)
            DoubleBuffered = True
            Font = New Font("Segoe UI", 12.0F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
            Tip.SetImage(Me, Nothing)
            Margin = New Padding(4)
            Name = "LogViewerControl"
            Size = New Size(637, 429)
            Tip.SetText(Me, Nothing)
            PanelTop.ResumeLayout(False)
            PanelTop.PerformLayout()
            ResumeLayout(False)
        End Sub

        Friend WithEvents BtnPrevious As Button
        Friend WithEvents PanelTop As Panel
        Friend WithEvents TxtBoxSearch As TextBox
        Friend WithEvents BtnNext As Button
        Friend WithEvents BtnRefresh As Button
        Friend WithEvents RTBLog As RichTextBox
        Friend WithEvents BtnHighlightAll As Button
        Public WithEvents Tip As ToolTipEX

    End Class

End Namespace