Namespace UI.Log

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class LogViewer
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
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
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LogViewer))
            LogViewerControl1 = New LogViewerControl()
            SuspendLayout()
            ' 
            ' LogViewerControl1
            ' 
            LogViewerControl1.Dock = DockStyle.Fill
            LogViewerControl1.Font = New Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
            LogViewerControl1.Location = New Point(0, 0)
            LogViewerControl1.Margin = New Padding(4)
            LogViewerControl1.Name = "LogViewerControl1"
            LogViewerControl1.Size = New Size(800, 450)
            LogViewerControl1.TabIndex = 0
            ' 
            ' LogViewer
            ' 
            AutoScaleDimensions = New SizeF(7F, 15F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New Size(800, 450)
            Controls.Add(LogViewerControl1)
            Icon = CType(resources.GetObject("$this.Icon"), Icon)
            Name = "LogViewer"
            Text = "Log Viewer"
            ResumeLayout(False)
        End Sub

        Friend WithEvents LogViewerControl1 As LogViewerControl
    End Class

End Namespace
