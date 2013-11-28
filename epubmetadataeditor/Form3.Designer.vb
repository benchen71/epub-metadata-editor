<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.Button1 = New System.Windows.Forms.Button
        Me.VistaTreeView1 = New EPubMetadataEditor.VistaTreeView
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(130, 26)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Checked = True
        Me.ToolStripMenuItem1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(129, 22)
        Me.ToolStripMenuItem1.Text = "Expand All"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(147, 617)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(114, 26)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'VistaTreeView1
        '
        Me.VistaTreeView1.BackColor = System.Drawing.Color.White
        Me.VistaTreeView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.VistaTreeView1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.VistaTreeView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VistaTreeView1.FontColorSelected = System.Drawing.Color.Black
        Me.VistaTreeView1.FontHotTracking = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VistaTreeView1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.VistaTreeView1.ItemHeight = 17
        Me.VistaTreeView1.Location = New System.Drawing.Point(9, 12)
        Me.VistaTreeView1.Name = "VistaTreeView1"
        Me.VistaTreeView1.NodeBaseColor = System.Drawing.Color.White
        Me.VistaTreeView1.NodeColor = System.Drawing.Color.White
        Me.VistaTreeView1.Size = New System.Drawing.Size(384, 597)
        Me.VistaTreeView1.TabIndex = 2
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 649)
        Me.Controls.Add(Me.VistaTreeView1)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form3"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "EPub Metadata Editor - Table Of Contents"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VistaTreeView1 As EPubMetadataEditor.VistaTreeView
End Class
