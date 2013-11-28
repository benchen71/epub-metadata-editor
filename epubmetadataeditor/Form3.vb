Public Class Form3
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Me.Disposed
        Form1.Button23.Enabled = True
        Me.Hide()
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        If ToolStripMenuItem1.Checked = True Then
            VistaTreeView1.CollapseAll()
            ToolStripMenuItem1.Checked = False
        Else
            VistaTreeView1.ExpandAll()
            ToolStripMenuItem1.Checked = True
        End If
    End Sub
End Class