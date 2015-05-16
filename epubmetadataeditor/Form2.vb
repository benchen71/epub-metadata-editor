Public Class Form2

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fileeditorreturn = False
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        fileeditorreturn = True
        filecontents = RichTextBox1.Text
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Clipboard.SetText("<guide>" + Chr(10) + Chr(9) + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(10) + "</guide>")
    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If (DirectCast(Me.ActiveControl, RichTextBox).CanUndo) Then
            ContextMenuStrip1.Items(0).Enabled = True
        Else
            ContextMenuStrip1.Items(0).Enabled = False
        End If

        If (DirectCast(Me.ActiveControl, RichTextBox).SelectedText.Length = 0) Then
            ContextMenuStrip1.Items(2).Enabled = False
            ContextMenuStrip1.Items(3).Enabled = False
        Else
            ContextMenuStrip1.Items(2).Enabled = True
            ContextMenuStrip1.Items(3).Enabled = True
        End If

        If (Clipboard.ContainsText()) Then
            ContextMenuStrip1.Items(4).Enabled = True
        Else
            ContextMenuStrip1.Items(4).Enabled = False
        End If

        If (DirectCast(Me.ActiveControl, RichTextBox).Text.Length = 0) Then
            ContextMenuStrip1.Items(6).Enabled = False
        Else
            ContextMenuStrip1.Items(6).Enabled = True
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        DirectCast(Me.ActiveControl, RichTextBox).Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        DirectCast(Me.ActiveControl, RichTextBox).Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        DirectCast(Me.ActiveControl, RichTextBox).Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        DirectCast(Me.ActiveControl, RichTextBox).SelectAll()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        DirectCast(Me.ActiveControl, RichTextBox).Undo()
    End Sub

    Private Sub RichTextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles RichTextBox1.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            e.IsInputKey = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim opffiletext As String

        opffiletext = RichTextBox1.Text

        ' delete stuff
        opffiletext = opffiletext.Replace(Chr(13), "")
        opffiletext = opffiletext.Replace(Chr(10), "")
        opffiletext = opffiletext.Replace(Chr(9), "")
        While (opffiletext.Contains("> "))
            opffiletext = opffiletext.Replace("> ", ">")
        End While

        ' add stuff back
        opffiletext = opffiletext.Replace("><", ">" + Chr(10) + "<")
        opffiletext = opffiletext.Replace("<metadata", Chr(9) + "<metadata")
        opffiletext = opffiletext.Replace("</metadata", Chr(9) + "</metadata")
        opffiletext = opffiletext.Replace("<manifest", Chr(9) + "<manifest")
        opffiletext = opffiletext.Replace("</manifest", Chr(9) + "</manifest")
        opffiletext = opffiletext.Replace("<spine", Chr(9) + "<spine")
        opffiletext = opffiletext.Replace("</spine", Chr(9) + "</spine")
        opffiletext = opffiletext.Replace("<guide", Chr(9) + "<guide")
        opffiletext = opffiletext.Replace("</guide", Chr(9) + "</guide")
        opffiletext = opffiletext.Replace("<dc:", Chr(9) + Chr(9) + "<dc:")
        opffiletext = opffiletext.Replace("<meta ", Chr(9) + Chr(9) + "<meta ")
        opffiletext = opffiletext.Replace("<item", Chr(9) + Chr(9) + "<item")
        opffiletext = opffiletext.Replace("<reference", Chr(9) + Chr(9) + "<reference")
        RichTextBox1.Text = opffiletext
    End Sub
End Class