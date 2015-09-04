Public Class Form2

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fileeditorreturn = False
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        fileeditorreturn = True
        filecontents = RichTextBox1.Text.Replace("¶", "")
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

    Private Sub RichTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RichTextBox1.KeyPress
        If ((e.KeyChar = Chr(13)) And (CheckBox1.Checked)) Then
            Dim position = RichTextBox1.SelectionStart
            Dim cnt As Integer = 0
            For Each c As Char In Mid(RichTextBox1.Text, 1, position)
                If c = "¶" Then cnt += 1
            Next
            CheckBox1.Checked = False
            RichTextBox1.SelectionStart = position - cnt
        End If
    End Sub

    Private Sub RichTextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles RichTextBox1.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            e.IsInputKey = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        RichTextBox1.Text = Form1.Regularise(RichTextBox1.Text.Replace("¶", ""))
        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text.Replace(Chr(10), "¶" + Chr(10))
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text.Replace(Chr(10), "¶" + Chr(10))
        Else
            RichTextBox1.Text = RichTextBox1.Text.Replace("¶", "")
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim metadatafile, newmetadatafile As String
        Dim prevpos, startpos, endpos As Integer
        metadatafile = RichTextBox1.Text
        prevpos = 1
        newmetadatafile = ""
        startpos = InStr(metadatafile, "<meta ")
        While (startpos <> 0)
            endpos = InStr(startpos, metadatafile, "/>")
            Dim tempstring = Mid(metadatafile, startpos, endpos - startpos).ToLower
            If (InStr(Mid(metadatafile, startpos, endpos - startpos).ToLower, "name=" + Chr(34) + "cover") = 0) Then
                newmetadatafile = newmetadatafile + Mid(metadatafile, prevpos, startpos - prevpos - 1)
                prevpos = endpos + 2
            Else
                If Not CheckBox2.Checked Then
                    newmetadatafile = newmetadatafile + Mid(metadatafile, prevpos, startpos - prevpos - 1)
                    prevpos = endpos + 2
                End If
            End If
            startpos = InStr(startpos + 1, metadatafile, "<meta ")
        End While
        newmetadatafile = newmetadatafile + Mid(metadatafile, prevpos)
        RichTextBox1.Text = newmetadatafile
    End Sub
End Class