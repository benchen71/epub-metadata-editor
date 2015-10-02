Public Class Form2
    Dim start As Integer
    Dim indexOfSearchText As Integer = 0
    Dim newsearch As Boolean = True
    Dim mytempfile As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        fileeditorreturn = False
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        fileeditorreturn = True
        filecontents = RichTextBox1.Text.Replace("�", "")
        Me.Close()
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
                If c = "�" Then cnt += 1
            Next
            CheckBox1.Checked = False
            RichTextBox1.SelectionStart = position - cnt
        End If
    End Sub

    Private Sub RichTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub RichTextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles RichTextBox1.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            e.IsInputKey = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        RichTextBox1.Text = Form1.Regularise(RichTextBox1.Text.Replace("�", ""))
        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text.Replace(Chr(10), "�" + Chr(10))
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text.Replace(Chr(10), "�" + Chr(10))
        Else
            RichTextBox1.Text = RichTextBox1.Text.Replace("�", "")
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Panel1.Visible Then
            Panel1.Visible = False
            Label2.Visible = False
            Button7.Text = "Go"
            'start = 0
            'indexOfSearchText = 0
            newsearch = True
            Dim position As Integer = RichTextBox1.SelectionStart
            RichTextBox1.SelectAll()
            RichTextBox1.SelectionBackColor = Color.Transparent
            RichTextBox1.Select(position, 0)
            RichTextBox1.Focus()
        Else
            Panel1.Visible = True
            txtSearch.Focus()
            start = RichTextBox1.SelectionStart
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' Code based on http://www.dotnetcurry.com/ShowArticle.aspx?ID=146
        Dim startindex As Integer
        Dim endindex As Integer
        Dim numentries As Integer

        ' Highlight all search strings in yellow
        If newsearch Then
            numentries = 0
            If txtSearch.Text.Length > 0 Then
                startindex = FindMyText(txtSearch.Text, 0, RichTextBox1.Text.Length)
                While startindex >= 0
                    numentries = numentries + 1
                    newsearch = False
                    RichTextBox1.SelectionBackColor = Color.Yellow
                    endindex = txtSearch.Text.Length
                    RichTextBox1.Select(startindex, endindex)
                    startindex = FindMyText(txtSearch.Text, startindex + endindex, RichTextBox1.Text.Length)
                End While
                indexOfSearchText = 0
                Label2.Text = numentries & " occurrences"
                Label2.Visible = True
            End If
        End If

        ' Highlight next search string in yellowgreen
        If txtSearch.Text.Length > 0 Then
            startindex = FindMyText(txtSearch.Text, start, RichTextBox1.Text.Length)
        End If

        ' If string was found in the RichTextBox, highlight it
        If startindex >= 0 Then
            RichTextBox1.SelectionBackColor = Color.YellowGreen
            endindex = txtSearch.Text.Length
            RichTextBox1.Select(startindex, endindex)
            RichTextBox1.ScrollToCaret()
            start = startindex + endindex
            Button7.Text = "Next"
        Else
            If InStr(RichTextBox1.Text, txtSearch.Text) Then
                start = 0
                indexOfSearchText = 0
                Button7_Click(sender, e)
            Else
                Beep()
                Button7.Text = "Go"
                newsearch = True
                indexOfSearchText = 0
                Label2.Visible = False
            End If
        End If
    End Sub
    Public Function FindMyText(ByVal txtToSearch As String, ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer
        ' Unselect the previously searched string
        If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
            RichTextBox1.SelectionBackColor = Color.Yellow

        End If

        ' Set the return value to -1 by default.
        Dim retVal As Integer = -1

        ' A valid starting index should be specified.
        ' if indexOfSearchText = -1, the end of search
        If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then
            ' A valid ending index
            If searchEnd > searchStart OrElse searchEnd = -1 Then
                ' Find the position of search string in RichTextBox
                indexOfSearchText = RichTextBox1.Find(txtToSearch, searchStart, searchEnd, RichTextBoxFinds.None)
                ' Determine whether the text was found in richTextBox1.
                If indexOfSearchText <> -1 Then
                    ' Return the index to the specified search text.
                    retVal = indexOfSearchText
                End If
            End If
        End If
        Return retVal
    End Function

    Private Sub txtSearch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Button7_Click(sender, e)
        End If
    End Sub

    Private Sub txtSearch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        If Not newsearch Then
            Dim position As Integer = RichTextBox1.SelectionStart
            RichTextBox1.SelectAll()
            RichTextBox1.SelectionBackColor = Color.Transparent
            RichTextBox1.Select(position, 0)
        End If
        Button7.Text = "Go"
        indexOfSearchText = 0
        newsearch = True
        Label2.Visible = False
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Button6_Click(sender, e)
    End Sub

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        txtSearch.Text = ""
        Button7.Text = "Go"
        start = 0
        indexOfSearchText = 0
        newsearch = True
        Panel1.Visible = False
        Label2.Visible = False
        If (System.IO.File.Exists(mytempfile)) Then System.IO.File.Delete(mytempfile)
        Form6.Button7.Visible = False
    End Sub

    Private Sub Form2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button3.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button4.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button5.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub CheckBox2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox2.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub CheckBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox1.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button6_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button6.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button7_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button7.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Public Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If (System.IO.File.Exists(mytempfile)) Then System.IO.File.Delete(mytempfile)
        mytempfile = IO.Path.GetDirectoryName(Form1.OpenFileDialog5.FileName) + "\mytempfile" + IO.Path.GetFileName(Form1.OpenFileDialog5.FileName)
        Form1.SaveUnicodeFile(mytempfile, RichTextBox1.Text)
        Dim myuri As New Uri(mytempfile)
        Form6.WebBrowser1.Url = myuri
        Form6.WebBrowser1.Refresh()
        Form6.Button7.Visible = True
        Form6.Text = "EPUB Metadata Editor - Preview of " + IO.Path.GetFileName(Form1.OpenFileDialog5.FileName)
        Form6.Show()
    End Sub

    Private Sub Button9_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button9.KeyUp
        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub
End Class