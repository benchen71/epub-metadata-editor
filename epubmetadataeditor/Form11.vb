Public Class Form11
    Dim charactersDisallowed As String = "/:*?<>|" + Chr(34)
    Dim SelectionText As String
    Dim ComboSelectionStart As Integer
    Dim ComboSelectionLength As Integer
    Dim firsttimeopen As Boolean = True
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim insertText = "%Creator%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim insertText = "%CreatorFileAs%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Dim insertText = "%Title%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel5_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Dim insertText = "%TitleFileAs%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel6_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Dim insertText = "%Date%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel7_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Dim insertText = "%Series%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub LinkLabel8_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Dim insertText = "%SeriesIndex%"
        Dim temp
        If Not ComboBox1.Text.Contains(insertText) Then
            MakeInsertion(insertText)
        Else
            temp = MsgBox("The Template can only contain one of each field.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
        End If
    End Sub

    Private Sub MakeInsertion(ByVal insertText)
        Dim insertPos As Integer = ComboSelectionStart
        Dim currentText As String = ComboBox1.Text
        If ComboSelectionLength <> 0 Then
            currentText = Mid(currentText, 1, ComboSelectionStart) + Mid(currentText, ComboSelectionStart + ComboSelectionLength + 1)
        End If
        If My.Computer.Keyboard.ShiftKeyDown Then
            ComboBox1.Text = currentText.Insert(insertPos, insertText.ToUpper)
        Else
            ComboBox1.Text = currentText.Insert(insertPos, insertText)
        End If

        ComboBox1.Focus()
        ComboBox1.SelectionStart = insertPos + insertText.Length
        ComboBox1.SelectionLength = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        UpdateMetadataPanel()
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If e.KeyChar = "%" Then
            Label5.Visible = True
        Else
            Label5.Visible = False
        End If
    End Sub

    Private Sub ComboBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyUp
        SelectionText = ComboBox1.SelectedText
        ComboSelectionStart = ComboBox1.SelectionStart
        ComboSelectionLength = ComboBox1.SelectionLength
    End Sub

    Private Sub ComboBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseUp
        SelectionText = ComboBox1.SelectedText
        ComboSelectionStart = ComboBox1.SelectionStart
        ComboSelectionLength = ComboBox1.SelectionLength
    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        Dim theText As String = ComboBox1.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = ComboBox1.SelectionStart
        Dim Change As Integer

        ' reset metadata preview
        DataGridView1.Rows.Clear()
        Button3.Enabled = False

        For x As Integer = 0 To ComboBox1.Text.Length - 1
            Letter = ComboBox1.Text.Substring(x, 1)
            If charactersDisallowed.Contains(Letter) Then
                If SelectionText = "" Then
                    theText = theText.Replace(Letter, String.Empty)
                    Change = 1
                    ComboBox1.Text = theText
                    ComboBox1.Select(SelectionIndex - Change, 0)
                Else
                    theText = theText.Replace(Letter, SelectionText)
                    Change = Len(SelectionText)
                    ComboBox1.Text = theText
                    ComboBox1.Select(SelectionIndex - 1, Change)
                End If
                Label4.Visible = True
                Label5.Visible = False
                Exit Sub
            End If
        Next
        Label4.Visible = False
    End Sub

    Private Sub Form4_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        ComboBox1.SelectionStart = 0
        ComboBox1.SelectionLength = 0
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label4.Text = "The following characters are not allowed:" + Chr(10) + charactersDisallowed
        Label5.Text = "To obtain a single '%'," + Chr(10) + "use '%%' in the Template."
        Label4.Visible = False
        Label5.Visible = False
        'TextBox3.Text = Mid(Form1.ListBox1.Items(0), InStrRev(Form1.ListBox1.Items(0), "\") + 1)
        'TextBox3.Text = Mid(TextBox3.Text, 1, Len(TextBox3.Text) - 5)
        TextBox3.Text = Form1.ListBox1.Items(0)
        Button3.Enabled = False
        If ComboBox1.Text <> "" Then
            UpdateMetadataPanel()
        End If
    End Sub

    Public Sub UpdateMetadataPanel()
        Dim slashloc, currposTemplate, currposFilename, endpos, endMetadata As Integer
        Dim currentField, currentSearchText, currentFilename, currentMetadata As String
        Dim foundfield As Boolean
        Dim temp
        currentSearchText = ""
        currentMetadata = ""
        currentField = ""
        currposTemplate = 0
        currposFilename = 1
        foundfield = False

        DataGridView1.Rows.Clear()

        ' check for (non-space) characters between fields


        Try
            ' get location of last "\"
            slashloc = InStrRev(Form1.ListBox1.Items(0), "\")

            ' include folders if Template contains "\" characters
            Dim numslash As Integer = Len(ComboBox1.Text) - Len(Replace(ComboBox1.Text, "\", ""))
            While numslash > 0
                slashloc = InStrRev(Form1.ListBox1.Items(0), "\", slashloc - 1)
                numslash = numslash - 1
            End While

            ' get filename
            currentFilename = Mid(Form1.ListBox1.Items(0), slashloc + 1)
            currentFilename = Mid(currentFilename, 1, Len(currentFilename) - 5)

            ' parse template
            While (currposTemplate < Len(ComboBox1.Text))
                currposTemplate = currposTemplate + 1

                ' look for field marker
                If (Mid(ComboBox1.Text, currposTemplate, 1) = "%") Then
                    If (Mid(ComboBox1.Text, currposTemplate + 1, 1) = "%") Then
                        ' found '%%' (replace with '%')
                        currentSearchText = currentSearchText + "%"
                        currposTemplate = currposTemplate + 1
                    Else
                        If foundfield Then
                            ' new field found so use currentSearchText to extract metadata for currentField
                            If Replace(currentSearchText, " ", "") = "" Then
                                temp = MsgBox("Oops! There has to be non-space characters between fields. Fix the Template and try again.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                                Exit Sub
                            End If
                            endMetadata = InStr(currposFilename, currentFilename, currentSearchText) - 1
                            currentMetadata = Mid(currentFilename, currposFilename, endMetadata - currposFilename + 1)

                            ' Add row to DataGridView
                            DataGridView1.Rows.Add(New String() {currentField, currentMetadata})
                            currposFilename = endMetadata + Len(currentSearchText) + 1
                            currentSearchText = ""
                        End If

                        ' look for end field marker
                        endpos = InStr(currposTemplate + 1, ComboBox1.Text, "%")
                        If (endpos <> 0) Then
                            ' end field marker found
                            currentField = Mid(ComboBox1.Text, currposTemplate + 1, endpos - currposTemplate - 1)
                            foundfield = True
                            currposTemplate = endpos
                        End If
                    End If
                Else
                    currentSearchText = currentSearchText + Mid(ComboBox1.Text, currposTemplate, 1)
                End If
            End While
            ' Get metadata for last field
            currentMetadata = Mid(currentFilename, currposFilename)

            ' Add last row to DataGridView
            DataGridView1.Rows.Add(New String() {currentField, currentMetadata})

            ' If we got this far, then extraction can procede!
            Button3.Enabled = True
        Catch
            If Not firsttimeopen Then
                temp = MsgBox("Oops! I can't process that. Fix the Template and try again.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            End If            
        End Try
        firsttimeopen = False
    End Sub
End Class