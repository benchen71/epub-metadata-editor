Imports System
Imports System.IO
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Win32
Imports System.Drawing.Imaging


Public Class Form1
    Dim projectchanged As Boolean
    Dim tempdirectory As String
    Dim ebookdirectory As String
    Dim opfdirectory As String
    Dim coverimagefile As String
    Dim relativecoverimagefile As String
    Dim coverimagefilename As String
    Dim coverfile As String
    Dim opffile As String
    Dim tocfile As String
    Dim pagemapfile As String
    Dim m_MouseIsDown As Boolean
    Dim appdatafolder As String
    Dim versioninfo As String
    Dim tocncxfile As String
    Dim updateinfo As String
    Dim fixcovermetadata As Boolean
    Dim fixcovermanifest As Boolean
    Dim CaptionString As String
    Dim searchResults As String()
    Dim currentfilenumber As Integer
    Dim refreshfilelist As Boolean = True
    Dim possibleDRM As Boolean = False
    Dim keepcombobox As Boolean = False
    Dim subjectseparator As String
    Dim WordsNotToCapitalise As String
    Dim idcount As Integer

    Public Sub DeleteDirContents(ByVal dir As IO.DirectoryInfo)
        Dim fa() As IO.FileInfo
        Dim f As IO.FileInfo

        fa = dir.GetFiles

        For Each f In fa
            f.Delete()
        Next

        Dim da() As IO.DirectoryInfo
        Dim d1 As IO.DirectoryInfo

        da = dir.GetDirectories
        For Each d1 In da
            DeleteDirContents(d1)
        Next
    End Sub

    Private Sub ClearInterface()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        WebBrowser1.DocumentText = ""
        WebBrowser1.Visible = False
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        projectchanged = False
        refreshfilelist = True
        If keepcombobox = False Then
            ComboBox3.SelectedIndex = -1
            CaptionString = "EPUB Metadata Editor"
            Me.Text = CaptionString
        End If
        PictureBox1.Image = Nothing
        Label4.Visible = False
        Button1.Visible = False
        Button42.Visible = False
        Button35.Visible = False
        Button27.Visible = False
        Label25.Visible = False
        GroupBox1.Visible = False
        ListBox2.Items.Clear()
        Label27.Visible = False
        Label23.Visible = False
        CheckBox5.Visible = False
        Label31.Visible = False
    End Sub

    Private Sub SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, _
    TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, TextBox6.TextChanged, TextBox7.TextChanged, TextBox8.TextChanged, _
    TextBox9.TextChanged, TextBox10.TextChanged, TextBox11.TextChanged, TextBox14.TextChanged, TextBox15.TextChanged, TextBox16.TextChanged, TextBox17.TextChanged
        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString

        If ComboBox1.SelectedIndex = -1 Then ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString

        If ComboBox2.SelectedIndex = -1 Then ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim MyFiles() As String
        Dim i As Integer
        Dim ext As String

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Assign the file(s) to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
        Else
            Exit Sub
        End If

        ' Loop through the array
        For i = 0 To MyFiles.Length - 1
            ext = Path.GetExtension(MyFiles(i)).ToLower()
            If (ext = ".epub") Then
                If (MyFiles(i) <> OpenFileDialog1.FileName) Then
                    ' Different file already open
                    If projectchanged Then
                        DialogResult = Dialog1.ShowDialog
                        If DialogResult = Windows.Forms.DialogResult.Cancel Then
                            Exit Sub
                        End If
                        If DialogResult = Windows.Forms.DialogResult.Yes Then
                            SaveEpub(OpenFileDialog1.FileName, False)
                        End If
                    End If

                    ' Delete previous temp directory (if it exists)
                    If tempdirectory <> "" Then
                        ChDir(tempdirectory)
                        If ebookdirectory <> "" Then
                            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                                Try
                                    My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                                Catch ex As Exception
                                    Dim instance = Convert.ToInt16(Mid(ebookdirectory, InStrRev(ebookdirectory, "B" + 1)))
                                    instance = instance + 1
                                    ebookdirectory = tempdirectory + "EPUB" + Trim(Str(instance))
                                End Try
                            End If
                        End If
                    End If

                    ClearInterface()
                    SaveImageAsToolStripMenuItem.Enabled = False
                    AddImageToolStripMenuItem.Enabled = False
                    ChangeImageToolStripMenuItem.Enabled = False
                    OpenFileDialog1.FileName = MyFiles(i)
                    OpenEPub()
                    Button3.Enabled = False

                    ' Put cursor in Title box
                    TextBox1.Focus()
                End If
            End If
        Next
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If projectchanged Then
            DialogResult = Dialog1.ShowDialog()
            If DialogResult = Windows.Forms.DialogResult.No Then
                ChDir(tempdirectory)
                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                    End Try
                End If
                End
            ElseIf DialogResult = Windows.Forms.DialogResult.Yes Then
                SaveEpub(OpenFileDialog1.FileName, False)
            Else
                e.Cancel = True
            End If
        Else
            ChDir(tempdirectory)
            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                Try
                    My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception
                End Try
            End If
            End
        End If
    End Sub
    Private Function LoadUnicodeFile(ByVal filename) As String
        Dim UnicodeText As String
        Dim sr As StreamReader = Nothing
        UnicodeText = ""
        Try
            sr = OpenText(filename)
            UnicodeText = sr.ReadToEnd
            sr.Close()
            sr.Dispose()
        Catch
            'Always check to make sure the object isnt nothing (to avoid nullreference exceptions)
            If Not sr Is Nothing Then
                sr.Close()
                sr = Nothing
            End If
        End Try
        Return UnicodeText
    End Function
    Public Sub SaveUnicodeFile(ByVal filename, ByVal UnicodeText)
        Dim sw As StreamWriter = Nothing
        Try
            sw = New StreamWriter(filename, False)
            sw.WriteLine(UnicodeText)
            sw.Close()
            sw.Dispose()
        Catch
            If Not sw Is Nothing Then
                sw.Close()
                sw = Nothing
            End If
        End Try
    End Sub
    Public Sub OpenEPUB()
        Dim metadatafile As String
        Dim instance As Integer

        'Unzip epub to temp directory
        tempdirectory = System.IO.Path.GetTempPath
        instance = 1
        ebookdirectory = tempdirectory + "EPUB" + Trim(Str(instance))
        While (My.Computer.FileSystem.DirectoryExists(ebookdirectory))
            Try
                My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.ThrowIfDirectoryNonEmpty)
                GoTo founddirectory
            Catch ex As Exception
                instance = instance + 1
                ebookdirectory = tempdirectory + "EPUB" + Trim(Str(instance))
            End Try
        End While
founddirectory:
        MkDir(ebookdirectory)
        ChDir(ebookdirectory)

        Try
            Dim zip As ZipStorer
            zip = ZipStorer.Open(OpenFileDialog1.FileName, FileAccess.Read)
            Dim dir = zip.ReadCentralDir()
            Dim item As ZipStorer.ZipFileEntry
            For Each item In dir
                zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
            Next
            zip.Close()
        Catch ex1 As Exception
            Console.Error.WriteLine("exception: {0}", ex1.ToString)
            DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "This ebook cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Exit Sub
        End Try

        'Search for .opf file
        searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

        'Open .opf file into RichTextBox
        If searchResults.Length < 1 Then
            DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "This ebook is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Button19.Enabled = False
            opffile = ""
            Return
        Else
            opffile = searchResults(0)
            If InStr(opffile, "_MACOSX") Then
                If searchResults.Length > 1 Then
                    opffile = searchResults(1)
                End If
            End If
            opfdirectory = Path.GetDirectoryName(opffile)
            RichTextBox1.Text = LoadUnicodeFile(opffile)
            Button19.Enabled = True
        End If

        'Process .opf file to determine EPUB version
        Dim opffiletext As String
        Dim packagepos, endpos, versionpos As Integer
        opffiletext = LoadUnicodeFile(opffile)
        packagepos = InStr(opffiletext, "<package")
        If packagepos <> 0 Then
            endpos = InStr(packagepos, opffiletext, ">")
            versionpos = InStr(packagepos, opffiletext, "version=")
            If versionpos < endpos Then
                versioninfo = Mid(opffiletext, versionpos + 9, 3)
            End If
        End If

        If versioninfo = "3.0" Then
            'Search for toc.ncx file (included in some EPUB3 files for forward compatibility)
            searchResults = Directory.GetFiles(ebookdirectory, "*.ncx", SearchOption.AllDirectories)
            If searchResults.Length < 1 Then
                Button34.Visible = False
                LinkLabel5.Visible = False
                tocncxfile = ""
            Else
                tocncxfile = searchResults(0)
                If InStr(tocncxfile, "_MACOSX") Then
                    If searchResults.Length > 1 Then
                        tocncxfile = searchResults(1)
                    End If
                End If
                Button34.Visible = True
                LinkLabel5.Visible = True
            End If

            'Search for nav file
            Dim itempos, tocpos, hrefpos, itemend As Integer
            itempos = InStr(opffiletext, "<item ")
            While itempos <> 0
                itemend = InStr(itempos, opffiletext, "/>")
                tocpos = InStr(opffiletext, "properties=" + Chr(34) + "nav" + Chr(34))
                If tocpos <> 0 Then
                    If tocpos < itemend Then
                        'Found nav file
                        hrefpos = InStr(itempos, opffiletext, "href=")
                        endpos = InStr(hrefpos + 6, opffiletext, Chr(34))
                        tocfile = opfdirectory + "\" + Mid(opffiletext, hrefpos + 6, endpos - hrefpos - 6).Replace("/", "\")
                        Button20.Enabled = True
                        Button23.Enabled = True
                        Button20.Text = "Edit nav file"
                        GoTo lookforpagemap
                    Else
                        'Go to next item
                        itempos = InStr(itemend, opffiletext, "<item ")
                    End If
                Else
                    'No nav document
                    DialogResult = MsgBox("ERROR: Table of Contents file not found." + Chr(10) + "This ebook is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Button20.Enabled = False
                    Button23.Enabled = False
                    tocfile = ""
                    Button20.Text = "Edit nav file"
                    GoTo lookforpagemap
                End If
            End While
        Else
            'Search for toc.ncx file
            searchResults = Directory.GetFiles(ebookdirectory, "*.ncx", SearchOption.AllDirectories)
            If searchResults.Length < 1 Then
                DialogResult = MsgBox("ERROR: Table of Contents file not found." + Chr(10) + "This ebook is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Button20.Enabled = False
                Button23.Enabled = False
                tocfile = ""
            Else
                tocfile = searchResults(0)
                If InStr(tocfile, "_MACOSX") Then
                    If searchResults.Length > 1 Then
                        tocfile = searchResults(1)
                    End If
                End If
                Button20.Enabled = True
                Button23.Enabled = True
            End If
            Button20.Text = "Edit toc.ncx file"
        End If

lookforpagemap:
        'Search for page-map.xml file
        searchResults = Directory.GetFiles(ebookdirectory, "page-map.xml", SearchOption.AllDirectories)
        If searchResults.Length < 1 Then
            Button24.Enabled = False
            pagemapfile = ""
        Else
            pagemapfile = searchResults(0)
            If InStr(pagemapfile, "_MACOSX") Then
                If searchResults.Length > 1 Then
                    pagemapfile = searchResults(1)
                End If
            End If
            Button24.Enabled = True
        End If

        Button26.Enabled = True
        Button33.Enabled = True
        Button43.Enabled = True

        'Extract metadata into textboxes
        metadatafile = LoadUnicodeFile(opffile)
        ExtractMetadata(metadatafile, True)

        'Process current folder to locate other EPUB files
        searchResults = Directory.GetFiles(IO.Path.GetDirectoryName(OpenFileDialog1.FileName), "*.epub", SearchOption.TopDirectoryOnly)
        Array.Sort(searchResults)
        refreshfilelist = True
        ComboBox3.Items.Clear()
        Dim fi As String
        For Each fi In searchResults
            ComboBox3.Items.Add(fi.Substring(fi.LastIndexOf("\") + 1, fi.Length - fi.LastIndexOf("\") - 1))
        Next
        Dim x As Integer = 0
        While (searchResults(x) <> OpenFileDialog1.FileName)
            x = x + 1
        End While
        currentfilenumber = x + 1 'searchResults is zero based
        ComboBox3.SelectedIndex = x
        ComboBox3.Enabled = True

        'Update interface
        If Not possibleDRM Then
            CaptionString = IO.Path.GetFileName(OpenFileDialog1.FileName) + " [" + currentfilenumber.ToString + "/" + searchResults.Length.ToString + "] - EPUB Metadata Editor"
            Me.Text = CaptionString
            projectchanged = False
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            TextBox13.Enabled = True
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            PictureBox1.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            Button7.Enabled = True
            Button13.Enabled = True
            Button14.Enabled = True
            Button15.Enabled = True
            Button18.Enabled = True
            Button30.Enabled = True
            Button31.Enabled = True
            Button21.Enabled = True
            Button22.Enabled = True
            Button25.Enabled = True
            Button28.Enabled = True
            Button29.Enabled = True
            Button38.Enabled = True
            LinkLabel3.Enabled = True
            'SaveImageAsToolStripMenuItem.Enabled = True
            If versioninfo = "3.0" Then
                Label25.Visible = True
                'Title cannot have 'file-as' apparently
                'TextBox16.Enabled = False
                'Button21.Enabled = False
                'Button22.Enabled = False
                'Button18.Enabled = False
                Button28.Enabled = False
                'DialogResult = MsgBox("Warning: You are opening an EPUB3 file." + Chr(10) + "EPUB3 handing is in alpha-release only.", MsgBoxStyle.Exclamation, "EPUB Metadata Editor")
            End If
        Else
            Button34.Visible = False
            LinkLabel5.Visible = False
            Button20.Enabled = False
            Button23.Enabled = False
            Button24.Enabled = False
            Button26.Enabled = False
            Button33.Enabled = False
            Button33.Enabled = False
            ClearInterface()
            CaptionString = IO.Path.GetFileName(OpenFileDialog1.FileName) + " [" + currentfilenumber.ToString + "/" + searchResults.Length.ToString + "] - EPUB Metadata Editor"
            Me.Text = CaptionString
            projectchanged = False
            DialogResult = MsgBox("ERROR: Image file corrupted or encrypted." + Chr(10) + "Note that EPUB Metadata Editor cannot handle" + Chr(10) + "EPUB files locked by DRM.", MsgBoxStyle.Critical, "EPUB Metadata Editor")
        End If
    End Sub
    Private Sub ExtractMetadata(ByVal metadatafile As String, ByVal extractcover As Boolean)
        Dim startpos, namespacelen, endpos, endheader, lenheader, fileaspos, temploop, rolepos, coverfilepos, nextcharpos, firsttaglength As Integer
        Dim dcnamespace, rolestring, coverfiletext, langtext, hreftype, nextchar, tempstring As String
        Dim idpos, endheaderpos, startheaderpos, temppos, refinespos, oldstartpos, searchpos As Integer
        Dim idinfo, coverid, coverfileid As String

        'Check for non-standard dc namespace tags
        startpos = InStr(metadatafile, "=" + Chr(34) + "http://purl.org/dc/elements/1.1/")
        If startpos <> 0 Then
            ' work backwards to find the xmlns definition
            namespacelen = 0
            While (startpos - namespacelen <> 0)
                namespacelen = namespacelen + 1
                If Mid(metadatafile, startpos - namespacelen, 6) = "xmlns:" Then
                    Exit While
                End If
            End While
            If namespacelen < startpos Then
                dcnamespace = Mid(metadatafile, startpos - namespacelen + 6, namespacelen - 6)
                metadatafile = metadatafile.Replace(dcnamespace + ":", "dc:")
            End If
        End If

        'Check for non-standard opf namespace tags
        If (InStr(metadatafile, "<opf:metadata") Or InStr(metadatafile, "<opf:manifest")) Then
            metadatafile = metadatafile.Replace("<opf:", "<")
            metadatafile = metadatafile.Replace("</opf:", "</")
        End If

        'Tidy opf file first (just in case fields have carriage returns in them)
        metadatafile = Regularise(metadatafile)

        'Check for uppercase UUID scheme
        metadatafile = metadatafile.Replace("scheme=" + Chr(34) + "UUID", "scheme=" + Chr(34) + "uuid")

        'Get title
        Try
            startpos = InStr(metadatafile, "<dc:title")
            If startpos <> 0 Then
                endpos = InStr(metadatafile, "</dc:title>")
                lenheader = Len("<dc:title")
                If Mid(metadatafile, startpos + lenheader, 1) = ">" Then
                    TextBox1.Text = XMLInput(Mid(metadatafile, startpos + lenheader + 1, endpos - startpos - lenheader - 1))
                    ' Look for Calibre's title_sort meta tag
                    startpos = InStr(metadatafile, "calibre:title_sort")
                    If startpos <> 0 Then
                        temppos = startpos
                        startpos = InStrRev(metadatafile, "<meta ", startpos)
                        If startpos = 0 Then startpos = InStrRev(metadatafile, "<opf:meta ")
                        startpos = InStr(startpos, metadatafile, "content=")
                        endpos = InStr(startpos + 9, metadatafile, Chr(34))
                        If ((startpos <> 0) And (startpos < endpos)) Then
                            TextBox16.Text = XMLInput(Mid(metadatafile, startpos + 9, endpos - startpos - 9))
                        End If
                    End If
                Else
                    'Get optional attributes
                    fileaspos = InStr(startpos, metadatafile, "opf:file-as=")
                    If fileaspos <> 0 Then
                        For temploop = fileaspos + 13 To endpos
                            If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                TextBox16.Text = XMLInput(Mid(metadatafile, fileaspos + 13, temploop - fileaspos - 13))
                                Exit For
                            End If
                        Next
                    End If
                    For temploop = startpos To endpos
                        If Mid(metadatafile, temploop, 1) = ">" Then
                            TextBox1.Text = XMLInput(Mid(metadatafile, temploop + 1, endpos - temploop - 1))
                            Exit For
                        End If
                    Next
                End If
                If versioninfo = "3.0" Then
                    ' Look for Calibre's title_sort meta tag
                    startpos = InStr(metadatafile, "calibre:title_sort")
                    If startpos <> 0 Then
                        temppos = startpos
                        startpos = InStrRev(metadatafile, "<meta ", startpos)
                        If startpos = 0 Then startpos = InStrRev(metadatafile, "<opf:meta ")
                        endpos = InStr(startpos, metadatafile, "/>")
                        startpos = InStr(startpos, metadatafile, "content=")
                        If ((startpos <> 0) And (startpos < endpos)) Then
                            TextBox16.Text = XMLInput(Mid(metadatafile, startpos + 9, endpos - startpos - 10))
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox1.Text = "ERROR"
        End Try

        'Get creator
        Try
            startpos = InStr(metadatafile, "<dc:creator")
            If startpos <> 0 Then
                endpos = InStr(metadatafile, "</dc:creator>")
                lenheader = Len("<dc:creator")
                If Mid(metadatafile, startpos + lenheader + 1) = ">" Then
                    TextBox2.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                Else
                    If versioninfo = "3.0" Then
                        endheaderpos = InStr(startpos, metadatafile, ">")
                        TextBox2.Text = XMLInput(Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1))
                        'Get id
                        idpos = InStr(startpos, metadatafile, "id=")
                        idinfo = ""
                        If idpos <> 0 Then
                            For temploop = idpos + 4 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                    Exit For
                                End If
                            Next
                        End If

                        If idinfo = "" Then idinfo = "creator"

                        temppos = InStr(startpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo + Chr(34))
                        While temppos <> 0
                            endheaderpos = InStr(temppos, metadatafile, ">")
                            startheaderpos = InStrRev(metadatafile, "<", temppos)
                            endpos = InStr(temppos, metadatafile, "</meta>")
                            If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                            refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "file-as")
                            If refinespos <> 0 Then
                                If refinespos < endpos Then
                                    TextBox12.Text = XMLInput(Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1))
                                End If
                            End If
                            refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "role")
                            If refinespos <> 0 Then
                                If refinespos < endpos Then
                                    rolestring = Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1)
                                    If rolestring = "aut" Then
                                        ComboBox1.SelectedIndex = 0
                                    ElseIf rolestring = "edt" Then
                                        ComboBox1.SelectedIndex = 1
                                    ElseIf rolestring = "ill" Then
                                        ComboBox1.SelectedIndex = 2
                                    ElseIf rolestring = "trl" Then
                                        ComboBox1.SelectedIndex = 3
                                    Else
                                        ComboBox1.SelectedIndex = 0
                                    End If
                                End If
                            End If
                            temppos = InStr(endpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo + Chr(34))
                        End While
                    Else
                        'Get optional attributes
                        fileaspos = InStr(startpos, metadatafile, "opf:file-as=")
                        If ((fileaspos <> 0) And (fileaspos < endpos)) Then
                            For temploop = fileaspos + 13 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    TextBox12.Text = XMLInput(Mid(metadatafile, fileaspos + 13, temploop - fileaspos - 13))
                                    Exit For
                                End If
                            Next
                        End If

                        rolepos = InStr(startpos, metadatafile, "opf:role=")
                        If ((rolepos <> 0) And (rolepos < endpos)) Then
                            rolestring = Mid(metadatafile, rolepos + 10, 3)
                            If rolestring = "aut" Then
                                ComboBox1.SelectedIndex = 0
                            ElseIf rolestring = "edt" Then
                                ComboBox1.SelectedIndex = 1
                            ElseIf rolestring = "ill" Then
                                ComboBox1.SelectedIndex = 2
                            ElseIf rolestring = "trl" Then
                                ComboBox1.SelectedIndex = 3
                            Else
                                ComboBox1.SelectedIndex = 0
                            End If
                        Else
                            ComboBox1.SelectedIndex = 0
                        End If

                        For temploop = startpos To endpos
                            If Mid(metadatafile, temploop, 1) = ">" Then
                                TextBox2.Text = XMLInput(Mid(metadatafile, temploop + 1, endpos - temploop - 1))
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox2.Text = "ERROR"
        End Try

        If endpos = 0 Then GoTo skipsecondcreator

        'Look for second creator
        Try
            startpos = InStr(metadatafile, "<dc:creator")
            startpos = InStr(startpos + 1, metadatafile, "<dc:creator")
            If startpos <> 0 Then
                endpos = InStr(startpos, metadatafile, "</dc:creator>")
                lenheader = Len("<dc:creator")
                If Mid(metadatafile, startpos + 1) = ">" Then
                    TextBox3.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                Else
                    If versioninfo = "3.0" Then
                        endheaderpos = InStr(startpos, metadatafile, ">")
                        TextBox3.Text = XMLInput(Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1))
                        ' get id
                        idpos = InStr(startpos, metadatafile, "id=")
                        idinfo = ""
                        If idpos <> 0 Then
                            For temploop = idpos + 4 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                    Exit For
                                End If
                            Next
                        End If

                        If idinfo <> "" Then
                            temppos = InStr(startpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo + Chr(34))
                            While temppos <> 0
                                endheaderpos = InStr(temppos, metadatafile, ">")
                                startheaderpos = InStrRev(metadatafile, "<", temppos)
                                endpos = InStr(temppos, metadatafile, "</meta>")
                                If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                                refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "file-as")
                                If refinespos <> 0 Then
                                    If refinespos < endpos Then
                                        TextBox13.Text = XMLInput(Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1))
                                    End If
                                End If
                                refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "role")
                                If refinespos <> 0 Then
                                    If refinespos < endpos Then
                                        rolestring = Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1)
                                        If rolestring = "aut" Then
                                            ComboBox2.SelectedIndex = 0
                                        ElseIf rolestring = "edt" Then
                                            ComboBox2.SelectedIndex = 1
                                        ElseIf rolestring = "ill" Then
                                            ComboBox2.SelectedIndex = 2
                                        ElseIf rolestring = "trl" Then
                                            ComboBox2.SelectedIndex = 3
                                        Else
                                            ComboBox2.SelectedIndex = 0
                                        End If
                                    End If
                                End If
                                temppos = InStr(endpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo + Chr(34))
                            End While
                        End If
                    Else
                        'Get optional attributes
                        fileaspos = InStr(startpos, metadatafile, "opf:file-as=")
                        If ((fileaspos <> 0) And (fileaspos < endpos)) Then
                            For temploop = fileaspos + 13 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    TextBox13.Text = XMLInput(Mid(metadatafile, fileaspos + 13, temploop - fileaspos - 13))
                                    Exit For
                                End If
                            Next
                        End If

                        rolepos = InStr(startpos, metadatafile, "opf:role=")
                        If ((rolepos <> 0) And (rolepos < endpos)) Then
                            rolestring = Mid(metadatafile, rolepos + 10, 3)
                            If rolestring = "aut" Then ComboBox2.SelectedIndex = 0
                            If rolestring = "edt" Then ComboBox2.SelectedIndex = 1
                            If rolestring = "ill" Then ComboBox2.SelectedIndex = 2
                            If rolestring = "trl" Then ComboBox2.SelectedIndex = 3
                        End If

                        For temploop = startpos To endpos
                            If Mid(metadatafile, temploop, 1) = ">" Then
                                TextBox3.Text = XMLInput(Mid(metadatafile, temploop + 1, endpos - temploop - 1))
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox3.Text = "ERROR"
        End Try

skipsecondcreator:

        'Get (Calibre) Series and Series Index
        Try
            startpos = InStr(metadatafile, "calibre:series")
            If startpos <> 0 Then
                ' The following line is necessary because some EPUBS have calibre:series_index but no calibre:series or calibre:series_index comes before calibre:series
                startpos = InStr(metadatafile, "calibre:series" + Chr(34))
                If startpos <> 0 Then
                    ' find start of entry
                    temppos = startpos
                    startpos = InStrRev(metadatafile, "<meta", startpos)
                    If startpos = 0 Then startpos = InStrRev(metadatafile, "<opf:meta")
                    ' find content
                    startpos = InStr(startpos, metadatafile, "content=" & Chr(34))
                    If startpos <> 0 Then
                        lenheader = Len("content=" & Chr(34))
                        endpos = InStr(startpos + lenheader, metadatafile, Chr(34))
                        TextBox15.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                    End If
                End If

                startpos = InStr(metadatafile, "calibre:series_index" + Chr(34))
                If startpos <> 0 Then
                    ' find start of entry
                    temppos = startpos
                    startpos = InStrRev(metadatafile, "<meta", startpos)
                    If startpos = 0 Then startpos = InStrRev(metadatafile, "<opf:meta")
                    ' find content
                    startpos = InStr(startpos, metadatafile, "content=" & Chr(34))
                    If startpos <> 0 Then
                        lenheader = Len("content=" & Chr(34))
                        endpos = InStr(startpos + lenheader, metadatafile, Chr(34))
                        TextBox14.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox15.Text = "ERROR"
        End Try

        'Get Description
        Try
            metadatafile = metadatafile.Replace("<dc:description xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:description />")
            If (InStr(metadatafile, "<dc:description />") = 0) Then
                startpos = InStr(metadatafile, "<dc:description/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:description")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<description")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:description>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</description>")
                        TextBox4.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                        TextBox4.Text = TextBox4.Text.Replace("<![CDATA[", "")
                        TextBox4.Text = TextBox4.Text.Replace("]]>", "")
                        Application.DoEvents()
                        WebBrowser1.DocumentText = "<head><style type=" + Chr(34) + "text/css" + Chr(34) + ">" + "body {margin:0;padding:0}" + "</style></head>" + TextBox4.Text.Replace(Chr(10), "<br>")
                        WebBrowser1.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox4.Text = "ERROR"
            WebBrowser1.Visible = False
        End Try

        'Get Publisher
        Try
            metadatafile = metadatafile.Replace("<dc:publisher xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:publisher />")
            If (InStr(metadatafile, "<dc:publisher />") = 0) Then
                startpos = InStr(metadatafile, "<dc:publisher/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:publisher")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<publisher")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:publisher>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</publisher>")
                        TextBox5.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox5.Text = "ERROR"
        End Try

        'Get Date
        Try
            Label6.Text = "Date"
            TextBox6.Width = 304
            TextBox6.Left = 81
            startpos = InStr(metadatafile, "<dc:date")
            firsttaglength = 8
            If startpos = 0 Then
                startpos = InStr(metadatafile, "<date")
                firsttaglength = 5
            End If
            If startpos <> 0 Then
                endheader = InStr(startpos, metadatafile, ">")
                lenheader = endheader - startpos + 1
                endpos = InStr(metadatafile, "</dc:date>")
                If endpos = 0 Then endpos = InStr(metadatafile, "</date>")
                If (Mid(metadatafile, startpos + firsttaglength, 1) = ">") Then
                    TextBox6.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                Else
                    'Get optional attribute: event
                    fileaspos = InStr(startpos, metadatafile, "opf:event=")
                    If (fileaspos <> 0) Then
                        If (fileaspos < endpos) Then
                            TextBox6.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                            Label6.Text = "Date (" + Mid(metadatafile, fileaspos + 11, endheader - fileaspos - 12) + ")"
                            TextBox6.Width = 255
                            TextBox6.Left = 130
                        Else
                            TextBox6.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                        End If
                    Else
                        TextBox6.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox6.Text = "ERROR"
        End Try

        'Get Subject
        Try
            metadatafile = metadatafile.Replace("<dc:subject xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:subject />")
            If (InStr(metadatafile, "<dc:subject />") = 0) Then
                startpos = InStr(metadatafile, "<dc:subject/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:subject")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<subject")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:subject>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</subject>")
                        TextBox17.Text = XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                        oldstartpos = startpos
                        startpos = InStr(oldstartpos + 1, metadatafile, "<dc:subject")
                        If startpos = 0 Then startpos = InStr(oldstartpos + 1, metadatafile, "<subject")
                        While startpos <> 0
                            endheader = InStr(startpos, metadatafile, ">")
                            lenheader = endheader - startpos + 1
                            endpos = InStr(startpos, metadatafile, "</dc:subject>")
                            If endpos = 0 Then endpos = InStr(startpos, metadatafile, "</subject>")
                            TextBox17.Text = TextBox17.Text + subjectseparator + XMLInput(Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader))
                            oldstartpos = startpos
                            startpos = InStr(oldstartpos + 1, metadatafile, "<dc:subject")
                            If startpos = 0 Then startpos = InStr(oldstartpos + 1, metadatafile, "<subject")
                        End While
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox17.Text = "ERROR"
        End Try

        'Get Type
        Try
            metadatafile = metadatafile.Replace("<dc:type xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:type />")
            If (InStr(metadatafile, "<dc:type />") = 0) Then
                startpos = InStr(metadatafile, "<dc:type/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:type")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<type")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:type>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</type>")
                        TextBox7.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox7.Text = "ERROR"
        End Try

        'Get Format
        Try
            metadatafile = metadatafile.Replace("<dc:format xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:format />")
            If (InStr(metadatafile, "<dc:format />") = 0) Then
                startpos = InStr(metadatafile, "<dc:format/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:format/>")
                    If startpos = 0 Then
                        startpos = InStr(metadatafile, "<dc:format")
                        If startpos = 0 Then startpos = InStr(metadatafile, "<format")
                        If startpos <> 0 Then
                            endheader = InStr(startpos, metadatafile, ">")
                            lenheader = endheader - startpos + 1
                            endpos = InStr(metadatafile, "</dc:format>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</format>")
                            TextBox8.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox8.Text = "ERROR"
        End Try

        'Get Identifier
        Try
            'Look for multiple identifiers
            idcount = countString(metadatafile, "<dc:identifier")
            If idcount > 1 Then
                Label31.Text = "Multiple" + Chr(10) + "Identifiers"
                Label31.Visible = True
            Else
                Label31.Visible = False
            End If

            'before looking for first identifier, look for scheme="uuid"
            startpos = InStr(metadatafile, "opf:scheme=" + Chr(34) + "uuid")
            If startpos <> 0 Then
                'Scan backwards to <dc:identifier
                startpos = InStrRev(metadatafile, "<dc:identifier", startpos)
            Else
                'Find the first <dc:identifier
                startpos = InStr(metadatafile, "<dc:identifier")
            End If
            firsttaglength = 14
            If startpos = 0 Then
                startpos = InStr(metadatafile, "<identifier")
                firsttaglength = 11
            End If
            If startpos <> 0 Then
                Dim nocontent As Boolean
                nocontent = False
                endheader = InStr(startpos, metadatafile, ">")
                lenheader = endheader - startpos + 1
                endpos = InStr(startpos, metadatafile, "</dc:identifier>")
                If endpos = 0 Then endpos = InStr(startpos, metadatafile, "</identifier>")
                If endpos = 0 Then
                    endpos = InStr(startpos, metadatafile, " />")
                    If endpos <> 0 Then nocontent = True
                End If
                If endpos <> 0 Then
                    If (Mid(metadatafile, startpos + firsttaglength, 1) = ">") Then
                        TextBox9.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                        Label9.Text = "Identifier"
                        TextBox9.Width = 304
                        TextBox9.Left = 81
                    Else
                        If versioninfo = "3.0" Then
                            endheaderpos = InStr(startpos, metadatafile, ">")
                            TextBox9.Text = Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1)
                            Label9.Text = "Identifier"
                            TextBox9.Width = 304
                            TextBox9.Left = 81
                            'Get id
                            idpos = InStr(startpos, metadatafile, "id=")
                            idinfo = ""
                            If idpos <> 0 Then
                                For temploop = idpos + 4 To endpos
                                    If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                        idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                        Exit For
                                    End If
                                Next
                            End If
                            If idinfo <> "" Then
                                temppos = InStr(startpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo)
                                While temppos <> 0
                                    endheaderpos = InStr(temppos, metadatafile, ">")
                                    startheaderpos = InStrRev(metadatafile, "<", temppos)
                                    endpos = InStr(temppos, metadatafile, "</meta>")
                                    If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                                    refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "identifier-type")
                                    If refinespos <> 0 Then
                                        If refinespos < endpos Then
                                            refinespos = InStr(refinespos, metadatafile, "scheme=" + Chr(34))
                                            If refinespos <> 0 Then
                                                If refinespos < endpos Then
                                                    Label9.Text = "Identifier (" + Mid(metadatafile, refinespos + 8, endheaderpos - refinespos - 10) + "=" + Mid(metadatafile, endheaderpos + 1, endpos - endheaderpos - 1) + ")"
                                                    TextBox9.Width = 255
                                                    TextBox9.Left = 130
                                                End If
                                            End If
                                        End If
                                    End If
                                    temppos = InStr(endpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo)
                                End While
                            End If
                        Else
                            'Get optional attribute: scheme
                            fileaspos = InStr(startpos, metadatafile, "opf:scheme=")
                            If fileaspos <> 0 Then
                                For temploop = fileaspos + 13 To endpos
                                    If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                        Label9.Text = "Identifier (" + Mid(metadatafile, fileaspos + 12, temploop - fileaspos - 12) + ")"
                                        Exit For
                                    End If
                                Next
                                If nocontent = False Then
                                    TextBox9.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                                Else
                                    'Get id
                                    idpos = InStr(startpos, metadatafile, "id=")
                                    TextBox9.Text = ""
                                    If idpos <> 0 Then
                                        For temploop = idpos + 4 To endpos
                                            If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                                TextBox9.Text = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                                TextBox9.Width = 255
                                TextBox9.Left = 130
                            Else
                                If nocontent = False Then
                                    TextBox9.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                                Else
                                    'Get id
                                    idpos = InStr(startpos, metadatafile, "id=")
                                    TextBox9.Text = ""
                                    If idpos <> 0 Then
                                        For temploop = idpos + 4 To endpos
                                            If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                                TextBox9.Text = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                                Label9.Text = "Identifier"
                                TextBox9.Width = 304
                                TextBox9.Left = 81
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox9.Text = "ERROR"
            TextBox9.Width = 304
            TextBox9.Left = 81
        End Try

        'Get source
        Try
            metadatafile = metadatafile.Replace("<dc:source xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:source />")
            If (InStr(metadatafile, "<dc:source />") = 0) Then
                startpos = InStr(metadatafile, "<dc:source/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:source")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<source")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:source>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</source>")
                        TextBox10.Text = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox10.Text = "ERROR"
        End Try

        'Get Language
        Try
            metadatafile = metadatafile.Replace("<dc:language xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:language />")
            If (InStr(metadatafile, "<dc:language />") = 0) Then
                startpos = InStr(metadatafile, "<dc:language/>")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<dc:language")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<language")
                    If startpos <> 0 Then
                        endheader = InStr(startpos, metadatafile, ">")
                        lenheader = endheader - startpos + 1
                        endpos = InStr(metadatafile, "</dc:language>")
                        If endpos = 0 Then endpos = InStr(metadatafile, "</language>")
                        langtext = Mid(metadatafile, startpos + lenheader, endpos - startpos - lenheader)
                        TextBox11.Text = langtext
                    End If
                End If
            End If
        Catch ex As Exception
            TextBox11.Text = "ERROR"
        End Try

        If extractcover = False Then GoTo didnotfindhref

        'Get Cover
        Try
            possibleDRM = False
            Label24.Visible = False

            'Start by searching for <meta ... name="cover"
            startpos = InStr(metadatafile, "<metadata")
            If startpos = 0 Then
                startpos = InStr(metadatafile, "<opf:metadata")
            End If
            endpos = InStr(metadatafile, "</metadata")
            If endpos = 0 Then
                endpos = InStr(metadatafile, "</opf:metadata")
            End If
            If (startpos <> 0) Then
                ' look for "name="cover""
                hreftype = "name=" + Chr(34) + "cover" + Chr(34)
                searchpos = InStr(startpos, metadatafile, hreftype)
                If ((searchpos <> 0) And (searchpos < endpos)) Then
                    ' Found "name="cover"" in the metadata section
                    ' Search backwards for "<meta " (just in case)
                    nextcharpos = searchpos - 1
                    nextchar = Mid(metadatafile, nextcharpos, 1)
                    tempstring = ""
                    While nextchar <> "<"
                        nextcharpos = nextcharpos - 1
                        nextchar = Mid(metadatafile, nextcharpos, 1)
                    End While
                    tempstring = Mid(metadatafile, nextcharpos, 6)
                    If tempstring <> "<meta " Then GoTo didnotfindhref
                    ' search forwards for content="id of cover"
                    nextcharpos = searchpos + 1
                    nextchar = Mid(metadatafile, nextcharpos, 1)
                    tempstring = ""
                    While nextchar <> ">"
                        tempstring = Mid(metadatafile, nextcharpos, 8)
                        If tempstring = "content=" Then Exit While
                        nextcharpos = nextcharpos + 1
                        nextchar = Mid(metadatafile, nextcharpos, 1)
                    End While
                    If tempstring <> "content=" Then
                        ' search backwards for content="id of cover"
                        nextcharpos = searchpos - 1
                        nextchar = Mid(metadatafile, nextcharpos, 1)
                        tempstring = ""
                        While nextchar <> "<"
                            tempstring = Mid(metadatafile, nextcharpos, 8)
                            If tempstring = "content=" Then Exit While
                            nextcharpos = nextcharpos - 1
                            nextchar = Mid(metadatafile, nextcharpos, 1)
                        End While
                    End If
                    If tempstring = "content=" Then
                        coverfileid = nextcharpos
                        endpos = InStr(coverfileid + 9, metadatafile, Chr(34))
                        coverid = Mid(metadatafile, coverfileid + 9, endpos - coverfileid - 9)
                        ' Now search for coverid in <manifest
                        startpos = InStr(metadatafile, "<manifest")
                        If startpos = 0 Then
                            startpos = InStr(metadatafile, "<opf:manifest")
                        End If
                        If startpos <> 0 Then
                            hreftype = "id=" + Chr(34) + coverid + Chr(34)
                            coverfilepos = InStr(startpos, metadatafile, hreftype)
                            If coverfilepos > 0 Then
                                GoTo foundcoverid
                            End If
                        Else
                            coverfilepos = 0
                        End If
                    Else
                        coverfilepos = 0
                    End If
                End If
            End If

            'Try alternatives
            startpos = InStr(metadatafile, "<guide")
            If startpos = 0 Then
                startpos = InStr(metadatafile, "<opf:guide")
            End If
            If startpos = 0 Then
                '<guide> is now deprecated
                startpos = InStr(metadatafile, "<manifest")
                If startpos = 0 Then
                    startpos = InStr(metadatafile, "<opf:manifest")
                End If
                If startpos <> 0 Then
                    hreftype = "id=" + Chr(34) + "cov" + Chr(34)
                    coverfilepos = InStr(startpos, metadatafile, hreftype)
                    If coverfilepos = 0 Then
                        hreftype = "id=" + Chr(34) + "cover" + Chr(34)
                        coverfilepos = InStr(startpos, metadatafile, hreftype)
                        If coverfilepos = 0 Then
                            hreftype = "id=" + Chr(34) + "coverpage" + Chr(34)
                            coverfilepos = InStr(startpos, metadatafile, hreftype)
                            If coverfilepos = 0 Then
                                hreftype = "properties=" + Chr(34) + "cover-image" + Chr(34)
                                coverfilepos = InStr(startpos, metadatafile, hreftype)
                            End If
                        End If
                    End If
                End If
            Else
                hreftype = "type=" + Chr(34) + "cover" + Chr(34)
                coverfilepos = InStr(startpos, metadatafile, hreftype)
                If coverfilepos = 0 Then
                    'some Sony Reader Library files require different processing
                    hreftype = "title=" + Chr(34) + "Cover" + Chr(34)
                    coverfilepos = InStr(startpos, metadatafile, hreftype)
                    If coverfilepos = 0 Then
                        hreftype = "title=" + Chr(34) + "Cover Page" + Chr(34)
                        coverfilepos = InStr(startpos, metadatafile, hreftype)
                        If coverfilepos = 0 Then
                            hreftype = "type=" + Chr(34) + "coverimagestandard" + Chr(34)
                            coverfilepos = InStr(startpos, metadatafile, hreftype)
                        End If
                    End If
                End If
            End If

            If coverfilepos = 0 Then
                ' Last ditch effort: search for <meta name="cover"
            End If

            If coverfilepos = 0 Then GoTo didnotfindhref

foundcoverid:
            'find href (scanning forwards) 
            nextcharpos = coverfilepos + 1
            nextchar = Mid(metadatafile, nextcharpos, 1)
            While nextchar <> ">"
                tempstring = Mid(metadatafile, nextcharpos, 5)
                If tempstring = "href=" Then GoTo foundhref
                nextcharpos = nextcharpos + 1
                nextchar = Mid(metadatafile, nextcharpos, 1)
            End While

            'find href (scanning backwards) 
            nextcharpos = coverfilepos - 1
            nextchar = Mid(metadatafile, nextcharpos, 1)
            While nextchar <> "<"
                tempstring = Mid(metadatafile, nextcharpos, 5)
                If tempstring = "href=" Then GoTo foundhref
                nextcharpos = nextcharpos - 1
                nextchar = Mid(metadatafile, nextcharpos, 1)
            End While
            GoTo didnotfindhref
foundhref:
            coverfilepos = nextcharpos
            endpos = InStr(coverfilepos + 6, metadatafile, Chr(34))
            coverfile = Path.GetDirectoryName(opffile) + "\" + Mid(metadatafile, coverfilepos + 6, endpos - coverfilepos - 6).Replace("/", "\")

            If ((Path.GetExtension(coverfile) = ".jpg") Or (Path.GetExtension(coverfile) = ".jpeg") Or (Path.GetExtension(coverfile) = ".png")) Then
                If System.IO.File.Exists(coverfile) Then
                    coverimagefile = coverfile
                    PictureBox1.ImageLocation = coverfile
                    Try
                        PictureBox1.Load()
                        SaveImageAsToolStripMenuItem.Enabled = True
                        ChangeImageToolStripMenuItem.Enabled = True
                        UseExistingImageToolStripMenuItem.Enabled = True
                        AddImageToolStripMenuItem.Enabled = False
                    Catch ex As Exception
                        possibleDRM = True
                        GoTo exitsub
                    End Try
                    GoTo updateinterface
                End If
            Else

parsecoverfile:
                'Parse coverfile for image information
                If System.IO.File.Exists(coverfile) Then
                    coverfiletext = LoadUnicodeFile(coverfile)
                    startpos = InStr(coverfiletext, "<svg")
                    If startpos <> 0 Then
                        Label4.Visible = True
                        Button1.Visible = True
                        Button42.Visible = True
                    End If
                    startpos = InStr(coverfiletext, "<img")
                    If startpos <> 0 Then
                        startpos = InStr(startpos, coverfiletext, "src")
                        startpos = InStr(startpos, coverfiletext, Chr(34))
                        If startpos <> 0 Then
                            endpos = InStr(startpos + 1, coverfiletext, Chr(34))
                            If endpos <> 0 Then
                                relativecoverimagefile = Mid(coverfiletext, startpos + 1, endpos - startpos - 1)
                                coverimagefile = Path.GetDirectoryName(coverfile) + "\" + Mid(coverfiletext, startpos + 1, endpos - startpos - 1).Replace("/", "\")
                                If System.IO.File.Exists(coverimagefile) Then
                                    PictureBox1.ImageLocation = coverimagefile
                                    Try
                                        PictureBox1.Load()
                                        SaveImageAsToolStripMenuItem.Enabled = True
                                        ChangeImageToolStripMenuItem.Enabled = True
                                        UseExistingImageToolStripMenuItem.Enabled = True
                                        AddImageToolStripMenuItem.Enabled = False
                                    Catch ex As Exception
                                        possibleDRM = True
                                        GoTo exitsub
                                    End Try
                                    GoTo updateinterface
                                End If
                            End If
                        End If
                    Else
                        startpos = InStr(coverfiletext, "<image")
                        If startpos <> 0 Then
                            startpos = InStr(startpos, coverfiletext, "href")
                            startpos = InStr(startpos, coverfiletext, Chr(34))
                            If startpos <> 0 Then
                                endpos = InStr(startpos + 1, coverfiletext, Chr(34))
                                If endpos <> 0 Then
                                    relativecoverimagefile = Mid(coverfiletext, startpos + 1, endpos - startpos - 1)
                                    coverimagefile = Path.GetDirectoryName(coverfile) + "\" + Mid(coverfiletext, startpos + 1, endpos - startpos - 1).Replace("/", "\")
                                    If System.IO.File.Exists(coverimagefile) Then
                                        PictureBox1.ImageLocation = coverimagefile
                                        Try
                                            PictureBox1.Load()
                                            ChangeImageToolStripMenuItem.Enabled = True
                                            UseExistingImageToolStripMenuItem.Enabled = True
                                            AddImageToolStripMenuItem.Enabled = False
                                        Catch ex As Exception
                                            possibleDRM = True
                                            GoTo exitsub
                                        End Try
                                        GoTo updateinterface
                                    End If
                                End If
                            End If
                        Else
                            startpos = InStr(coverfiletext, "<svg:image")
                            If startpos <> 0 Then
                                startpos = InStr(startpos, coverfiletext, "href")
                                startpos = InStr(startpos, coverfiletext, Chr(34))
                                If startpos <> 0 Then
                                    endpos = InStr(startpos + 1, coverfiletext, Chr(34))
                                    If endpos <> 0 Then
                                        relativecoverimagefile = Mid(coverfiletext, startpos + 1, endpos - startpos - 1)
                                        coverimagefile = Path.GetDirectoryName(coverfile) + "\" + Mid(coverfiletext, startpos + 1, endpos - startpos - 1).Replace("/", "\")
                                        If System.IO.File.Exists(coverimagefile) Then
                                            PictureBox1.ImageLocation = coverimagefile
                                            Try
                                                PictureBox1.Load()
                                                ChangeImageToolStripMenuItem.Enabled = True
                                                UseExistingImageToolStripMenuItem.Enabled = True
                                                AddImageToolStripMenuItem.Enabled = False
                                            Catch ex As Exception
                                                possibleDRM = True
                                                GoTo exitsub
                                            End Try
                                            GoTo updateinterface
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Label24.Visible = True
        End Try


didnotfindhref:
        ChangeImageToolStripMenuItem.Enabled = False
        UseExistingImageToolStripMenuItem.Enabled = True
        AddImageToolStripMenuItem.Enabled = True

        If extractcover Then
            'Look for existing images
            startpos = InStr(metadatafile, "<manifest")
            endpos = InStr(metadatafile, "</manifest>")
            Dim imgpos, hrefpos, endhrefpos, imgnum As Integer
            Dim href, imgfilename As String
            imgpos = startpos
            imgnum = 0
            While (imgpos < endpos)
                imgpos = InStr(imgpos + 1, metadatafile, "media-type=" + Chr(34) + "image/jpeg")
                If ((imgpos = 0) Or (imgpos > endpos)) Then
                    Exit While
                End If

                'Scan backwards looking for start of <item>
                temppos = imgpos
                While (temppos > startpos)
                    temppos = temppos - 1
                    If (Mid(metadatafile, temppos, 5) = "<item") Then
                        Exit While
                    End If
                End While
                hrefpos = InStr(temppos, metadatafile, "href=")
                endhrefpos = InStr(hrefpos + 6, metadatafile, Chr(34))
                href = Mid(metadatafile, hrefpos + 6, endhrefpos - hrefpos - 6)
                href = href.Replace("%20", " ")
                ListBox2.Items.Add(href)
                imgnum = imgnum + 1
                If ((InStr(href.ToLower, "cover") <> 0) And (InStr(href.ToLower, "backcover") = 0)) Then
                    ListBox2.SelectedIndex = imgnum - 1
                End If
            End While

            imgpos = startpos
            While (imgpos < endpos)
                imgpos = InStr(imgpos + 1, metadatafile, "media-type=" + Chr(34) + "image/png")
                If ((imgpos = 0) Or (imgpos > endpos)) Then
                    Exit While
                End If

                'Scan backwards looking for start of <item>
                temppos = imgpos
                While (temppos > startpos)
                    temppos = temppos - 1
                    If (Mid(metadatafile, temppos, 5) = "<item") Then
                        Exit While
                    End If
                End While
                hrefpos = InStr(temppos, metadatafile, "href=")
                endhrefpos = InStr(hrefpos + 6, metadatafile, Chr(34))
                href = Mid(metadatafile, hrefpos + 6, endhrefpos - hrefpos - 6)
                href = href.Replace("%20", " ")
                ListBox2.Items.Add(href)
                imgnum = imgnum + 1
                If (InStr(href, "cover") <> 0) Then
                    ListBox2.SelectedIndex = imgnum - 1
                End If
            End While

            If ListBox2.Items.Count > 0 Then
                If ListBox2.SelectedIndex = -1 Then
                    ListBox2.SelectedIndex = 0
                End If
                'Show preview
                imgfilename = Path.GetDirectoryName(opffile) + "\" + ListBox2.SelectedItem.ToString.Replace("/", "\")
                If System.IO.File.Exists(imgfilename) Then
                    PictureBox2.ImageLocation = imgfilename
                    Try
                        PictureBox2.Load()
                    Catch ex As Exception
                        possibleDRM = True
                        GoTo exitsub
                    End Try
                End If
                GroupBox1.Visible = True
                Label24.Visible = False
            End If
        End If
        GoTo exitsub

updateinterface:
        ' Check to see if cover has been prioritised already
        If ((My.Computer.FileSystem.FileExists(ebookdirectory + "\0000Cover.jpg")) Or (My.Computer.FileSystem.FileExists(ebookdirectory + "\0000Cover.jpeg")) Or (My.Computer.FileSystem.FileExists(ebookdirectory + "\0000Cover.png"))) Then
            Button27.Visible = False
            Label23.Visible = False
            CheckBox5.Visible = False
            If ((Button1.Visible = False) And (Button35.Visible = False)) Then
                Button42.Visible = False
            End If
        Else
            Button27.Visible = True
            Label23.Visible = True
            CheckBox5.Visible = True
            Button42.Visible = True
        End If

        fixcovermetadata = False
        fixcovermanifest = False
        Button35.Visible = False
        Label27.Visible = False
        If ((Button27.Visible = False) And (Button1.Visible = False)) Then
            Button42.Visible = False
        End If

        If relativecoverimagefile <> "" Then
            Dim pos As Integer

            ' Check to see if cover image information is in metadata
            ' e.g. <meta content="cover.jpg" name="cover"/>
            pos = relativecoverimagefile.LastIndexOf("/")
            coverimagefilename = Mid(relativecoverimagefile, pos + 2)
            pos = InStr(metadatafile, "name=" + Chr(34) + "cover" + Chr(34))
            If pos = 0 Then
                Button35.Visible = True
                Label27.Visible = True
                fixcovermetadata = True
                Button42.Visible = True
            End If

            ' Check to see if cover image information is in manifest
            ' e.g. <item href="Images/cover.jpg" id="cover" media-type="image/jpeg"/>
            startpos = InStr(metadatafile, "<manifest")
            pos = InStr(startpos, metadatafile, "id=" + Chr(34) + "cover" + Chr(34))
            If pos = 0 Then
                Button35.Visible = True
                Label27.Visible = True
                fixcovermanifest = True
                Button42.Visible = True
            Else
                endpos = InStr(metadatafile, "</manifest>")
                If ((pos < startpos) Or (pos > endpos)) Then
                    Button35.Visible = True
                    Label27.Visible = True
                    fixcovermanifest = True
                    Button42.Visible = True
                End If
            End If
        End If
exitsub:
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim imgfilename As String
        imgfilename = Path.GetDirectoryName(opffile) + "\" + ListBox2.SelectedItem.ToString.Replace("/", "\")
        If System.IO.File.Exists(imgfilename) Then
            PictureBox2.ImageLocation = imgfilename
            Try
                PictureBox2.Load()
            Catch ex As Exception
                possibleDRM = True
            End Try
        End If
    End Sub
    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        Dim CoverFile, metadatafile As String
        Dim startpos, endpos As Integer
        CoverFile = ListBox2.SelectedItem
        RichTextBox1.Text = LoadUnicodeFile(opffile)
        metadatafile = LoadUnicodeFile(opffile)

        If (InStr(metadatafile, "<guide>") = 0) Then
            endpos = InStr(metadatafile, "</package>")
            If endpos <> 0 Then
                metadatafile = Mid(metadatafile, 1, endpos - 1) + "<guide>" + Chr(13) + Chr(10) + Chr(9) + "<reference href=" + Chr(34) + CoverFile + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + "</guide>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos)
            End If
        Else
            startpos = InStr(metadatafile, "<guide>")
            endpos = InStr(startpos, metadatafile, "type=" + Chr(34) + "cover")
            If endpos = 0 Then
                metadatafile = Mid(metadatafile, 1, startpos + 7) + Chr(9) + "<reference href=" + Chr(34) + CoverFile + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, startpos + 8)
            Else
                While (Mid(metadatafile, endpos, 5) <> "href=")
                    endpos = endpos - 1
                End While
                startpos = endpos
                endpos = InStr(startpos + 7, metadatafile, Chr(34))
                metadatafile = Mid(metadatafile, 1, startpos + 5) + CoverFile + Mid(metadatafile, endpos)
            End If
        End If

        RichTextBox1.Text = metadatafile
        SaveUnicodeFile(opffile, metadatafile)

        SaveImageAsToolStripMenuItem.Enabled = True
        AddImageToolStripMenuItem.Enabled = False
        ChangeImageToolStripMenuItem.Enabled = True

        GroupBox1.Visible = False

        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString

        ' Need to update metadata
        RichTextBox1.Text = LoadUnicodeFile(opffile)
        metadatafile = LoadUnicodeFile(opffile)
        ExtractMetadata(metadatafile, True)

    End Sub
    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        AddImageToolStripMenuItem.PerformClick()
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If Control.ModifierKeys = (Keys.Control Or Keys.Shift) Then
            If e.KeyCode = System.Windows.Forms.Keys.S Then
                Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
                Me.SizeGripStyle = Windows.Forms.SizeGripStyle.Show
            End If
        End If
    End Sub
    Public Function DealWithPreviousFile() As String
        If projectchanged Then
            DialogResult = Dialog1.ShowDialog
            If DialogResult = Windows.Forms.DialogResult.Cancel Then
                Return ("cancel")
            End If
            If DialogResult = Windows.Forms.DialogResult.Yes Then
                SaveEpub(OpenFileDialog1.FileName, False)
            End If
        End If

        ' Delete previous temp directory (if it exists)
        If tempdirectory <> "" Then
            ChDir(tempdirectory)
            If ebookdirectory <> "" Then
                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                        Dim instance = Convert.ToInt16(Mid(ebookdirectory, InStrRev(ebookdirectory, "B" + 1)))
                        instance = instance + 1
                        ebookdirectory = tempdirectory + "EPUB" + Trim(Str(instance))
                    End Try
                End If
            End If
        End If

        keepcombobox = True
        ClearInterface()
        keepcombobox = False
        SaveImageAsToolStripMenuItem.Enabled = False
        AddImageToolStripMenuItem.Enabled = False
        ChangeImageToolStripMenuItem.Enabled = False
        Return ("proceed")
    End Function
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Check for external viewer
        Dim fileCheck As String
        fileCheck = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor\EPubMetadataEditor.ini"
        If (System.IO.File.Exists(fileCheck) = False) Then
            fileCheck = Application.StartupPath() + "\EPubMetadataEditor.ini"
        End If
        Dim objIniFile As New IniFile(fileCheck)
        Dim ViewerPath As String = objIniFile.GetString("Viewer", "Path", "(none)")
        WordsNotToCapitalise = objIniFile.GetString("Editor", "Words", "a,an,the,at,by,for,in,of,on,to,up,and,as,but,or,nor")
        If ViewerPath <> "(none)" Then
            LinkLabel1.Text = "Change external viewer"
        End If

        'Me.Width = 913
        'Me.Height = 670
        Me.ClientSize = New System.Drawing.Size(905, 640)

        tempdirectory = System.IO.Path.GetTempPath
        ChDir(tempdirectory)

        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1

        subjectseparator = "|"
        ToolTip1.SetToolTip(Me.TextBox17, ToolTip1.GetToolTip(Me.TextBox17) + subjectseparator)

        OpenFileDialog1.FileName = Command()
        If OpenFileDialog1.FileName <> "" Then
            'Check for quotation marks at start and end of commandline and delete them if found
            If ((Mid(OpenFileDialog1.FileName, 1, 1) = Chr(34)) And (Mid(OpenFileDialog1.FileName, Len(OpenFileDialog1.FileName), 1) = Chr(34))) Then
                OpenFileDialog1.FileName = Mid(OpenFileDialog1.FileName, 2, Len(OpenFileDialog1.FileName) - 2)
            End If
            OpenEPUB()
            Button3.Enabled = False

            ' Check for external viewer
            If ViewerPath <> "(none)" Then
                Button8.Enabled = True
                LinkLabel1.Text = "Change external viewer"
            Else
                Button8.Enabled = False
            End If
        Else
            Button8.Enabled = False
        End If

        PictureBox1.AllowDrop = True

        ' Start check for update as background task
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub ProtectClipboard()
        If Clipboard.ContainsText Then
            Dim cliptext As String = Clipboard.GetText
            Clipboard.SetDataObject(cliptext, True)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ProtectClipboard()
        Try
            If projectchanged Then
                DialogResult = Dialog1.ShowDialog
                If DialogResult = Windows.Forms.DialogResult.No Then
                    ChDir(tempdirectory)
                    If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                        Try
                            My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        Catch ex As Exception
                        End Try
                    End If
                    End
                ElseIf DialogResult = Windows.Forms.DialogResult.Yes Then
                    SaveEpub(OpenFileDialog1.FileName, False)
                    ChDir(tempdirectory)
                    If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                        Try
                            My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        Catch ex As Exception
                        End Try
                    End If
                    End
                End If
            Else
                ChDir(tempdirectory)
                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                    End Try
                End If
                End
            End If
        Catch ex As Exception
            End
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If projectchanged Then
            DialogResult = Dialog1.ShowDialog
            If DialogResult = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
            If DialogResult = Windows.Forms.DialogResult.Yes Then
                SaveEpub(OpenFileDialog1.FileName, False)
            End If
        End If

        ' Delete previous temp directory (if it exists)
        If tempdirectory <> "" Then
            ChDir(tempdirectory)
            If ebookdirectory <> "" Then
                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        My.Computer.FileSystem.DeleteDirectory(ebookdirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception
                        Dim instance = Convert.ToInt16(Mid(ebookdirectory, InStrRev(ebookdirectory, "B" + 1)))
                        instance = instance + 1
                        ebookdirectory = tempdirectory + "EPUB" + Trim(Str(instance))
                    End Try
                End If
            End If
        End If

        'File chooser
        OpenFileDialog1.Filter = "EPUB Files (*.epub)|*.epub|All files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.FileName = ""
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            ClearInterface()
            SaveImageAsToolStripMenuItem.Enabled = False
            AddImageToolStripMenuItem.Enabled = False
            ChangeImageToolStripMenuItem.Enabled = False
            OpenEPUB()
            Button3.Enabled = False

            ' Check for external viewer
            Dim fileCheck = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor\EPubMetadataEditor.ini"
            If (System.IO.File.Exists(fileCheck) = False) Then
                fileCheck = Application.StartupPath() + "\EPubMetadataEditor.ini"
            End If
            Dim objIniFile As New IniFile(fileCheck)
            Dim ViewerPath As String = _
                objIniFile.GetString("Viewer", "Path", "(none)")
            If ViewerPath <> "(none)" Then
                Button8.Enabled = True
                LinkLabel1.Text = "Change external viewer"
            Else
                Button8.Enabled = False
                LinkLabel1.Text = "Set external viewer"
            End If

            ' Put cursor in Title box
            TextBox1.Focus()

            ' Enable "Find EPUB3" button
            Button44.Enabled = True
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SaveEpub(OpenFileDialog1.FileName, False)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox12.Text = TextBox2.Text
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox2.Text = TextBox12.Text
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim tempname, nextchar, newname As String
        Dim temppos As Integer

        If My.Computer.Keyboard.ShiftKeyDown Then
            tempname = TextBox2.Text
            If InStr(tempname, ",") <> 0 Then
                temppos = Len(tempname)
                nextchar = Mid(tempname, temppos, 1)
                While (nextchar <> ",")
                    If temppos = 1 Then
                        GoTo errortext
                    End If
                    temppos = temppos - 1
                    nextchar = Mid(tempname, temppos, 1)
                End While
                newname = Mid(tempname, temppos + 2) + " " + Mid(tempname, 1, temppos - 1)
                TextBox2.Text = newname
            End If
        Else
            tempname = TextBox2.Text
            If InStr(tempname, " ") <> 0 Then
                temppos = Len(tempname)
                nextchar = Mid(tempname, temppos, 1)
                While (nextchar <> " ")
                    If temppos = 1 Then
                        GoTo errortext
                    End If
                    temppos = temppos - 1
                    nextchar = Mid(tempname, temppos, 1)
                End While
                newname = Mid(tempname, temppos + 1) + ", " + Mid(tempname, 1, temppos - 1)
                TextBox12.Text = newname
            Else
                TextBox12.Text = tempname
            End If
        End If
errortext:
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        OpenFileDialog3.Filter = "EPUB Files (*.epub)|*.epub"
        OpenFileDialog3.FilterIndex = 1
        OpenFileDialog3.FileName = ""
        If OpenFileDialog3.ShowDialog = Windows.Forms.DialogResult.OK Then
            If projectchanged Then
                DialogResult = Dialog1.ShowDialog
                If DialogResult = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
                If DialogResult = Windows.Forms.DialogResult.Yes Then
                    SaveEpub(OpenFileDialog1.FileName, False)
                End If
            End If

            ClearInterface()

            Dim file As String
            For Each file In OpenFileDialog3.FileNames
                ListBox1.Items.Add(file)
                Button10.Enabled = True
                Button32.Enabled = True
                Button41.Enabled = True
                Button46.Enabled = True
            Next

            DisableInterface()
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim filenum, x As Integer
        Dim metadatafile, tempstring As String
        Dim temp As DialogResult

        ' do some checks first
        If ((CheckBox1.Checked = False) And (CheckBox2.Checked = False) And (CheckBox3.Checked = False) And (CheckBox4.Checked = False) And (CheckBox6.Checked = False) And (CheckBox7.Checked = False) And (CheckBox8.Checked = False) And (CheckBox9.Checked = False) And (CheckBox10.Checked = False) And (CheckBox11.Checked = False) And (CheckBox12.Checked = False) And (CheckBox13.Checked = False)) Then
            MsgBox("You need to check one of the batch task boxes!")
            Exit Sub
        End If

        If ((CheckBox9.Checked) And (TextBox18.Text = "")) Then
            MsgBox("You need to enter the series title!")
            TextBox18.Focus()
            Exit Sub
        End If

        If ((CheckBox11.Checked) And ((Form10.CheckBox1.Checked = False) And (Form10.CheckBox2.Checked = False) And (Form10.CheckBox3.Checked = False))) Then
            MsgBox("You need to select at least one cover fix!")
            Form10.ShowDialog()
            Exit Sub
        End If

        If ((CheckBox12.Checked) And (Form7.CheckedListBox1.CheckedItems.Count = 0)) Then
            MsgBox("You need to select a field to replace!")
            Form7.ShowDialog()
            Exit Sub
        End If

        If ((CheckBox12.Checked) And (Form7.TextBox1.Text = "")) Then
            temp = MsgBox("Are you sure you want to replace that field with an empty string?", MsgBoxStyle.YesNo)
            If temp = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        If ((CheckBox13.Checked) And (Form9.CheckedListBox1.CheckedItems.Count = 0)) Then
            MsgBox("You need to select a field first!")
            Form9.ShowDialog()
            Exit Sub
        End If

        ClearInterface()
        tempdirectory = System.IO.Path.GetTempPath
        ebookdirectory = tempdirectory + "EPUB"

        filenum = ListBox1.Items.Count
        ProgressBar1.Maximum = filenum - 1
        ProgressBar1.Visible = True

        For x = 1 To filenum
            ChDir(tempdirectory)
            ProgressBar1.Value = x - 1
            ProgressBar1.Update()
            Application.DoEvents()

            ' open file
            'Unzip epub to temp directory

            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                Try
                    'delete contents of temp directory
                    DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                Catch
                    wait(500)
                    'try again
                    DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                End Try
            Else
                MkDir(ebookdirectory)
            End If
            ChDir(ebookdirectory)

            Try
                Dim zip As ZipStorer
                zip = ZipStorer.Open(ListBox1.Items(x - 1).ToString, FileAccess.Read)
                Dim dir = zip.ReadCentralDir()
                Dim item As ZipStorer.ZipFileEntry
                For Each item In dir
                    zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
                Next
                zip.Close()
            Catch ex1 As Exception
                Console.Error.WriteLine("exception: {0}", ex1.ToString)
                DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Exit Sub
            End Try

            'Search for .opf file
            searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

            'Open .opf file into RichTextBox
            If searchResults.Length < 1 Then
                DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Return
            Else
                opffile = searchResults(0)
                If InStr(opffile, "_MACOSX") Then
                    If searchResults.Length > 1 Then
                        opffile = searchResults(1)
                    End If
                End If
                opfdirectory = Path.GetDirectoryName(opffile)
                RichTextBox1.Text = LoadUnicodeFile(opffile)
            End If

            'Extract metadata into textboxes
            metadatafile = LoadUnicodeFile(opffile)
            If CheckBox11.Checked = True Then
                ' Need to extract cover
                ExtractMetadata(metadatafile, True)
            Else
                ' No need to extract cover
                ExtractMetadata(metadatafile, False)
            End If

            WebBrowser1.Visible = False

            'Do the processing
            If ((CheckBox1.Checked = True) And (CheckBox8.Checked = True)) Then
                'swap 'file as' and 'creator'
                tempstring = TextBox12.Text
                TextBox12.Text = TextBox2.Text
                TextBox2.Text = tempstring

                tempstring = TextBox13.Text
                TextBox13.Text = TextBox3.Text
                TextBox3.Text = tempstring
            Else
                If CheckBox1.Checked = True Then
                    ' copy 'file as' to 'creator'
                    If TextBox12.Text <> "" Then
                        TextBox2.Text = TextBox12.Text
                    End If

                    If TextBox13.Text <> "" Then
                        TextBox3.Text = TextBox13.Text
                    End If
                End If

                If CheckBox8.Checked = True Then
                    ' copy 'creator' to 'file as'
                    If TextBox2.Text <> "" Then
                        TextBox12.Text = TextBox2.Text
                    End If

                    If TextBox3.Text <> "" Then
                        TextBox13.Text = TextBox3.Text
                    End If
                End If
            End If


            If ((CheckBox6.Checked = True) And (CheckBox7.Checked = True)) Then
                'swap 'file as' and 'title'
                tempstring = TextBox16.Text
                TextBox16.Text = TextBox1.Text
                TextBox1.Text = tempstring
            Else
                If CheckBox6.Checked = True Then
                    ' copy 'file as' to 'title'
                    If TextBox16.Text <> "" Then
                        TextBox1.Text = TextBox16.Text
                    End If
                End If

                If CheckBox7.Checked = True Then
                    ' copy 'title' to 'file as'
                    If TextBox1.Text <> "" Then
                        TextBox16.Text = TextBox1.Text
                    End If
                End If
            End If

            If CheckBox2.Checked = True Then
                ' apply Title Case to 'title'
                TextBox1.Text = TitleCase(TextBox1.Text)
            End If

            If CheckBox3.Checked = True Then
                ' remove Title's 'File As'
                TextBox16.Text = ""
            End If

            If CheckBox4.Checked = True Then
                ' Autogenerate Creator's 'File as'
                Button7_Click(sender, e)
                If TextBox3.Text <> "" Then Button13.PerformClick()
            End If

            If CheckBox9.Checked = True Then
                ' Serialise
                TextBox15.Text = TextBox18.Text
                TextBox14.Text = x.ToString
            End If

            If CheckBox10.Checked = True Then
                ' Swap Title and Creator
                tempstring = TextBox1.Text
                TextBox1.Text = TextBox2.Text
                TextBox2.Text = tempstring
            End If

            If CheckBox12.Checked = True Then
                ' Replace contents of selected field with text
                Dim indexChecked As Integer

                For Each indexChecked In Form7.CheckedListBox1.CheckedIndices
                    If indexChecked = 0 Then
                        ' Title
                        TextBox1.Text = ReplaceField(TextBox1.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 1 Then
                        ' Title File as
                        TextBox16.Text = ReplaceField(TextBox16.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 2 Then
                        ' Creator1
                        TextBox2.Text = ReplaceField(TextBox2.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 3 Then
                        ' Creator1 File as
                        TextBox12.Text = ReplaceField(TextBox12.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 4 Then
                        ' Creator2
                        TextBox3.Text = ReplaceField(TextBox3.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 5 Then
                        ' Creator2 File as
                        TextBox13.Text = ReplaceField(TextBox13.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 6 Then
                        ' Series
                        TextBox15.Text = ReplaceField(TextBox15.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 7 Then
                        ' Series index
                        TextBox14.Text = ReplaceField(TextBox14.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 8 Then
                        ' Description
                        TextBox4.Text = ReplaceField(TextBox4.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 9 Then
                        ' Publisher
                        TextBox5.Text = ReplaceField(TextBox5.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 10 Then
                        ' Date
                        TextBox6.Text = ReplaceField(TextBox6.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 11 Then
                        ' Subject
                        TextBox17.Text = ReplaceField(TextBox17.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 12 Then
                        ' Type
                        TextBox7.Text = ReplaceField(TextBox7.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 13 Then
                        ' Format
                        TextBox8.Text = ReplaceField(TextBox8.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 14 Then
                        ' Identifier
                        TextBox9.Text = ReplaceField(TextBox9.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 15 Then
                        ' Source
                        TextBox10.Text = ReplaceField(TextBox10.Text, Form7.TextBox1.Text)
                    End If
                    If indexChecked = 16 Then
                        ' Language
                        TextBox11.Text = ReplaceField(TextBox11.Text, Form7.TextBox1.Text)
                    End If
                Next
            End If

            If CheckBox13.Checked = True Then
                Dim indexChecked As Integer
                For Each indexChecked In Form9.CheckedListBox1.CheckedIndices
                    If indexChecked = 0 Then
                        ' Title
                        TextBox1.Text = TextBox1.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 1 Then
                        ' Title File as
                        TextBox16.Text = TextBox16.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 2 Then
                        ' Creator1
                        TextBox2.Text = TextBox2.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 3 Then
                        ' Creator1 File as
                        TextBox12.Text = TextBox12.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 4 Then
                        ' Creator2
                        TextBox3.Text = TextBox3.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 5 Then
                        ' Creator2 File as
                        TextBox13.Text = TextBox13.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 6 Then
                        ' Series
                        TextBox15.Text = TextBox15.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 7 Then
                        ' Series index
                        TextBox14.Text = TextBox14.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 8 Then
                        ' Description
                        TextBox4.Text = TextBox4.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 9 Then
                        ' Publisher
                        TextBox5.Text = TextBox5.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 10 Then
                        ' Date
                        TextBox6.Text = TextBox6.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 11 Then
                        ' Subject
                        TextBox17.Text = TextBox17.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 12 Then
                        ' Type
                        TextBox7.Text = TextBox7.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 13 Then
                        ' Format
                        TextBox8.Text = TextBox8.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 14 Then
                        ' Identifier
                        TextBox9.Text = TextBox9.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 15 Then
                        ' Source
                        TextBox10.Text = TextBox10.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                    If indexChecked = 16 Then
                        ' Language
                        TextBox11.Text = TextBox11.Text.Replace(Form9.TextBox1.Text, Form9.TextBox2.Text)
                    End If
                Next
            End If

            Application.DoEvents()

            ' Do the cover fixes now
            If CheckBox11.Checked = True Then
                If ((Button35.Visible) And (Form10.CheckBox1.Checked)) Then
                    Button35_Click(sender, e)
                End If
                If ((Button1.Visible) And (Form10.CheckBox2.Checked)) Then
                    Button1_Click(sender, e)
                End If
                If ((Button27.Visible) And (Form10.CheckBox3.Checked)) Then
                    Button27_Click(sender, e)
                End If
            End If

            Application.DoEvents()

            ' Save file
            SaveEpub(ListBox1.Items(x - 1), False)

            ClearInterface()

        Next
        ProgressBar1.Value = 0
        ProgressBar1.Update()
        ProgressBar1.Visible = False
        projectchanged = False
        Button3.Enabled = False
        ClearInterface()
        DialogResult = MsgBox("All done!", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")

        If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
            'delete contents of temp directory
            DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
        End If

        projectchanged = False
        CaptionString = "EPUB Metadata Editor"
        Me.Text = CaptionString
    End Sub
    Private Function ReplaceField(ByVal OldString, ByVal ReplacementString) As String
        'Convert metadata into new filename
        Dim currpos, endpos, temppos, field, nextchar, insertText, NewField
        NewField = ""
        currpos = 0
        While (currpos < Len(ReplacementString))
            currpos = currpos + 1

            ' look for field marker
            If (Mid(ReplacementString, currpos, 1) = "%") Then
                If (Mid(ReplacementString, currpos + 1, 1) = "%") Then
                    ' found '%%' (replace with '%')
                    NewField = NewField + "%"
                    currpos = currpos + 1
                Else
                    ' look for end field marker
                    endpos = InStr(currpos + 1, ReplacementString, "%")
                    If (endpos <> 0) Then
                        ' end field marker found
                        field = Mid(ReplacementString, currpos + 1, endpos - currpos - 1)
                        insertText = ""
                        If field = "CurrentContents" Then
                            insertText = OldString
                        ElseIf field = "CURRENTCONTENTS" Then
                            insertText = OldString.ToUpper
                        ElseIf field = "currentcontents" Then
                            insertText = OldString.ToLower
                        ElseIf field = "CurrentContentsTitleCase" Then
                            insertText = TitleCase(OldString)
                        ElseIf field = "CurrentContentsTitleCaseAll" Then
                            insertText = StrConv(OldString, VbStrConv.ProperCase)
                        ElseIf field = "Creator" Then
                            insertText = TextBox2.Text
                        ElseIf field = "CreatorFileAs" Then
                            insertText = TextBox12.Text
                        ElseIf field = "CreatorSurnameOnly" Then
                            insertText = TextBox2.Text
                            If InStr(insertText, " ") <> 0 Then
                                temppos = Len(insertText)
                                nextchar = Mid(insertText, temppos, 1)
                                While (nextchar <> " ")
                                    If temppos = 1 Then
                                        GoTo errortext
                                    End If
                                    temppos = temppos - 1
                                    nextchar = Mid(insertText, temppos, 1)
                                End While
                                insertText = Mid(insertText, temppos + 1)
                                If (Mid(TextBox2.Text, temppos - 1, 1) = ",") Then
                                    insertText = TextBox2.Text
                                    temppos = 1
                                    nextchar = Mid(insertText, temppos, 1)
                                    While (nextchar <> ",")
                                        If temppos = Len(insertText) Then
                                            GoTo errortext
                                        End If
                                        temppos = temppos + 1
                                        nextchar = Mid(insertText, temppos, 1)
                                    End While
                                    insertText = Mid(insertText, 1, temppos - 1)
                                End If
                            End If
                        ElseIf field = "CreatorFirstInitial" Then
                            insertText = TextBox2.Text
                            If InStr(insertText, " ") <> 0 Then
                                temppos = Len(insertText)
                                nextchar = Mid(insertText, temppos, 1)
                                While (nextchar <> " ")
                                    If temppos = 1 Then
                                        GoTo errortext
                                    End If
                                    temppos = temppos - 1
                                    nextchar = Mid(insertText, temppos, 1)
                                End While
                                insertText = Mid(insertText, temppos + 1)
                                If (Mid(TextBox2.Text, temppos - 1, 1) = ",") Then
                                    insertText = TextBox2.Text
                                    temppos = 1
                                    nextchar = Mid(insertText, temppos, 1)
                                    While (nextchar <> ",")
                                        If temppos = Len(insertText) Then
                                            GoTo errortext
                                        End If
                                        temppos = temppos + 1
                                        nextchar = Mid(insertText, temppos, 1)
                                    End While
                                    insertText = Mid(insertText, 1, temppos - 1)
                                End If
                            End If
                            If Len(insertText) > 1 Then
                                insertText = Mid(insertText, 1, 1)
                            End If
                        ElseIf field = "Title" Then
                            insertText = TextBox1.Text
                        ElseIf field = "TitleFileAs" Then
                            insertText = TextBox16.Text
                        ElseIf field = "Series" Then
                            insertText = TextBox15.Text
                        ElseIf field = "SeriesIndex" Then
                            insertText = TextBox14.Text
                        ElseIf field = "Date" Then
                            insertText = TextBox6.Text
                        Else
                            insertText = ""
                        End If
errortext:
                        NewField = NewField + insertText
                        currpos = endpos
                    End If
                End If
            Else
                NewField = NewField + Mid(ReplacementString, currpos, 1)
            End If
        End While
        Return NewField
    End Function
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        ListBox1.Items.Clear()
        Button10.Enabled = False
        Button32.Enabled = False
        ClearInterface()
        Button39.Enabled = False
        Button40.Enabled = False
        Button41.Enabled = False
    End Sub

    Private Sub SaveImageAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveImageAsToolStripMenuItem.Click
        SaveFileDialog1.Filter = "Image files (*.jpg, *.jpeg, *.png) | *.jpg; *.jpeg; *.png|All files (*.*)|*.*"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.FileName = Path.GetFileName(coverimagefile)
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            System.IO.File.Copy(coverimagefile, SaveFileDialog1.FileName, True)
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        TextBox13.Text = TextBox3.Text
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim tempname, nextchar, newname As String
        Dim temppos As Integer

        If My.Computer.Keyboard.ShiftKeyDown Then
            tempname = TextBox3.Text
            If InStr(tempname, ",") <> 0 Then
                temppos = Len(tempname)
                nextchar = Mid(tempname, temppos, 1)
                While (nextchar <> ",")
                    If temppos = 1 Then
                        GoTo errortext
                    End If
                    temppos = temppos - 1
                    nextchar = Mid(tempname, temppos, 1)
                End While
                newname = Mid(tempname, temppos + 2) + " " + Mid(tempname, 1, temppos - 1)
                TextBox3.Text = newname
            End If
        Else
            tempname = TextBox3.Text
            If InStr(tempname, " ") <> 0 Then
                temppos = Len(tempname)
                If temppos = 0 Then GoTo errortext
                nextchar = Mid(tempname, temppos, 1)
                While (nextchar <> " ")
                    If temppos = 1 Then
                        GoTo errortext
                    End If
                    temppos = temppos - 1
                    nextchar = Mid(tempname, temppos, 1)
                End While

                newname = Mid(tempname, temppos + 1) + ", " + Mid(tempname, 1, temppos - 1)
                TextBox13.Text = newname
            Else
                TextBox13.Text = tempname
            End If
        End If
errortext:
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        TextBox3.Text = TextBox13.Text
    End Sub

    Private Sub ChangeImageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeImageToolStripMenuItem.Click
        Dim oldextension, extension As String
        OpenFileDialog2.Filter = "Image files (*.jpg, *.jpeg, *.png) | *.jpg; *.jpeg; *.png|All files (*.*)|*.*"
        OpenFileDialog2.FilterIndex = 1
        OpenFileDialog2.FileName = ""
        If OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            oldextension = Path.GetExtension(coverimagefile)
            IO.File.Delete(coverimagefile)

            'find extension of new cover image
            extension = Path.GetExtension(OpenFileDialog2.FileName)

            If (extension <> oldextension) Then
                'convert new cover image to filetype of previous cover
                System.IO.File.Copy(OpenFileDialog2.FileName, coverimagefile, True)
                Dim image As Bitmap = Drawing.Image.FromFile(OpenFileDialog2.FileName)
                Dim encoderParameters As New EncoderParameters(1)
                encoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 100L)
                If oldextension = ".png" Then
                    image.Save(coverimagefile, System.Drawing.Imaging.ImageFormat.Png)
                Else
                    image.Save(coverimagefile, System.Drawing.Imaging.ImageFormat.Jpeg)
                End If
            Else
                IO.File.Copy(OpenFileDialog2.FileName, coverimagefile, True)
                wait(500)
            End If

            PictureBox1.ImageLocation = coverimagefile
            PictureBox1.Load()
            projectchanged = True
            Button3.Enabled = True
            Button27.Visible = True
            Label23.Visible = True
            Button42.Visible = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub AddImageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddImageToolStripMenuItem.Click
        Dim outputfile, metadatafile, newlineandspace, insertion, extension As String
        Dim startpos, insertpos As Integer

        OpenFileDialog2.Filter = "Image files (*.jpg, *.jpeg, *.png) | *.jpg; *.jpeg; *.png|All files (*.*)|*.*"
        OpenFileDialog2.FilterIndex = 1
        OpenFileDialog2.FileName = ""
        If OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
            'copy to outputfile
            extension = Path.GetExtension(OpenFileDialog2.FileName)
            outputfile = "cover" + extension
            coverimagefile = outputfile
            System.IO.File.Copy(OpenFileDialog2.FileName, opfdirectory + "\" + outputfile, True)
            wait(500)
            PictureBox1.ImageLocation = opfdirectory + "\" + outputfile
            PictureBox1.Load()

            'output html file
            RichTextBox1.Text = GetHTMLCoverFile(outputfile)
            RichTextBox1.SaveFile(opfdirectory + "\coverpage.xhtml", RichTextBoxStreamType.PlainText)

            'make changes to opf file
            RichTextBox1.Text = LoadUnicodeFile(opffile)
            metadatafile = LoadUnicodeFile(opffile)
            newlineandspace = Chr(10)

            If versioninfo = "3.0" Then
            Else
                'add item to metadata
                startpos = InStr(metadatafile, "</dc:title>")
                If startpos <> 0 Then
                    insertpos = InStr(startpos + 1, metadatafile, "<")
                    If insertpos <> 0 Then
                        newlineandspace = Mid(metadatafile, startpos + 11, insertpos - startpos - 11)
                        insertion = newlineandspace + "<meta name=" + Chr(34) + "cover" + Chr(34) + " content=" + Chr(34) + "cover" + Chr(34) + "/>" + newlineandspace
                        metadatafile = Mid(metadatafile, 1, startpos + 10) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                    End If
                End If
            End If

            'add items to manifest
            startpos = InStr(metadatafile, "<manifest")
            If startpos <> 0 Then
                insertpos = InStr(startpos + 1, metadatafile, "<")
                If insertpos <> 0 Then
                    newlineandspace = Mid(metadatafile, startpos + 10, insertpos - startpos - 10)
                    If extension = ".png" Then
                        insertion = newlineandspace + "<item href=" + Chr(34) + outputfile + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/png" + Chr(34) + "/>" + newlineandspace
                    Else
                        insertion = newlineandspace + "<item href=" + Chr(34) + outputfile + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>" + newlineandspace
                    End If
                    insertion = insertion + "<item href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " id=" + Chr(34) + "coverpage" + Chr(34) + " media-type=" + Chr(34) + "application/xhtml+xml" + Chr(34) + "/>" + newlineandspace
                    metadatafile = Mid(metadatafile, 1, startpos + 9) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                End If
            End If

            'add item to spine
            startpos = InStr(metadatafile, "<spine")
            If startpos <> 0 Then
                insertpos = InStr(startpos + 1, metadatafile, "<")
                If insertpos <> 0 Then
                    startpos = insertpos
                    While (Mid(metadatafile, startpos, 1) <> Chr(10))
                        startpos = startpos - 1
                        If startpos = 0 Then Exit While
                    End While
                    If startpos = 0 Then
                        startpos = insertpos
                        newlineandspace = ""
                    Else
                        newlineandspace = Mid(metadatafile, startpos, insertpos - startpos)
                    End If
                    insertion = "<itemref idref=" + Chr(34) + "coverpage" + Chr(34) + "/>"
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + newlineandspace + insertion + newlineandspace + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                End If
            End If

            If versioninfo = "3.0" Then
            Else
                'add reference to guide
                startpos = InStr(metadatafile, "<guide>")
                If startpos <> 0 Then
                    insertpos = InStr(startpos + 1, metadatafile, "<")
                    If insertpos <> 0 Then
                        newlineandspace = Mid(metadatafile, startpos + 7, insertpos - startpos - 7)
                        insertion = newlineandspace + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + newlineandspace
                        metadatafile = Mid(metadatafile, 1, startpos + 6) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                    End If
                Else
                    'look for <guide/> and delete it
                    metadatafile = metadatafile.Replace("<guide/>", "")

                    'now create <guide>
                    startpos = InStr(metadatafile, "</package>")
                    If startpos <> 0 Then
                        insertion = "<guide>" + newlineandspace + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + newlineandspace + "</guide>" + newlineandspace
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + insertion + Mid(metadatafile, startpos, Len(metadatafile) - startpos + 1)
                    End If
                End If
            End If

            'save opf file
            SaveUnicodeFile(opffile, metadatafile)

            ' Need to update metadata
            RichTextBox1.Text = LoadUnicodeFile(opffile)
            metadatafile = LoadUnicodeFile(opffile)
            ExtractMetadata(metadatafile, True)

            'update interface
            projectchanged = True
            Button3.Enabled = True
            Button27.Visible = True
            Label23.Visible = True
            Button42.Visible = True
            SaveImageAsToolStripMenuItem.Enabled = True
            ChangeImageToolStripMenuItem.Enabled = True
            AddImageToolStripMenuItem.Enabled = False
            Me.Text = "*" + CaptionString

            GroupBox1.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'output html file
        RichTextBox1.Text = GetHTMLCoverFile(relativecoverimagefile)
        RichTextBox1.SaveFile(coverfile, RichTextBoxStreamType.PlainText)
        Button1.Visible = False
        Label4.Visible = False
        If ((Button27.Visible = False) And (Button35.Visible = False)) Then
            Button42.Visible = False
        End If
        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString
    End Sub
    Private Function GetHTMLCoverFile(ByVal imagefile As String) As String
        Dim returnstring As String
        returnstring = "<?xml version='1.0' encoding='utf-8'?>" + Chr(13) + Chr(10)
        returnstring = returnstring + "<html xmlns=" + Chr(34) + "http://www.w3.org/1999/xhtml" + Chr(34) + " xml:lang=" + Chr(34) + "en" + Chr(34) + ">" + Chr(13) + Chr(10)
        returnstring = returnstring + "   <head>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       <meta http-equiv=" + Chr(34) + "Content-Type" + Chr(34) + " content=" + Chr(34) + "text/html; charset=UTF-8" + Chr(34) + "/>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       <meta name=" + Chr(34) + "calibre:cover" + Chr(34) + " content=" + Chr(34) + "true" + Chr(34) + "/>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       <title>Cover</title>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       <style type=" + Chr(34) + "text/css" + Chr(34) + " title=" + Chr(34) + "override_css" + Chr(34) + ">" + Chr(13) + Chr(10)
        returnstring = returnstring + "           @page {padding: 0pt; margin:0pt}" + Chr(13) + Chr(10)
        returnstring = returnstring + "           body { text-align: center; padding:0pt; margin: 0pt }" + Chr(13) + Chr(10)
        returnstring = returnstring + "           div { padding:0pt; margin: 0pt }" + Chr(13) + Chr(10)
        returnstring = returnstring + "           img { padding:0pt; margin: 0pt }" + Chr(13) + Chr(10)
        returnstring = returnstring + "       </style>" + Chr(13) + Chr(10)
        returnstring = returnstring + "   </head>" + Chr(13) + Chr(10)
        returnstring = returnstring + "   <body>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       <div>" + Chr(13) + Chr(10)
        returnstring = returnstring + "           <img src=" + Chr(34) + imagefile + Chr(34) + " alt=" + Chr(34) + "cover" + Chr(34) + " style=" + Chr(34) + "height: 100%" + Chr(34) + "/>" + Chr(13) + Chr(10)
        returnstring = returnstring + "       </div>" + Chr(13) + Chr(10)
        returnstring = returnstring + "   </body>" + Chr(13) + Chr(10)
        returnstring = returnstring + "</html>"
        Return returnstring
    End Function
    Private Function FileInUse(ByVal sFile As String) As Boolean
        If System.IO.File.Exists(sFile) Then
            Try
                Dim F As Short = FreeFile()
                FileOpen(F, sFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(F)
            Catch
                Return True
            End Try
        End If
    End Function

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim tempfilename As String
        Dim fileCheck As String
        fileCheck = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor\EPubMetadataEditor.ini"
        If (System.IO.File.Exists(fileCheck) = False) Then
            fileCheck = Application.StartupPath() + "\EPubMetadataEditor.ini"
        End If
        Dim objIniFile As New IniFile(fileCheck)
        Dim ViewerPath As String = _
            objIniFile.GetString("Viewer", "Path", "(none)")
        If ViewerPath <> "(none)" Then
            If ((My.Computer.Keyboard.ShiftKeyDown) Or (Not projectchanged)) Then
                Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process()
                myProcess.StartInfo.FileName = ViewerPath
                myProcess.StartInfo.Arguments = Chr(34) + OpenFileDialog1.FileName + Chr(34)
                myProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                myProcess.Start()
                myProcess.WaitForExit()
                myProcess.Close()
            Else
                'store current content of OPF file
                Dim tempstring = LoadUnicodeFile(opffile)

                'save any OPF changes to OPF file
                SaveEpub(OpenFileDialog1.FileName, True)

                Try
                    Dim zip As ZipStorer
                    tempfilename = tempdirectory & "\" & System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    zip = ZipStorer.Create(tempfilename, "")
                    Dim dir = Directory.GetDirectories(ebookdirectory)
                    Dim item As String
                    For Each item In dir
                        zip.AddDirectory(ZipStorer.Compression.Deflate, item, "", "")
                    Next
                    Dim files = Directory.GetFiles(ebookdirectory)
                    For Each item In files
                        zip.AddFile(ZipStorer.Compression.Deflate, item, Path.GetFileName(item), "")
                    Next
                    zip.Close()
                    Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process()
                    myProcess.StartInfo.FileName = ViewerPath
                    myProcess.StartInfo.Arguments = Chr(34) + tempfilename + Chr(34)
                    myProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                    myProcess.Start()
                    myProcess.WaitForExit()
                    myProcess.Close()

                    System.IO.File.Delete(tempfilename)
                Catch ex As Exception
                    DialogResult = MsgBox("There was a problem saving current file to a temporary file.  Original file will be viewed instead.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process()
                    myProcess.StartInfo.FileName = ViewerPath
                    myProcess.StartInfo.Arguments = Chr(34) + OpenFileDialog1.FileName + Chr(34)
                    myProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                    myProcess.Start()
                    myProcess.WaitForExit()
                    myProcess.Close()
                End Try

                'restore current content of OPF file
                SaveUnicodeFile(opffile, tempstring)
            End If
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        'Me.Width = 1257
        Me.ClientSize = New System.Drawing.Size(1250, 640)
        Button16.Visible = False
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        'Me.Width = 913
        Me.ClientSize = New System.Drawing.Size(905, 640)
        Button16.Visible = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim viewerfilename, inidirectory, inifilename As String
        inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
        inifilename = inidirectory + "\EPubMetadataEditor.ini"
        If System.IO.File.Exists(inifilename) = False Then
            If System.IO.Directory.Exists(inidirectory) = False Then
                System.IO.Directory.CreateDirectory(inidirectory)
            End If
            Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
            Dim s As New StreamWriter(fs)
            s.WriteLine("[Viewer]")
            s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
            s.Close()
        End If

        Dim objIniFile As New IniFile(inifilename)
        viewerfilename = objIniFile.GetString("Viewer", "Path", "(none)")
        If viewerfilename <> "(none)" Then
            OpenFileDialog4.FileName = Path.GetFileName(viewerfilename)
            OpenFileDialog4.InitialDirectory = Path.GetDirectoryName(viewerfilename)
        Else
            ' look for ini file in old location
            Dim tempinifile = Application.StartupPath() + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(tempinifile) = True Then
                Dim tempobjIniFile As New IniFile(tempinifile)
                viewerfilename = tempobjIniFile.GetString("Viewer", "Path", "(none)")
                If viewerfilename <> "(none)" Then
                    OpenFileDialog4.FileName = Path.GetFileName(viewerfilename)
                    OpenFileDialog4.InitialDirectory = Path.GetDirectoryName(viewerfilename)
                Else
                    OpenFileDialog4.FileName = ""
                End If
            Else
                OpenFileDialog4.FileName = ""
            End If
        End If

        OpenFileDialog4.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
        OpenFileDialog4.FilterIndex = 1

        Dim myDialogResult = OpenFileDialog4.ShowDialog

        If myDialogResult = Windows.Forms.DialogResult.OK Then
            objIniFile.WriteString("Viewer", "Path", Chr(34) + OpenFileDialog4.FileName + Chr(34))
            LinkLabel1.Text = "Change external editor"
            Button8.Enabled = Button23.Enabled
        ElseIf myDialogResult = Windows.Forms.DialogResult.Cancel Then
            If viewerfilename <> "(none)" Then
                ' save old ini information to the new location
                objIniFile.WriteString("Viewer", "Path", Chr(34) + viewerfilename + Chr(34))
            End If
        End If
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        TextBox16.Text = ""
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim metadatafile As String

        fileeditorreturn = False
        filecontents = ""
        Form2.Button3.Visible = True
        Form2.Button4.Visible = True
        Form2.Button5.Visible = True
        Form2.Button9.Visible = False
        Form2.CheckBox2.Visible = True
        Form2.RichTextBox1.Text = LoadUnicodeFile(opffile)
        Form2.ShowDialog()
        If fileeditorreturn = True Then
            keepcombobox = True
            ClearInterface()
            keepcombobox = False
            RichTextBox1.Text = filecontents
            SaveUnicodeFile(opffile, RichTextBox1.Text)
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString

            ' Need to update metadata
            RichTextBox1.Text = LoadUnicodeFile(opffile)
            metadatafile = LoadUnicodeFile(opffile)
            ExtractMetadata(metadatafile, True)
            refreshfilelist = False
        End If
        Form2.Button4.Visible = False
        Form2.Button5.Visible = False
        Form2.CheckBox2.Visible = False
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        fileeditorreturn = False
        filecontents = ""
        Form2.Button3.Visible = False
        Form2.Button9.Visible = False
        Form2.RichTextBox1.Text = LoadUnicodeFile(tocfile)
        Form2.ShowDialog()
        If fileeditorreturn = True Then
            RichTextBox1.Text = filecontents
            SaveUnicodeFile(tocfile, RichTextBox1.Text)
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        TextBox16.Text = TextBox1.Text
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        TextBox1.Text = TextBox16.Text
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        If versioninfo = "3.0" Then
            Dim myuri As New Uri(tocfile)
            Form5.WebBrowser1.Url = myuri
            Button23.Enabled = False
            Form5.Show()
        Else
            ProcessTOCNCX(tocfile)
            Form3.Show()
        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        fileeditorreturn = False
        filecontents = ""
        Form2.Button3.Visible = False
        Form2.Button9.Visible = False
        Form2.RichTextBox1.LoadFile(pagemapfile, RichTextBoxStreamType.PlainText)
        Form2.ShowDialog()
        If fileeditorreturn = True Then
            RichTextBox1.Text = filecontents
            RichTextBox1.SaveFile(pagemapfile, RichTextBoxStreamType.PlainText)
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If My.Computer.Keyboard.ShiftKeyDown Then
            TextBox1.Text = StrConv(TextBox1.Text, VbStrConv.ProperCase)
        Else
            TextBox1.Text = TitleCase(TextBox1.Text)
        End If
    End Sub

    Private Sub PictureBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PictureBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PictureBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PictureBox1.DragDrop
        Dim MyFiles() As String
        Dim i As Integer
        Dim metadatafile, newlineandspace, insertion, oldextension, extension As String
        Dim startpos, insertpos As Integer
        Dim mime As String

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Assign the file(s) to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
        Else
            Exit Sub
        End If

        ' Loop through the array
        For i = 0 To MyFiles.Length - 1
            mime = GetMimeType(MyFiles(i))
            If (mime.StartsWith("image")) Then
                If (MyFiles(i) <> coverimagefile) Then
                    If IO.File.Exists(coverimagefile) Then
                        oldextension = Path.GetExtension(coverimagefile)
                        IO.File.Delete(coverimagefile)

                        'find extension of new cover image
                        extension = Path.GetExtension(MyFiles(i))

                        If (extension <> oldextension) Then
                            'convert new cover image to filetype of previous cover
                            System.IO.File.Copy(MyFiles(i), coverimagefile, True)
                            Dim image As Bitmap = Drawing.Image.FromFile(MyFiles(i))
                            Dim encoderParameters As New EncoderParameters(1)
                            encoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 100L)
                            If oldextension = ".png" Then
                                image.Save(coverimagefile, System.Drawing.Imaging.ImageFormat.Png)
                            Else
                                image.Save(coverimagefile, System.Drawing.Imaging.ImageFormat.Jpeg)
                            End If
                        Else
                            IO.File.Copy(MyFiles(i), coverimagefile, True)
                            wait(500)
                        End If
                    Else
                        'find extension of new cover image
                        extension = Path.GetExtension(MyFiles(i))

                        'need to create a new cover file
                        coverimagefile = opfdirectory + "\cover" + extension

                        'output html file
                        RichTextBox1.Text = GetHTMLCoverFile("cover" + extension)
                        RichTextBox1.SaveFile(opfdirectory + "\coverpage.xhtml", RichTextBoxStreamType.PlainText)

                        'make changes to opf file
                        RichTextBox1.Text = LoadUnicodeFile(opffile)
                        metadatafile = LoadUnicodeFile(opffile)
                        newlineandspace = Chr(10)

                        If versioninfo = "3.0" Then
                        Else
                            'add item to metadata
                            startpos = InStr(metadatafile, "</dc:title>")
                            If startpos <> 0 Then
                                insertpos = InStr(startpos + 1, metadatafile, "<")
                                If insertpos <> 0 Then
                                    newlineandspace = Mid(metadatafile, startpos + 11, insertpos - startpos - 11)
                                    insertion = newlineandspace + "<meta name=" + Chr(34) + "cover" + Chr(34) + " content=" + Chr(34) + "cover" + Chr(34) + "/>" + newlineandspace
                                    metadatafile = Mid(metadatafile, 1, startpos + 10) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                                End If
                            End If
                        End If

                        'add items to manifest
                        startpos = InStr(metadatafile, "<manifest")
                        If startpos <> 0 Then
                            insertpos = InStr(startpos + 1, metadatafile, "<")
                            If insertpos <> 0 Then
                                newlineandspace = Mid(metadatafile, startpos + 10, insertpos - startpos - 10)
                                If extension = ".png" Then
                                    insertion = newlineandspace + "<item href=" + Chr(34) + "cover" + extension + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/png" + Chr(34) + "/>" + newlineandspace
                                Else
                                    insertion = newlineandspace + "<item href=" + Chr(34) + "cover" + extension + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>" + newlineandspace
                                End If
                                insertion = insertion + "<item href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " id=" + Chr(34) + "coverpage" + Chr(34) + " media-type=" + Chr(34) + "application/xhtml+xml" + Chr(34) + "/>" + newlineandspace
                                metadatafile = Mid(metadatafile, 1, startpos + 9) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                            End If
                        End If

                        'add item to spine
                        startpos = InStr(metadatafile, "<spine")
                        If startpos <> 0 Then
                            insertpos = InStr(startpos + 1, metadatafile, "<")
                            If insertpos <> 0 Then
                                startpos = insertpos
                                While (Mid(metadatafile, startpos, 1) <> Chr(10))
                                    startpos = startpos - 1
                                    If startpos = 0 Then Exit While
                                End While
                                If startpos = 0 Then
                                    startpos = insertpos
                                    newlineandspace = ""
                                Else
                                    newlineandspace = Mid(metadatafile, startpos, insertpos - startpos)
                                End If
                                insertion = "<itemref idref=" + Chr(34) + "coverpage" + Chr(34) + "/>"
                                metadatafile = Mid(metadatafile, 1, startpos - 1) + newlineandspace + insertion + newlineandspace + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                            End If
                        End If

                        If versioninfo = "3.0" Then
                        Else
                            'add reference to guide
                            startpos = InStr(metadatafile, "<guide>")
                            If startpos <> 0 Then
                                insertpos = InStr(startpos + 1, metadatafile, "<")
                                If insertpos <> 0 Then
                                    newlineandspace = Mid(metadatafile, startpos + 7, insertpos - startpos - 7)
                                    insertion = newlineandspace + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + newlineandspace
                                    metadatafile = Mid(metadatafile, 1, startpos + 6) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                                End If
                            Else
                                'look for <guide/> and delete it
                                metadatafile = metadatafile.Replace("<guide/>", "")

                                'now create <guide>
                                startpos = InStr(metadatafile, "</package>")
                                If startpos <> 0 Then
                                    insertion = "<guide>" + newlineandspace + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + newlineandspace + "</guide>" + newlineandspace
                                    metadatafile = Mid(metadatafile, 1, startpos - 1) + insertion + Mid(metadatafile, startpos, Len(metadatafile) - startpos + 1)
                                End If
                            End If
                        End If

                        'save opf file
                        SaveUnicodeFile(opffile, metadatafile)

                        IO.File.Copy(MyFiles(i), coverimagefile, True)
                        wait(500)
                    End If


                    ' update interface
                    RichTextBox1.Text = LoadUnicodeFile(opffile)
                    metadatafile = LoadUnicodeFile(opffile)
                    ExtractMetadata(metadatafile, True)
                    ChangeImageToolStripMenuItem.Enabled = True
                    AddImageToolStripMenuItem.Enabled = False
                    projectchanged = True
                    Button3.Enabled = True
                    Button27.Visible = True
                    Label23.Visible = True
                    Button42.Visible = True
                    Me.Text = "*" + CaptionString
                End If
            End If
        Next
    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        If Not PictureBox1.ImageLocation Is Nothing Then
            ' Set a flag to show that the mouse is down.
            m_MouseIsDown = True
        End If
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        If m_MouseIsDown Then
            ' Initiate dragging
            Dim DropList As New System.Collections.Specialized.StringCollection
            Dim DragPaths As New DataObject()
            DropList.Add(PictureBox1.ImageLocation)
            DragPaths.SetFileDropList(DropList)
            PictureBox1.DoDragDrop(DragPaths, DragDropEffects.Copy)
        End If
        m_MouseIsDown = False
    End Sub

    Private Sub PasteImageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteImageToolStripMenuItem.Click
        Dim pic As Image
        Dim metadatafile, newlineandspace, insertion As String
        Dim startpos, insertpos As Integer

        If Clipboard.ContainsImage Then
            pic = Clipboard.GetImage
            If coverimagefile = "" Then
                'need to create a new cover file
                coverimagefile = opfdirectory + "\cover.jpg"

                'output html file
                RichTextBox1.Text = GetHTMLCoverFile("cover.jpg")
                RichTextBox1.SaveFile(opfdirectory + "\coverpage.xhtml", RichTextBoxStreamType.PlainText)

                'make changes to opf file
                RichTextBox1.Text = LoadUnicodeFile(opffile)
                metadatafile = LoadUnicodeFile(opffile)

                If versioninfo = "3.0" Then
                Else
                    'add item to metadata
                    startpos = InStr(metadatafile, "</dc:title>")
                    If startpos <> 0 Then
                        insertpos = InStr(startpos + 1, metadatafile, "<")
                        If insertpos <> 0 Then
                            newlineandspace = Mid(metadatafile, startpos + 11, insertpos - startpos - 11)
                            insertion = newlineandspace + "<meta name=" + Chr(34) + "cover" + Chr(34) + " content=" + Chr(34) + "cover" + Chr(34) + "/>" + newlineandspace
                            metadatafile = Mid(metadatafile, 1, startpos + 10) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                        End If
                    End If
                End If

                'add items to manifest
                startpos = InStr(metadatafile, "<manifest")
                If startpos <> 0 Then
                    insertpos = InStr(startpos + 1, metadatafile, "<")
                    If insertpos <> 0 Then
                        newlineandspace = Mid(metadatafile, startpos + 10, insertpos - startpos - 10)
                        insertion = newlineandspace + "<item href=" + Chr(34) + "cover.jpg" + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>" + newlineandspace
                        insertion = insertion + "<item href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " id=" + Chr(34) + "coverpage" + Chr(34) + " media-type=" + Chr(34) + "application/xhtml+xml" + Chr(34) + "/>" + newlineandspace
                        metadatafile = Mid(metadatafile, 1, startpos + 9) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                    End If
                End If

                'add item to spine
                startpos = InStr(metadatafile, "<spine")
                If startpos <> 0 Then
                    insertpos = InStr(startpos + 1, metadatafile, "<")
                    If insertpos <> 0 Then
                        startpos = insertpos
                        While (Mid(metadatafile, startpos, 1) <> Chr(10))
                            startpos = startpos - 1
                        End While
                        newlineandspace = Mid(metadatafile, startpos, insertpos - startpos)
                        insertion = "<itemref idref=" + Chr(34) + "coverpage" + Chr(34) + "/>"
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + newlineandspace + insertion + newlineandspace + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                    End If
                End If

                If versioninfo = "3.0" Then
                Else
                    'add reference to guide
                    startpos = InStr(metadatafile, "<guide>")
                    If startpos <> 0 Then
                        insertpos = InStr(startpos + 1, metadatafile, "<")
                        If insertpos <> 0 Then
                            newlineandspace = Mid(metadatafile, startpos + 7, insertpos - startpos - 7)
                            insertion = newlineandspace + "<reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + newlineandspace
                            metadatafile = Mid(metadatafile, 1, startpos + 6) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
                        End If
                    Else
                        'look for <guide/> and delete it
                        metadatafile = metadatafile.Replace("<guide/>", "")

                        'now create <guide>
                        startpos = InStr(metadatafile, "</package>")
                        If startpos <> 0 Then
                            insertion = "<guide>" + Chr(13) + Chr(10) + "    <reference href=" + Chr(34) + "coverpage.xhtml" + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + "</guide>" + Chr(13) + Chr(10)
                            metadatafile = Mid(metadatafile, 1, startpos - 1) + insertion + Mid(metadatafile, startpos, Len(metadatafile) - startpos + 1)
                        End If
                    End If
                End If

                'save opf file
                SaveUnicodeFile(opffile, metadatafile)
            End If

            'save image to file
            pic.Save(coverimagefile)
            wait(500)

            'update interface
            RichTextBox1.Text = LoadUnicodeFile(opffile)
            metadatafile = LoadUnicodeFile(opffile)
            ExtractMetadata(metadatafile, True)
            ChangeImageToolStripMenuItem.Enabled = True
            AddImageToolStripMenuItem.Enabled = False
            projectchanged = True
            Button3.Enabled = True
            Button27.Visible = True
            Label23.Visible = True
            Button42.Visible = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then CheckBox4.Checked = False
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then CheckBox1.Checked = False
    End Sub
    Private Function CleanOPF(ByVal metadatafile As String) As String
        Dim dcnamespace As String
        Dim startpos, namespacelen, endpos, extracheck, temppos As Integer

        Try
            'Check for corrupted xml first line
            endpos = InStr(metadatafile, Chr(10))
            Dim tempstring As String = Mid(metadatafile, 1, endpos)
            If InStr(tempstring, " xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34)) Then
                tempstring = tempstring.Replace(" xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34), "")
                metadatafile = tempstring + Mid(metadatafile, endpos + 1)
            End If

            'Check for non-standard dc namespace tags
            startpos = InStr(metadatafile, "=" + Chr(34) + "http://purl.org/dc/elements/1.1/")
            If ((startpos <> 0) And (startpos < endpos)) Then
                ' work backwards to find the xmlns definition
                namespacelen = 0
                While (startpos - namespacelen <> 0)
                    namespacelen = namespacelen + 1
                    If Mid(metadatafile, startpos - namespacelen, 6) = "xmlns:" Then
                        Exit While
                    End If
                End While
                If namespacelen < startpos Then
                    dcnamespace = Mid(metadatafile, startpos - namespacelen + 6, namespacelen - 6)
                    metadatafile = metadatafile.Replace(dcnamespace + ":", "dc:")
                    metadatafile = metadatafile.Replace("xmlns:" + dcnamespace, "xmlns:dc")
                End If
            Else
                metadatafile = metadatafile.Replace(" xmlns=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34), "")
            End If

            'Check for non-standard opf namespace tags
            If (InStr(metadatafile, "<opf:metadata") Or InStr(metadatafile, "<opf:manifest")) Then
                metadatafile = metadatafile.Replace("<opf:", "<")
                metadatafile = metadatafile.Replace("</opf:", "</")
            End If
            If (InStr(metadatafile, "<ns0:metadata") Or InStr(metadatafile, "<ns0:manifest")) Then
                metadatafile = metadatafile.Replace("<ns0:", "<")
                metadatafile = metadatafile.Replace("<ns1:", "<dc:")
                metadatafile = metadatafile.Replace("</ns0:", "</")
                metadatafile = metadatafile.Replace("</ns1:", "</dc:")
            End If

            'Search for multiple xmlns:dc="http://purl.org/dc/elements/1.1/"
            metadatafile = metadatafile.Replace(" xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34), "")
            startpos = InStr(metadatafile, "<metadata")
            metadatafile = Mid(metadatafile, 1, startpos + 8) + " xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + Mid(metadatafile, startpos + 9)

            'Search for xmlns:opf="http://www.idpf.org/2007/opf"
            startpos = InStr(metadatafile, "xmlns:opf=" + Chr(34) + "http://www.idpf.org/2007/opf" + Chr(34))
            temppos = InStr(metadatafile, "<dc:")
            If ((startpos = 0) Or (startpos > temppos)) Then
                metadatafile = metadatafile.Replace(" xmlns:opf=" + Chr(34) + "http://www.idpf.org/2007/opf" + Chr(34), "")
                'Add it to <metadata > tag
                startpos = InStr(metadatafile, "<metadata")
                startpos = InStr(startpos, metadatafile, ">") - 1
                extracheck = InStr(metadatafile, "xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34))
                If extracheck = 0 Then
                    metadatafile = Mid(metadatafile, 1, startpos) + " xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " xmlns:opf=" + Chr(34) + "http://www.idpf.org/2007/opf" + Chr(34) + Mid(metadatafile, startpos + 1)
                Else
                    metadatafile = Mid(metadatafile, 1, startpos) + " xmlns:opf=" + Chr(34) + "http://www.idpf.org/2007/opf" + Chr(34) + Mid(metadatafile, startpos + 1)
                End If
            End If

            'Search for xmlns:calibre="http://calibre.kovidgoyal.net/2009/metadata"
            startpos = InStr(metadatafile, "xmlns:calibre=" + Chr(34) + "http://calibre.kovidgoyal.net/2009/metadata" + Chr(34))
            temppos = InStr(metadatafile, "calibre:series")
            If ((startpos = 0) And ((temppos > 0) Or (TextBox15.Text <> ""))) Then
                'Add it to <metadata > tag
                startpos = InStr(metadatafile, "<metadata")
                startpos = InStr(startpos, metadatafile, ">") - 1
                metadatafile = Mid(metadatafile, 1, startpos) + " xmlns:calibre=" + Chr(34) + "http://calibre.kovidgoyal.net/2009/metadata" + Chr(34) + Mid(metadatafile, startpos + 1)
            End If

            Return metadatafile
        Catch ex As Exception
            MsgBox("ERROR: opf file is corrupt. You will need to edit this file manually before you can save this EPUB file.", MsgBoxStyle.Critical, "EPUB Metadata Editor")
            Return metadatafile
        End Try
    End Function
    Private Sub SaveEpub(ByVal EpubFileName As String, ByVal SaveOPFOnly As Boolean)
        Dim metadatafile, optionaltext, optionaltext2 As String
        Dim startpos, endtag, endpos, extracheck, lenheader, checktag, lookforID, endID As Integer
        Dim temporarydirectory, newheader, ID As String
        Dim idpos, temploop, temppos, startheaderpos, endheaderpos, refinespos, testpos, extrachars As Integer
        Dim idinfo, rolestring, identifierscheme, temptext As String
        Dim creatorfileasplaced, creatorroleplaced, creator2fileasplaced, creator2roleplaced, schemeplaced, opfmeta, founduuid As Boolean
        Dim tocfiletext As String
        Dim tocstartpos, toccontentloc, toccontentend As Integer

        Dim fi As New FileInfo(EpubFileName)

        If Not SaveOPFOnly Then
            If fi.IsReadOnly Then
                MsgBox("ERROR: File is read-only!  Cannot save changes.", MsgBoxStyle.Critical, "EPUB Metadata Editor")
                Exit Sub
            End If

            If FileInUse(fi.FullName) Then
                MsgBox("ERROR: File in use!  Cannot save changes.", MsgBoxStyle.Critical, "EPUB Metadata Editor")
                Exit Sub
            End If
        End If

        RichTextBox1.Text = LoadUnicodeFile(opffile)
        metadatafile = LoadUnicodeFile(opffile)
        'Get rid of empty creator field (if there is one creator field and one empty creator fields, the empty one causes problems)
        If (InStr(metadatafile, "<dc:creator />")) Then
            metadatafile = metadatafile.Replace("<dc:creator />", "")
        End If
        metadatafile = CleanOPF(metadatafile)
        metadatafile = Regularise(metadatafile)

        'Output title
        startpos = InStr(metadatafile, "<dc:title")
        If startpos <> 0 Then
            endpos = InStr(metadatafile, "</dc:title>")
            If endpos = 0 Then
                endpos = InStr(startpos, metadatafile, " />")
                extrachars = 3
                If endpos = 0 Then
                    InStr(startpos, metadatafile, "/>")
                    extrachars = 2
                End If
                If endpos <> 0 Then
                    metadatafile = Mid(metadatafile, 1, endpos + 2 - extrachars) + "</dc:title>" + Mid(metadatafile, endpos + extrachars)
                Else
                    DialogResult = MsgBox("Badly formed OPF file.  EPUB Metadata Editor is unable to save this file.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Exit Sub
                End If
            End If
            lenheader = Len("<dc:title")

            'If optional attributes
            If ((TextBox16.Text <> "") And (versioninfo <> "3.0")) Then
                optionaltext = " opf:file-as=" + Chr(34) + XMLOutput(TextBox16.Text) + Chr(34) + ">"
            Else
                optionaltext = ">"
            End If
            metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + optionaltext + XMLOutput(TextBox1.Text) + Mid(metadatafile, endpos)
        Else
            ' no title yet, so add it after <metadata... > tag
            startpos = InStr(metadatafile, "<metadata")
            startpos = InStr(startpos, metadatafile, ">") + 1
            'If optional attributes
            If ((TextBox16.Text <> "") And (versioninfo <> "3.0")) Then
                optionaltext = " opf:file-as=" + Chr(34) + XMLOutput(TextBox16.Text) + Chr(34) + ">"
            Else
                optionaltext = ">"
            End If
            metadatafile = Mid(metadatafile, 1, startpos) + "  <dc:title" + optionaltext + XMLOutput(TextBox1.Text) + "</dc:title>" + Mid(metadatafile, startpos)
        End If
        ' Handle Title "file as" in EPUB3
        If ((TextBox16.Text <> "") And (versioninfo = "3.0")) Then
            ' Look for Calibre's title_sort meta tag
            startpos = InStr(metadatafile, "calibre:title_sort")
            If startpos <> 0 Then
                temppos = startpos
                startpos = InStrRev(metadatafile, "<meta ", startpos)
                opfmeta = False
                If startpos = 0 Then
                    startpos = InStrRev(metadatafile, "<opf:meta ", temppos)
                    opfmeta = True
                End If
                endpos = InStr(startpos + 9, metadatafile, "/>")
                If ((startpos <> 0) And (startpos < endpos)) Then
                    If opfmeta Then
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + "<opf:meta name=" + Chr(34) + "calibre:title_sort" + Chr(34) + " content=" + Chr(34) + XMLOutput(TextBox16.Text) + Chr(34) + Mid(metadatafile, endpos)
                    Else
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + "<meta name=" + Chr(34) + "calibre:title_sort" + Chr(34) + " content=" + Chr(34) + XMLOutput(TextBox16.Text) + Chr(34) + Mid(metadatafile, endpos)
                    End If
                Else
                    ' Need a new metatag
                    endpos = InStr(metadatafile, "</dc:title>") + Len("</dc:title>")
                    metadatafile = Mid(metadatafile, 1, endpos) + Chr(10) + "<meta name=" + Chr(34) + "calibre:title_sort" + Chr(34) + " content=" + Chr(34) + XMLOutput(TextBox16.Text) + Chr(34) + "/>" + Mid(metadatafile, endpos + 1)
                End If
            End If
        End If
        If ((TextBox16.Text = "") Or ((TextBox16.Text <> "") And (versioninfo = "2.0"))) Then
            ' Look for Calibre's title_sort meta tag and remove it
            startpos = InStr(metadatafile, "calibre:title_sort")
            If startpos <> 0 Then
                temppos = startpos
                startpos = InStrRev(metadatafile, "<meta ", startpos)
                If startpos = 0 Then startpos = InStrRev(metadatafile, "<opf:meta ", temppos)
                endpos = InStr(startpos, metadatafile, "/>")
                If ((startpos <> 0) And (startpos < endpos)) Then
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endpos + 3)
                End If
            End If
        End If

        'Output first creator
        startpos = InStr(metadatafile, "<dc:creator")
        If startpos <> 0 Then
            endheaderpos = InStr(startpos, metadatafile, ">")
            endpos = InStr(startpos, metadatafile, "</dc:creator>")
            If endpos = 0 Then
                endpos = InStr(startpos, metadatafile, " />")
                extrachars = 3
                If endpos = 0 Then
                    InStr(startpos, metadatafile, "/>")
                    extrachars = 2
                End If
                If endpos <> 0 Then
                    metadatafile = Mid(metadatafile, 1, endpos + 2 - extrachars) + "</dc:creator>" + Mid(metadatafile, endpos + extrachars)
                Else
                    DialogResult = MsgBox("Badly formed OPF file.  EPUB Metadata Editor is unable to save this file.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Exit Sub
                End If
            End If
            lenheader = Len("<dc:creator")
            If versioninfo = "3.0" Then
                'metadatafile = Mid(metadatafile, 1, endheaderpos) + XMLOutput(TextBox2.Text) + Mid(metadatafile, endpos)
                ' get id
                idpos = InStr(startpos, metadatafile, "id=")
                idinfo = ""
                If idpos <> 0 Then
                    If idpos < endpos Then
                        For temploop = idpos + 4 To endpos
                            If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                metadatafile = Mid(metadatafile, 1, startpos - 1) + "<dc:creator id=" + Chr(34) + idinfo + Chr(34) + ">" + XMLOutput(TextBox2.Text) + Mid(metadatafile, endpos)
                                GoTo lookforrefines
                            End If
                        Next
                    Else
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + "<dc:creator id=" + Chr(34) + "creator" + Chr(34) + ">" + XMLOutput(TextBox2.Text) + Mid(metadatafile, endpos)
                        idinfo = "creator"
                    End If
                Else
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + "<dc:creator id=" + Chr(34) + "creator" + Chr(34) + ">" + XMLOutput(TextBox2.Text) + Mid(metadatafile, endpos)
                    idinfo = "creator"
                End If
lookforrefines:
                If idinfo <> "" Then
                    temppos = InStr(startpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo)
                    creatorfileasplaced = False
                    creatorroleplaced = False
                    rolestring = "aut"
                    While temppos <> 0
                        startheaderpos = InStrRev(metadatafile, "<", temppos)
                        endheaderpos = InStr(temppos, metadatafile, ">")
                        endpos = InStr(temppos, metadatafile, "</meta>")
                        If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                        refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "file-as")
                        If refinespos <> 0 Then
                            If refinespos < endpos Then
                                metadatafile = Mid(metadatafile, 1, endheaderpos) + XMLOutput(TextBox12.Text) + Mid(metadatafile, endpos)
                                creatorfileasplaced = True
                            End If
                        End If
                        refinespos = InStr(startheaderpos, metadatafile, "property=" + Chr(34) + "role")
                        If refinespos <> 0 Then
                            If refinespos < endpos Then
                                If ComboBox1.SelectedIndex = 1 Then rolestring = "edt"
                                If ComboBox1.SelectedIndex = 2 Then rolestring = "ill"
                                If ComboBox1.SelectedIndex = 3 Then rolestring = "trl"
                                metadatafile = Mid(metadatafile, 1, endheaderpos) + rolestring + Mid(metadatafile, endpos)
                                creatorroleplaced = True
                            End If
                        End If
                        temppos = InStr(endpos, metadatafile, "refines=" + Chr(34) + "#" + idinfo)
                    End While
                    If (((creatorroleplaced = False) Or (creatorfileasplaced = False)) And (TextBox12.Text <> "")) Then
                        startpos = InStr(metadatafile, "</dc:creator>") + 13 'end of creator
                        If ((creatorfileasplaced = False) And (creatorroleplaced = False)) Then
                            metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + _
                            "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox12.Text) + "</meta>" + Chr(13) + Chr(10) + _
                            "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Mid(metadatafile, startpos)
                        ElseIf ((creatorfileasplaced = False) And (creatorroleplaced = True)) Then
                            metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox12.Text) + "</meta>" + Mid(metadatafile, startpos)
                        ElseIf ((creatorfileasplaced = True) And (creatorroleplaced = False)) Then
                            startpos = InStr(startpos, metadatafile, "</meta>")
                            metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Mid(metadatafile, startpos)
                        End If
                    End If
                End If

                'Output second creator?
                endpos = InStr(metadatafile, "</dc:creator>") 'find end of first creator
                startpos = InStr(endpos, metadatafile, "<dc:creator") 'look for another one
                If ((TextBox3.Text <> "") Or (startpos <> 0)) Then
                    If startpos <> 0 Then
                        endheaderpos = InStr(startpos, metadatafile, ">")
                        endpos = InStr(startpos, metadatafile, "</dc:creator>")
                        metadatafile = Mid(metadatafile, 1, endheaderpos) + XMLOutput(TextBox3.Text) + Mid(metadatafile, endpos)

                        ' get id
                        idpos = InStr(startpos, metadatafile, "id=")
                        idinfo = ""
                        If idpos <> 0 Then
                            For temploop = idpos + 4 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                    metadatafile = Mid(metadatafile, 1, startpos - 1) + "<dc:creator id=" + Chr(34) + idinfo + Chr(34) + ">" + XMLOutput(TextBox3.Text) + Mid(metadatafile, endpos)
                                    GoTo lookforrefines2
                                End If
                            Next
                        Else
                            metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator id=" + Chr(34) + "creator2" + Chr(34) + ">" + XMLOutput(TextBox3.Text) + Mid(metadatafile, endpos)
                            idinfo = "creator2"
                        End If
lookforrefines2:
                        If idinfo <> "" Then
                            temppos = InStr(startpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                            If temppos = 0 Then temppos = InStr(startpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                            creator2fileasplaced = False
                            creator2roleplaced = False
                            rolestring = "aut"
                            While temppos <> 0
                                endheaderpos = InStr(temppos, metadatafile, ">")
                                endpos = InStr(temppos, metadatafile, "</meta>")
                                If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                                refinespos = InStr(temppos, metadatafile, "property=" + Chr(34) + "file-as")
                                If refinespos <> 0 Then
                                    If refinespos < endpos Then
                                        metadatafile = Mid(metadatafile, 1, endheaderpos) + XMLOutput(TextBox13.Text) + Mid(metadatafile, endpos)
                                        creator2fileasplaced = True
                                    End If
                                End If
                                refinespos = InStr(temppos, metadatafile, "property=" + Chr(34) + "role")
                                If refinespos <> 0 Then
                                    If refinespos < endpos Then
                                        If ComboBox2.SelectedIndex = 1 Then rolestring = "edt"
                                        If ComboBox2.SelectedIndex = 2 Then rolestring = "ill"
                                        If ComboBox2.SelectedIndex = 3 Then rolestring = "trl"
                                        metadatafile = Mid(metadatafile, 1, endheaderpos) + rolestring + Mid(metadatafile, endpos)
                                        creator2roleplaced = True
                                    End If
                                End If
                                temppos = InStr(endpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                                If temppos = 0 Then temppos = InStr(endpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                            End While
                            If (((creator2roleplaced = False) Or (creator2fileasplaced = False)) And (TextBox13.Text <> "")) Then
                                startpos = InStr(metadatafile, "</dc:creator>")
                                startpos = InStr(startpos + 1, metadatafile, "</dc:creator>") + 13 'end of second creator
                                If ((creator2fileasplaced = False) And (creator2roleplaced = False)) Then
                                    metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + _
                                    "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox13.Text) + "</meta>" + Chr(10) + _
                                    "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Mid(metadatafile, startpos)
                                ElseIf ((creator2fileasplaced = False) And (creator2roleplaced = True)) Then
                                    metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox13.Text) + "</meta>" + Mid(metadatafile, startpos)
                                ElseIf ((creator2fileasplaced = True) And (creator2roleplaced = False)) Then
                                    startpos = InStr(startpos, metadatafile, "</meta>")
                                    metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Mid(metadatafile, startpos)
                                End If
                            End If
                        End If
                    Else
                        'Second creator being added
                        startpos = InStr(metadatafile, "</dc:creator>")
                        startpos = InStr(startpos, metadatafile, "<dc:") - 1 'start of next item after first creator
                        rolestring = "aut"
                        If ComboBox2.SelectedIndex = 1 Then rolestring = "edt"
                        If ComboBox2.SelectedIndex = 2 Then rolestring = "ill"
                        If ComboBox2.SelectedIndex = 3 Then rolestring = "trl"
                        metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator id=" + Chr(34) + "creator2" + Chr(34) + ">" + XMLOutput(TextBox3.Text) + "</dc:creator>" + Chr(13) + Chr(10) + _
                        "    <meta refines=" + Chr(34) + "#creator2" + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox13.Text) + "</meta>" + Chr(13) + Chr(10) + _
                        "    <meta refines=" + Chr(34) + "#creator2" + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Chr(13) + Chr(10) + "    " + Mid(metadatafile, startpos)
                    End If
                End If
                If TextBox3.Text = "" Then
                    endpos = InStr(metadatafile, "</dc:creator>") 'find end of first creator
                    startpos = InStr(endpos, metadatafile, "<dc:creator") 'look for another one
                    If startpos <> 0 Then
                        'Second creator exists but needs to be deleted
                        endpos = InStr(startpos, metadatafile, "</dc:creator>")

                        'Get id
                        idpos = InStr(startpos, metadatafile, "id=")
                        idinfo = ""
                        If idpos <> 0 Then
                            For temploop = idpos + 4 To endpos
                                If Mid(metadatafile, temploop, 1) = Chr(34) Then
                                    idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                                    temppos = InStr(startpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                                    If temppos = 0 Then temppos = InStr(startpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                                    While temppos <> 0
                                        endpos = InStr(temppos, metadatafile, "</meta>")
                                        If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                                        metadatafile = Mid(metadatafile, 1, temppos - 1) + Mid(metadatafile, endpos + 8)
                                        startpos = InStr(metadatafile, "</dc:creator>")
                                        startpos = InStr(startpos, metadatafile, "<dc:creator")
                                        temppos = InStr(startpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                                        If temppos = 0 Then temppos = InStr(startpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                                    End While
                                    startpos = InStr(metadatafile, "</dc:creator>")
                                    startpos = InStr(startpos, metadatafile, "<dc:creator")
                                    endpos = InStr(startpos, metadatafile, "</dc:creator>")
                                    metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endpos + 14)
                                    Exit For
                                End If
                            Next
                        Else
                            'No id (therefore no refines)
                            metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endpos + 14)
                        End If
                    End If
                End If
            Else
                'If optional attributes
                If TextBox12.Text <> "" Then
                    optionaltext = ""
                    If ((ComboBox1.SelectedIndex = 0) Or (ComboBox1.SelectedIndex = -1)) Then optionaltext = " opf:role=" + Chr(34) + "aut" + Chr(34)
                    If ComboBox1.SelectedIndex = 1 Then optionaltext = " opf:role=" + Chr(34) + "edt" + Chr(34)
                    If ComboBox1.SelectedIndex = 2 Then optionaltext = " opf:role=" + Chr(34) + "ill" + Chr(34)
                    If ComboBox1.SelectedIndex = 3 Then optionaltext = " opf:role=" + Chr(34) + "trl" + Chr(34)
                    optionaltext = " opf:file-as=" + Chr(34) + XMLOutput(TextBox12.Text) + Chr(34) + optionaltext + ">"
                Else
                    optionaltext = ">"
                End If
                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + optionaltext + XMLOutput(TextBox2.Text) + Mid(metadatafile, endpos)

                'Output second creator?
                endpos = InStr(metadatafile, "</dc:creator>") 'find end of first creator
                startpos = InStr(endpos, metadatafile, "<dc:creator") 'look for another one
                If ((TextBox3.Text <> "") Or (startpos <> 0)) Then
                    'Get optional attributes
                    If TextBox13.Text <> "" Then
                        optionaltext = ""
                        If ((ComboBox2.SelectedIndex = 0) Or (ComboBox2.SelectedIndex = -1)) Then optionaltext = " opf:role=" + Chr(34) + "aut" + Chr(34)
                        If ComboBox2.SelectedIndex = 1 Then optionaltext = " opf:role=" + Chr(34) + "edt" + Chr(34)
                        If ComboBox2.SelectedIndex = 2 Then optionaltext = " opf:role=" + Chr(34) + "ill" + Chr(34)
                        If ComboBox2.SelectedIndex = 3 Then optionaltext = " opf:role=" + Chr(34) + "trl" + Chr(34)
                        optionaltext = " opf:file-as=" + Chr(34) + XMLOutput(TextBox13.Text) + Chr(34) + optionaltext + ">"
                    Else
                        optionaltext = ">"
                    End If
                    If startpos <> 0 Then
                        endpos = InStr(startpos, metadatafile, "</dc:creator>")
                        lenheader = Len("<dc:creator")
                        metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + optionaltext + XMLOutput(TextBox3.Text) + Mid(metadatafile, endpos)
                    Else
                        'Original file did not have second creator
                        metadatafile = Mid(metadatafile, 1, endpos + 13) + Chr(13) + Chr(10) + Chr(9) + "<dc:creator" + optionaltext + XMLOutput(TextBox3.Text) + "</dc:creator>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 14)
                    End If
                End If
                If TextBox3.Text = "" Then
                    endpos = InStr(metadatafile, "</dc:creator>") 'find end of first creator
                    startpos = InStr(endpos, metadatafile, "<dc:creator") 'look for another one
                    If startpos <> 0 Then
                        'Second creator exists but needs to be deleted
                        endpos = InStr(startpos, metadatafile, "</dc:creator>")
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endpos + 14)
                    End If
                End If
            End If
        Else
            If versioninfo = "3.0" Then
                'Creator being added
                startpos = InStr(metadatafile, "</dc:title>")
                startpos = InStr(startpos, metadatafile, "<dc:") - 1 'start of next item after title
                If TextBox12.Text <> "" Then
                    rolestring = "aut"
                    If ComboBox1.SelectedIndex = 1 Then rolestring = "edt"
                    If ComboBox1.SelectedIndex = 2 Then rolestring = "ill"
                    If ComboBox1.SelectedIndex = 3 Then rolestring = "trl"
                    metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator id=" + Chr(34) + "creator" + Chr(34) + ">" + XMLOutput(TextBox2.Text) + "</dc:creator>" + Chr(13) + Chr(10) + _
                    "    <meta refines=" + Chr(34) + "#creator" + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox12.Text) + "</meta>" + Chr(13) + Chr(10) + _
                    "    <meta refines=" + Chr(34) + "#creator" + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Chr(13) + Chr(10) + "    " + Mid(metadatafile, startpos)
                Else
                    metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator>" + XMLOutput(TextBox3.Text) + "</dc:creator>" + Chr(13) + Chr(10) + "    " + Mid(metadatafile, startpos)
                End If

                'Check for second author
                If (TextBox3.Text <> "") Then
                    'Second creator being added
                    startpos = InStr(metadatafile, "</dc:creator>")
                    startpos = InStr(startpos, metadatafile, "<dc:") - 1 'start of next item after first creator
                    'Get optional attributes
                    If TextBox13.Text <> "" Then
                        rolestring = "aut"
                        If ComboBox2.SelectedIndex = 1 Then rolestring = "edt"
                        If ComboBox2.SelectedIndex = 2 Then rolestring = "ill"
                        If ComboBox2.SelectedIndex = 3 Then rolestring = "trl"
                        metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator id=" + Chr(34) + "creator2" + Chr(34) + ">" + XMLOutput(TextBox3.Text) + "</dc:creator>" + Chr(13) + Chr(10) + _
                        "    <meta refines=" + Chr(34) + "#creator2" + Chr(34) + " property=" + Chr(34) + "file-as" + Chr(34) + ">" + XMLOutput(TextBox13.Text) + "</meta>" + Chr(13) + Chr(10) + _
                        "    <meta refines=" + Chr(34) + "#creator2" + Chr(34) + " property=" + Chr(34) + "role" + Chr(34) + " scheme=" + Chr(34) + "marc:relators" + Chr(34) + ">" + rolestring + "</meta>" + Chr(13) + Chr(10) + "    " + Mid(metadatafile, startpos)
                    Else
                        metadatafile = Mid(metadatafile, 1, startpos) + "<dc:creator>" + XMLOutput(TextBox3.Text) + "</dc:creator>" + Chr(13) + Chr(10) + "    " + Mid(metadatafile, startpos)
                    End If
                End If
            Else
                'No creator yet, so add it after <metadata... > tag
                startpos = InStr(metadatafile, "<metadata")
                startpos = InStr(startpos, metadatafile, ">") + 1
                'If optional attributes
                If TextBox12.Text <> "" Then
                    optionaltext = ""
                    If ((ComboBox1.SelectedIndex = 0) Or (ComboBox1.SelectedIndex = -1)) Then optionaltext = " opf:role=" + Chr(34) + "aut" + Chr(34)
                    If ComboBox1.SelectedIndex = 1 Then optionaltext = " opf:role=" + Chr(34) + "edt" + Chr(34)
                    If ComboBox1.SelectedIndex = 2 Then optionaltext = " opf:role=" + Chr(34) + "ill" + Chr(34)
                    If ComboBox1.SelectedIndex = 3 Then optionaltext = " opf:role=" + Chr(34) + "trl" + Chr(34)
                    optionaltext = " opf:file-as=" + Chr(34) + XMLOutput(TextBox12.Text) + Chr(34) + optionaltext + ">"
                Else
                    optionaltext = ">"
                End If

                ' check for second author
                If (TextBox3.Text <> "") Then
                    'Get optional attributes
                    If TextBox13.Text <> "" Then
                        optionaltext2 = ""
                        If ((ComboBox2.SelectedIndex = 0) Or (ComboBox2.SelectedIndex = -1)) Then optionaltext2 = " opf:role=" + Chr(34) + "aut" + Chr(34)
                        If ComboBox2.SelectedIndex = 1 Then optionaltext2 = " opf:role=" + Chr(34) + "edt" + Chr(34)
                        If ComboBox2.SelectedIndex = 2 Then optionaltext2 = " opf:role=" + Chr(34) + "ill" + Chr(34)
                        If ComboBox2.SelectedIndex = 3 Then optionaltext2 = " opf:role=" + Chr(34) + "trl" + Chr(34)
                        optionaltext2 = " opf:file-as=" + Chr(34) + XMLOutput(TextBox13.Text) + Chr(34) + optionaltext2 + ">"
                    Else
                        optionaltext2 = ">"
                    End If
                    ' output two creators
                    metadatafile = Mid(metadatafile, 1, startpos) + "  <dc:creator" + optionaltext + XMLOutput(TextBox2.Text) + "</dc:creator>" + Chr(13) + Chr(10) + "  <dc:creator" + optionaltext2 + XMLOutput(TextBox3.Text) + "</dc:creator>" + Mid(metadatafile, startpos)
                Else
                    ' output only one creator
                    metadatafile = Mid(metadatafile, 1, startpos) + "  <dc:creator" + optionaltext + XMLOutput(TextBox2.Text) + "</dc:creator>" + Mid(metadatafile, startpos)
                End If
            End If
        End If

        'Output (Calibre) series and series index
        startpos = InStr(metadatafile, "calibre:series" + Chr(34))
        If ((TextBox15.Text <> "") Or (startpos <> 0)) Then
            If startpos <> 0 Then
                opfmeta = False
                temppos = startpos
                startpos = InStrRev(metadatafile, "<meta", startpos)
                If startpos = 0 Then
                    startpos = InStrRev(metadatafile, "<opf:meta", temppos)
                    opfmeta = True
                End If
                endpos = InStr(startpos, metadatafile, "/>")
                If opfmeta Then
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + "<opf:meta name=" & Chr(34) & "calibre:series" & Chr(34) + " content=" + Chr(34) + XMLOutput(TextBox15.Text) + Chr(34) + Mid(metadatafile, endpos)
                Else
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + "<meta name=" & Chr(34) & "calibre:series" & Chr(34) + " content=" + Chr(34) + XMLOutput(TextBox15.Text) + Chr(34) + Mid(metadatafile, endpos)
                End If
            Else
                endpos = InStr(metadatafile, "</dc:title>")
                metadatafile = Mid(metadatafile, 1, endpos + 10) + Chr(13) + Chr(10) + Chr(9) + "<meta name=" & Chr(34) & "calibre:series" & Chr(34) & " content=" + Chr(34) + XMLOutput(TextBox15.Text) + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 11)
            End If
        End If
        startpos = InStr(metadatafile, "calibre:series_index" + Chr(34))
        If ((TextBox14.Text <> "") Or (startpos <> 0)) Then
            If startpos <> 0 Then
                startpos = InStrRev(metadatafile, "<meta", startpos)
                endpos = InStr(startpos, metadatafile, "/>")
                metadatafile = Mid(metadatafile, 1, startpos - 1) + "<meta name=" & Chr(34) & "calibre:series_index" & Chr(34) + " content=" + Chr(34) + TextBox14.Text + Chr(34) + Mid(metadatafile, endpos)
            Else
                endpos = InStr(metadatafile, "</dc:title>")
                metadatafile = Mid(metadatafile, 1, endpos + 10) + Chr(13) + Chr(10) + Chr(9) + "<meta name=" & Chr(34) & "calibre:series_index" & Chr(34) & " content=" + Chr(34) + TextBox14.Text + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 11)
            End If
        End If

        'Output description
        metadatafile = metadatafile.Replace("<dc:description xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:description />")
        testpos = InStr(metadatafile, "<dc:description />")
        If ((testpos <> 0) And (TextBox4.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:description/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:description />", "<dc:description>" + XMLOutput(TextBox4.Text) + "</dc:description>")
                Else
                    startpos = InStr(metadatafile, "<dc:description")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<description")
                    If ((TextBox4.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:description>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</description>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + XMLOutput(TextBox4.Text) + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + XMLOutput(TextBox4.Text) + "</dc:description>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:description>" + XMLOutput(TextBox4.Text) + "</dc:description>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:description/>", "<dc:description>" + XMLOutput(TextBox4.Text) + "</dc:description>")
            End If
        End If

        'Output publisher
        metadatafile = metadatafile.Replace("<dc:publisher xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:publisher />")
        testpos = InStr(metadatafile, "<dc:publisher />")
        If ((testpos <> 0) And (TextBox5.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:publisher/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:publisher />", "<dc:publisher>" + XMLOutput(TextBox5.Text) + "</dc:publisher>")
                Else
                    startpos = InStr(metadatafile, "<dc:publisher")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<publisher")
                    If ((TextBox5.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:publisher>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</publisher>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + XMLOutput(TextBox5.Text) + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + XMLOutput(TextBox5.Text) + "</dc:publisher>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:publisher>" + XMLOutput(TextBox5.Text) + "</dc:publisher>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:publisher/>", "<dc:publisher>" + XMLOutput(TextBox5.Text) + "</dc:publisher>")
            End If
        End If

        'Output date
        startpos = InStr(metadatafile, "<dc:date")
        If startpos = 0 Then startpos = InStr(metadatafile, "<date")
        If ((TextBox6.Text <> "") Or (startpos <> 0)) Then
            If startpos <> 0 Then
                endtag = InStr(startpos, metadatafile, ">")
                checktag = InStr(startpos, metadatafile, "opf:event")
                newheader = ""
                If ((checktag <> 0) And (checktag < endtag)) Then
                    If Label6.Text = "Date" Then
                        ' remove event
                        newheader = "<dc:date>"
                    Else
                        ' replace existing event
                        newheader = "<dc:date opf:event=" + Chr(34) + Mid(Label6.Text, 7, Len(Label6.Text) - 7) + Chr(34) + ">"
                    End If
                Else
                    ' there is no event in the tag
                    If Label6.Text <> "Date" Then
                        ' add event
                        newheader = "<dc:date opf:event=" + Chr(34) + Mid(Label6.Text, 7, Len(Label6.Text) - 7) + Chr(34) + ">"
                    Else
                        ' leave things as they are
                        newheader = Mid(metadatafile, startpos, endtag - startpos + 1)
                    End If
                End If
                endpos = InStr(metadatafile, "</dc:date>")
                If endpos = 0 Then
                    endpos = InStr(metadatafile, "</date>")
                    If endpos <> 0 Then
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox6.Text + "</dc:date>" + Mid(metadatafile, endpos + 7)
                    Else
                        endpos = InStr(startpos, metadatafile, " />")
                        If endpos <> 0 Then
                            metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox6.Text + "</dc:date>" + Mid(metadatafile, endpos + 3)
                        End If
                    End If
                Else
                    metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox6.Text + Mid(metadatafile, endpos)
                End If
            Else
                endpos = InStr(metadatafile, "</dc:title>")
                If Label6.Text = "Date" Then
                    metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:date>" + TextBox6.Text + "</dc:date>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                Else
                    metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:date opf:event=" + Chr(34) + Mid(Label6.Text, 7, Len(Label6.Text) - 7) + Chr(34) + ">" + TextBox6.Text + "</dc:date>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                End If
            End If
        End If

        'Output subject
        metadatafile = metadatafile.Replace("<dc:subject xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:subject />")
        temptext = XMLOutput(TextBox17.Text)
        ' delete any subjects present
        startpos = InStr(metadatafile, "<dc:subject>")
        If startpos = 0 Then startpos = InStr(metadatafile, "<subject>")
        While (startpos <> 0)
            endtag = InStr(startpos + 1, metadatafile, "<")
            endtag = InStr(endtag + 1, metadatafile, "<")
            metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endtag)
            startpos = InStr(metadatafile, "<dc:subject>")
            If startpos = 0 Then startpos = InStr(metadatafile, "<subject>")
        End While
        If TextBox17.Text <> "" Then
            If TextBox17.Text.Contains(subjectseparator) Then
                ' preformat TextBox17.text
                temptext = temptext.Replace(subjectseparator, "</dc:subject>" + Chr(13) + Chr(10) + Chr(9) + Chr(9) + "<dc:subject>")
            End If
        End If
        testpos = InStr(metadatafile, "<dc:subject />")
        If ((testpos <> 0) And (TextBox17.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:subject/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:subject />", "<dc:subject>" + temptext + "</dc:subject>")
                Else
                    endpos = InStr(metadatafile, "</dc:title>")
                    metadatafile = Mid(metadatafile, 1, endpos + 10) + Chr(9) + Chr(9) + "<dc:subject>" + temptext + "</dc:subject>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 11)
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:subject/>", "<dc:subject>" + temptext + "</dc:subject>")
            End If
            metadatafile = metadatafile.Replace("<dc:subject></dc:subject>", "<dc:subject/>")
        End If

        'Output type
        metadatafile = metadatafile.Replace("<dc:type xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:type />")
        testpos = InStr(metadatafile, "<dc:type />")
        If ((testpos <> 0) And (TextBox7.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:type/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:type />", "<dc:type>" + TextBox7.Text + "</dc:type>")
                Else
                    startpos = InStr(metadatafile, "<dc:type")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<type")
                    If ((TextBox7.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:type>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</type>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox7.Text + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox7.Text + "</dc:type>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:type>" + TextBox7.Text + "</dc:type>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:type/>", "<dc:type>" + TextBox7.Text + "</dc:type>")
            End If
        End If

        'Output format
        metadatafile = metadatafile.Replace("<dc:format xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:format />")
        testpos = InStr(metadatafile, "<dc:format />")
        If ((testpos <> 0) And (TextBox8.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:format/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:format />", "<dc:format>" + TextBox8.Text + "</dc:format>")
                Else
                    startpos = InStr(metadatafile, "<dc:format")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<format")
                    If ((TextBox8.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:format>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</format>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox8.Text + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox8.Text + "</dc:format>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:format>" + TextBox8.Text + "</dc:format>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:format/>", "<dc:format>" + TextBox8.Text + "</dc:format>")
            End If
        End If

        'Output identifier
        If versioninfo = "3.0" Then
            'before looking for first identifier, look for scheme="uuid"
            startpos = InStr(metadatafile, "scheme=" + Chr(34) + "uuid")
            If startpos <> 0 Then
                'Scan backwards to <dc:identifier
                startpos = InStrRev(metadatafile, "<dc:identifier", startpos)
            Else
                'Find the first <dc:identifier
                startpos = InStr(metadatafile, "<dc:identifier")
            End If
            If startpos <> 0 Then
                endheaderpos = InStr(startpos, metadatafile, ">")
                endpos = InStr(startpos, metadatafile, "</dc:identifier>")
                If endpos = 0 Then
                    endpos = InStr(startpos, metadatafile, " />")
                    extrachars = 3
                    If endpos = 0 Then
                        InStr(startpos, metadatafile, "/>")
                        extrachars = 2
                    End If
                    If endpos <> 0 Then
                        metadatafile = Mid(metadatafile, 1, endpos + 2 - extrachars) + "</dc:identifier>" + Mid(metadatafile, endpos + extrachars)
                    Else
                        DialogResult = MsgBox("Badly formed OPF file.  EPUB Metadata Editor is unable to save this file.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                        Exit Sub
                    End If
                End If
                lenheader = Len("<dc:identifier")
                metadatafile = Mid(metadatafile, 1, endheaderpos) + TextBox9.Text + Mid(metadatafile, endpos)

                If Label9.Text <> "Identifier" Then
                    identifierscheme = Mid(Label9.Text, 13, Len(Label9.Text) - 13)
                Else
                    GoTo outputsource
                End If

                ' get id
                idpos = InStr(startpos, metadatafile, "id=")
                idinfo = ""
                If idpos <> 0 Then
                    For temploop = idpos + 4 To endpos
                        If Mid(metadatafile, temploop, 1) = Chr(34) Then
                            idinfo = Mid(metadatafile, idpos + 4, temploop - idpos - 4)
                            GoTo lookforrefines5
                        End If
                    Next
                Else
                    metadatafile = Mid(metadatafile, 1, startpos) + "<dc:identifier id=" + Chr(34) + "pub-id" + Chr(34) + ">" + TextBox9.Text + Mid(metadatafile, endpos)
                    idinfo = "pub-id"
                End If
lookforrefines5:
                If idinfo <> "" Then
                    temppos = InStr(startpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                    If temppos = 0 Then temppos = InStr(startpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                    schemeplaced = False
                    While temppos <> 0
                        endheaderpos = InStr(temppos, metadatafile, ">")
                        endpos = InStr(temppos, metadatafile, "</meta>")
                        If endpos = 0 Then endpos = InStr(temppos, metadatafile, "</opf:meta>")
                        refinespos = InStr(temppos, metadatafile, "scheme=" + Chr(34))
                        If refinespos <> 0 Then
                            If refinespos < endpos Then
                                metadatafile = Mid(metadatafile, 1, refinespos + 7) + identifierscheme.Replace("=", Chr(34) + ">") + Mid(metadatafile, endpos)
                                schemeplaced = True
                            End If
                        End If
                        temppos = InStr(endpos, metadatafile, "<meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                        If temppos = 0 Then temppos = InStr(endpos, metadatafile, "<opf:meta refines=" + Chr(34) + "#" + idinfo + Chr(34))
                    End While
                    If (schemeplaced = False) Then
                        startpos = InStr(metadatafile, "</dc:identifier>") + 16 'end of identifier
                        metadatafile = Mid(metadatafile, 1, startpos) + Chr(13) + Chr(10) + "    <meta refines=" + Chr(34) + "#" + idinfo + Chr(34) + " property=" + Chr(34) + "identifier-type" + Chr(34) + " scheme=" + Chr(34) + identifierscheme.Replace("=", Chr(34) + ">") + "</meta>" + Mid(metadatafile, startpos)
                    End If
                End If
            End If
        Else
            'before looking for first identifier, look for scheme="uuid"
            startpos = InStr(metadatafile, "opf:scheme=" + Chr(34) + "uuid")
            If startpos <> 0 Then
                'Scan backwards to <dc:identifier
                startpos = InStrRev(metadatafile, "<dc:identifier", startpos)
            Else
                'Find the first <dc:identifier
                startpos = InStr(metadatafile, "<dc:identifier")
            End If
            If startpos = 0 Then startpos = InStr(metadatafile, "<identifier")
            If ((TextBox9.Text <> "") Or (startpos <> 0)) Then
                If startpos <> 0 Then
                    endtag = InStr(startpos, metadatafile, ">")
                    checktag = InStr(startpos, metadatafile, "opf:scheme")
                    lookforID = InStr(startpos, metadatafile, "id=")
                    If ((lookforID <> 0) And (lookforID < endtag)) Then
                        endID = InStr(lookforID + 5, metadatafile, Chr(34))
                        ID = Mid(metadatafile, lookforID + 4, endID - lookforID - 4)
                        If ID = "uuid_id" Then founduuid = True Else founduuid = False
                    Else
                        ID = ""
                        lookforID = 0
                    End If
                    newheader = ""
                    If ((checktag <> 0) And (checktag < endtag)) Then
                        If Label9.Text = "Identifier" Then
                            ' remove scheme
                            If lookforID = 0 Then
                                newheader = "<dc:identifier>"
                            Else
                                newheader = "<dc:identifier id=" + Chr(34) + ID + Chr(34) + ">"
                            End If
                        Else
                            ' replace existing scheme
                            If lookforID = 0 Then
                                If Mid(Label9.Text, 13, Len(Label9.Text) - 13) = "uuid" Then
                                    newheader = "<dc:identifier id=" + Chr(34) + "uuid_id" + Chr(34) + " opf:scheme=" + Chr(34) + "uuid" + Chr(34) + ">"
                                Else
                                    newheader = "<dc:identifier opf:scheme=" + Chr(34) + Mid(Label9.Text, 13, Len(Label9.Text) - 13) + Chr(34) + ">"
                                End If
                            Else
                                If Mid(Label9.Text, 13, Len(Label9.Text) - 13) = "uuid" Then
                                    newheader = "<dc:identifier id=" + Chr(34) + "uuid_id" + Chr(34) + " opf:scheme=" + Chr(34) + "uuid" + Chr(34) + ">"
                                Else
                                    newheader = "<dc:identifier id=" + Chr(34) + ID + Chr(34) + " opf:scheme=" + Chr(34) + Mid(Label9.Text, 13, Len(Label9.Text) - 13) + Chr(34) + ">"
                                End If
                            End If
                        End If
                    Else
                        ' there is no scheme in the tag
                        If Label9.Text <> "Identifier" Then
                            ' add scheme
                            If lookforID = 0 Then
                                If Mid(Label9.Text, 13, Len(Label9.Text) - 13) = "uuid" Then
                                    newheader = "<dc:identifier id=" + Chr(34) + "uuid_id" + Chr(34) + " opf:scheme=" + Chr(34) + "uuid" + Chr(34) + ">"
                                Else
                                    newheader = "<dc:identifier opf:scheme=" + Chr(34) + Mid(Label9.Text, 13, Len(Label9.Text) - 13) + Chr(34) + ">"
                                End If
                            Else
                                If Mid(Label9.Text, 13, Len(Label9.Text) - 13) = "uuid" Then
                                    newheader = "<dc:identifier id=" + Chr(34) + "uuid_id" + Chr(34) + " opf:scheme=" + Chr(34) + "uuid" + Chr(34) + ">"
                                Else
                                    newheader = "<dc:identifier id=" + Chr(34) + ID + Chr(34) + " opf:scheme=" + Chr(34) + Mid(Label9.Text, 13, Len(Label9.Text) - 13) + Chr(34) + ">"
                                End If
                            End If
                        Else
                            ' leave things as they are
                            If lookforID = 0 Then
                                newheader = "<dc:identifier>"
                            Else
                                newheader = "<dc:identifier id=" + Chr(34) + ID + Chr(34) + ">"
                            End If
                        End If
                    End If
                    endpos = InStr(startpos, metadatafile, "</dc:identifier>")
                    extracheck = InStr(startpos + 1, metadatafile, "<dc:")
                    If ((extracheck <> 0) And (endpos > extracheck)) Then endpos = 0 'look to see if field end is actually for a second identifier
                    If endpos = 0 Then
                        endpos = InStr(metadatafile, "</identifier>")
                        If endpos <> 0 Then
                            metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox9.Text + "</dc:identifier>" + Mid(metadatafile, endpos + 13)
                        Else
                            endpos = InStr(startpos, metadatafile, " />")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox9.Text + "</dc:identifier>" + Mid(metadatafile, endpos + 3)
                            End If
                        End If
                    Else
                        metadatafile = Mid(metadatafile, 1, startpos - 1) + newheader + TextBox9.Text + Mid(metadatafile, endpos)
                    End If

                    'If ID is uuid, then we also need to change toc.ncx
                    If founduuid = True Then
                        tocfiletext = LoadUnicodeFile(tocfile)
                        tocstartpos = InStr(tocfiletext, "name=" + Chr(34) + "dtb:uid")
                        If tocstartpos <> 0 Then
                            'scan back to start of line
                            tocstartpos = InStrRev(tocfiletext, "<meta", tocstartpos)
                            'scan forward to content=
                            toccontentloc = InStr(tocstartpos, tocfiletext, "content=")
                            toccontentloc = InStr(toccontentloc, tocfiletext, Chr(34))
                            toccontentend = InStr(toccontentloc + 1, tocfiletext, Chr(34))
                            tocfiletext = Mid(tocfiletext, 1, toccontentloc) + TextBox9.Text + Mid(tocfiletext, toccontentend)
                            SaveUnicodeFile(tocfile, tocfiletext)
                        End If
                    End If
                Else
                    endpos = InStr(metadatafile, "</dc:title>")
                    If Label9.Text = "Identifier" Then
                        metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:identifier>" + TextBox9.Text + "</dc:identifier>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                    Else
                        metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:identifier opf:scheme=" + Chr(34) + Mid(Label9.Text, 13, Len(Label9.Text) - 13) + Chr(34) + ">" + TextBox9.Text + "</dc:identifier>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                    End If
                End If
            End If

        End If

outputsource:
        'Output source
        metadatafile = metadatafile.Replace("<dc:source xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:source />")
        testpos = InStr(metadatafile, "<dc:source />")
        If ((testpos <> 0) And (TextBox10.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:source/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:source />", "<dc:source>" + TextBox10.Text + "</dc:source>")
                Else
                    startpos = InStr(metadatafile, "<dc:source")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<source")
                    If ((TextBox10.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:source>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</source>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox10.Text + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox10.Text + "</dc:source>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:source>" + TextBox10.Text + "</dc:source>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:source/>", "<dc:source>" + TextBox10.Text + "</dc:source>")
            End If
        End If

        'Output language
        metadatafile = metadatafile.Replace("<dc:language xmlns:dc=" + Chr(34) + "http://purl.org/dc/elements/1.1/" + Chr(34) + " />", "<dc:language />")
        testpos = InStr(metadatafile, "<dc:language />")
        If ((testpos <> 0) And (TextBox11.Text = "")) Then
        Else
            startpos = InStr(metadatafile, "<dc:language/>")
            If startpos = 0 Then
                If testpos <> 0 Then
                    metadatafile = metadatafile.Replace("<dc:language />", "<dc:language>" + TextBox11.Text + "</dc:language>")
                Else
                    startpos = InStr(metadatafile, "<dc:language")
                    If startpos = 0 Then startpos = InStr(metadatafile, "<language")
                    If ((TextBox11.Text <> "") Or (startpos <> 0)) Then
                        If startpos <> 0 Then
                            endtag = InStr(startpos, metadatafile, ">")
                            lenheader = endtag - startpos + 1
                            endpos = InStr(metadatafile, "</dc:language>")
                            If endpos = 0 Then endpos = InStr(metadatafile, "</language>")
                            If endpos <> 0 Then
                                metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox11.Text + Mid(metadatafile, endpos)
                            Else
                                endpos = InStr(startpos, metadatafile, " />")
                                If endpos <> 0 Then
                                    metadatafile = Mid(metadatafile, 1, startpos + lenheader - 1) + TextBox11.Text + "</dc:language>" + Mid(metadatafile, endpos + 3)
                                End If
                            End If
                        Else
                            endpos = InStr(metadatafile, "</dc:title>")
                            metadatafile = Mid(metadatafile, 1, endpos + 11) + Chr(13) + Chr(10) + Chr(9) + "<dc:language>" + TextBox11.Text + "</dc:language>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos + 12)
                        End If
                    End If
                End If
            Else
                metadatafile = metadatafile.Replace("<dc:language/>", "<dc:language>" + TextBox11.Text + "</dc:language>")
            End If
        End If

        'Regularise whitespace
        metadatafile = Regularise(metadatafile)

        RichTextBox1.Text = metadatafile
        SaveUnicodeFile(opffile, metadatafile)

        If Not SaveOPFOnly Then
            Dim tempEpubFileName As String
            Dim zip As ZipStorer = Nothing

            'Create temporary file (in case zip fails)
            tempEpubFileName = Mid(EpubFileName, 1, InStrRev(EpubFileName, ".") - 1) + "temp" + Mid(EpubFileName, InStrRev(EpubFileName, "."))

            'Zip temp directory to temp file
            temporarydirectory = CurDir()
            ChDir(ebookdirectory)

            'Delete mimetype file
            IO.File.Delete("mimetype")
            ChDir(temporarydirectory)
            Try
                zip = ZipStorer.Create(tempEpubFileName, "")
                Dim mimetype As New MemoryStream(System.Text.Encoding.UTF8.GetBytes("application/epub+zip"))
                zip.AddStream(ZipStorer.Compression.Store, "mimetype", mimetype, DateTime.Now, "")
                mimetype.Close()
                Dim dir = Directory.GetDirectories(ebookdirectory)
                Dim item As String
                For Each item In dir
                    zip.AddDirectory(ZipStorer.Compression.Deflate, item, "", "")
                Next
                Dim files = Directory.GetFiles(ebookdirectory)
                For Each item In files
                    zip.AddFile(ZipStorer.Compression.Deflate, item, Path.GetFileName(item), "")
                Next
                zip.Close()
                fi.Delete()
                wait(500)
                IO.File.Copy(tempEpubFileName, EpubFileName)
                wait(500)
                IO.File.Delete(tempEpubFileName)

                'Update interface
                Me.Text = CaptionString
                Button3.Enabled = False
                projectchanged = False
            Catch ex As Exception
                DialogResult = MsgBox("The EPUB failed to save properly.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                zip.Close()
                wait(500)
                IO.File.Delete(tempEpubFileName)
            End Try
        End If
    End Sub

    Public Function Regularise(ByVal metadatatext As String) As String
        'Regularise whitespace
        ' delete stuff
        metadatatext = metadatatext.Replace(Chr(13), "")
        metadatatext = metadatatext.Replace(Chr(10), "")
        metadatatext = metadatatext.Replace(Chr(9), "")
        While (metadatatext.Contains("> "))
            metadatatext = metadatatext.Replace("> ", ">")
        End While
        While (metadatatext.Contains("< "))
            metadatatext = metadatatext.Replace("< ", "<")
        End While

        ' add stuff back
        metadatatext = metadatatext.Replace("><", ">" + Chr(13) + Chr(10) + "<")
        metadatatext = metadatatext.Replace("<metadata", "  <metadata")
        metadatatext = metadatatext.Replace("</metadata", "  </metadata")
        metadatatext = metadatatext.Replace("<manifest", "  <manifest")
        metadatatext = metadatatext.Replace("</manifest", "  </manifest")
        metadatatext = metadatatext.Replace("<spine", "  <spine")
        metadatatext = metadatatext.Replace("</spine", "  </spine")
        metadatatext = metadatatext.Replace("<guide", "  <guide")
        metadatatext = metadatatext.Replace("</guide", "  </guide")
        metadatatext = metadatatext.Replace("<dc:", "    <dc:")
        metadatatext = metadatatext.Replace("<meta ", "    <meta ")
        metadatatext = metadatatext.Replace("<opf:meta ", "    <opf:meta ")
        metadatatext = metadatatext.Replace("<item", "    <item")
        metadatatext = metadatatext.Replace("<reference", "    <reference")
        metadatatext = metadatatext.Replace("<!--", "    <!--")
        metadatatext = metadatatext.Replace(Chr(10) + "</dc:", Chr(10) + "    </dc:")
        Return metadatatext
    End Function
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim metadatafile As String

        OpenFileDialog5.Filter = "All files (*.*)|*.*"
        OpenFileDialog5.FilterIndex = 1
        OpenFileDialog5.FileName = ""
        OpenFileDialog5.InitialDirectory = ebookdirectory
        If OpenFileDialog5.ShowDialog = Windows.Forms.DialogResult.OK Then
            fileeditorreturn = False
            filecontents = ""
            Dim fileextension = System.IO.Path.GetExtension(OpenFileDialog5.FileName).ToLower
            If ((fileextension = ".xhtml") Or (fileextension = ".html")) Then
                Form2.Button9.Visible = True
            Else
                Form2.Button9.Visible = False
            End If
            If (fileextension = ".opf") Then
                Form2.Button3.Visible = True
            Else
                Form2.Button3.Visible = False
            End If
            Form2.RichTextBox1.Text = LoadUnicodeFile(OpenFileDialog5.FileName)
            Form2.ShowDialog()
            If fileeditorreturn = True Then
                RichTextBox1.Text = filecontents
                SaveUnicodeFile(OpenFileDialog5.FileName, RichTextBox1.Text)
                projectchanged = True
                Button3.Enabled = True
                Me.Text = "*" + CaptionString

                ' Possibly need to update metadata
                keepcombobox = True
                ClearInterface()
                keepcombobox = False
                RichTextBox1.Text = LoadUnicodeFile(opffile)
                metadatafile = LoadUnicodeFile(opffile)
                ExtractMetadata(metadatafile, True)
                refreshfilelist = False
            End If
        End If
    End Sub

    Private Sub Button27_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button27.Click
        ' Copy cover to "0000Cover" in the root directory so that CDisplayEx shell extension can be used to display cover
        ' See: http://wiki.mobileread.com/wiki/Thumbnails and http://www.mobileread.com/forums/showthread.php?t=35505

        Dim currentcoverimageextension, newcoverimagefilename, newlineandspace, metadatafile, insertion, rest As String
        Dim startpos, insertpos, endpos, count, x As Integer

        RichTextBox1.Text = LoadUnicodeFile(opffile)
        metadatafile = LoadUnicodeFile(opffile)

        ' If 0000Cover file exists, then delete it
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(ebookdirectory, FileIO.SearchOption.SearchTopLevelOnly, "0000Cover.*")
            My.Computer.FileSystem.DeleteFile(foundFile)

            ' Delete the item from the opf file
            insertpos = InStr(metadatafile, "id=" + Chr(34) + "prioritorised_cover" + Chr(34))
            startpos = InStrRev(metadatafile, "<", insertpos)
            endpos = InStr(insertpos, metadatafile, ">")
            metadatafile = Mid(metadatafile, 1, startpos - 1) + Mid(metadatafile, endpos + 1)
        Next

        currentcoverimageextension = Path.GetExtension(coverimagefile)
        newcoverimagefilename = "0000Cover" + currentcoverimageextension

        If CheckBox5.Checked = False Then
            ' Copy cover file to root directory and give prioritised name
            My.Computer.FileSystem.CopyFile(coverimagefile, ebookdirectory + "\" + newcoverimagefilename)
        Else
            Try
                ' Resize cover file
                ' Code adapted from http://www.thedesilva.com/2010/01/resize-image-using-vb-net/
                Dim bm As New Bitmap(coverimagefile, False)
                Dim height As Integer = 256
                Dim percentResize As Decimal = height / bm.Height
                Dim width As Integer = bm.Width * percentResize 'calculate width maintaining aspect ratio
                Dim thumb As New Bitmap(width, height)
                Dim g As Graphics = Graphics.FromImage(thumb)
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)
                g.Dispose()
                bm.Dispose()
                thumb.Save(ebookdirectory + "\" + newcoverimagefilename)
                thumb.Dispose()
            Catch ex As Exception
                ' Copy cover file to root directory and give prioritised name
                My.Computer.FileSystem.CopyFile(coverimagefile, ebookdirectory + "\" + newcoverimagefilename)
            End Try
        End If

        If ebookdirectory <> opfdirectory Then
            startpos = InStr(opfdirectory, ebookdirectory) + Len(ebookdirectory)
            rest = Mid(opfdirectory, startpos)
            count = Len(rest.ToString) - Len(rest.ToString.Replace("\", ""))
            insertion = ""
            For x = 1 To count
                insertion = insertion + "../"
            Next
            newcoverimagefilename = insertion + newcoverimagefilename
        End If

        ' Add file to opf
        newlineandspace = Chr(10)
        insertion = ""
        startpos = InStr(metadatafile, "</manifest")
        If startpos <> 0 Then
            insertpos = InStr(startpos + 1, metadatafile, "<")
            If insertpos <> 0 Then
                newlineandspace = Mid(metadatafile, startpos + 10, insertpos - startpos - 10)
                insertion = newlineandspace + "<item href=" + Chr(34) + newcoverimagefilename + Chr(34) + " id=" + Chr(34) + "prioritorised_cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>" + newlineandspace
                metadatafile = Mid(metadatafile, 1, startpos + 9) + insertion + Mid(metadatafile, insertpos, Len(metadatafile) - insertpos + 1)
            End If
            ' alternate code to see if location in opf file caused issue with Microsoft Edge
            'insertion = "<item href=" + Chr(34) + newcoverimagefilename + Chr(34) + " id=" + Chr(34) + "prioritorised_cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>"
            'metadatafile = Mid(metadatafile, 1, startpos - 1) + insertion + Mid(metadatafile, startpos)
        End If

        'save opf file
        SaveUnicodeFile(opffile, metadatafile)

        'update interface
        Button27.Visible = False
        Label23.Visible = False
        If ((Button1.Visible = False) And (Button35.Visible = False)) Then
            Button42.Visible = False
        End If
        CheckBox5.Visible = False
        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim oldeventtype, neweventtype As String
        Dialog2.Text = "Edit Date's event type"
        If Label6.Text = "Date" Then
            oldeventtype = ""
        Else
            oldeventtype = Mid(Label6.Text, 7, Len(Label6.Text) - 7)
        End If
        Dialog2.TextBox1.Text = oldeventtype
        Dialog2.ShowDialog()
        neweventtype = Dialog2.TextBox1.Text
        If ((Dialog2.DialogResult = Windows.Forms.DialogResult.OK) And (oldeventtype <> neweventtype)) Then
            If neweventtype = "" Then
                Label6.Text = "Date"
                TextBox6.Width = 304
                TextBox6.Left = 81
            Else
                Label6.Text = "Date (" + neweventtype + ")"
                TextBox6.Width = 255
                TextBox6.Left = 130
            End If
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim oldschemetype, newschemetype As String
        Dialog2.Text = "Edit Identifier's scheme type:"
        If Label9.Text = "Identifier" Then
            oldschemetype = ""
        Else
            oldschemetype = Mid(Label9.Text, 13, Len(Label9.Text) - 13)
        End If
        Dialog2.TextBox1.Text = oldschemetype
        Dialog2.ShowDialog()
        newschemetype = Dialog2.TextBox1.Text
        If ((Dialog2.DialogResult = Windows.Forms.DialogResult.OK) And (oldschemetype <> newschemetype)) Then
            If newschemetype = "" Then
                Label9.Text = "Identifier"
                TextBox9.Width = 304
                TextBox9.Left = 81
            Else
                If versioninfo = "3.0" Then
                    If InStr(newschemetype, "=") = 0 Then
                        newschemetype = "xsd:string=" + newschemetype
                    End If
                End If
                Label9.Text = "Identifier (" + newschemetype + ")"
                TextBox9.Width = 255
                TextBox9.Left = 130
            End If
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        Dim tempstring As String
        tempstring = TextBox1.Text
        TextBox1.Text = TextBox2.Text
        TextBox2.Text = tempstring
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Dim tempstring As String
        Dim item As Integer
        tempstring = TextBox2.Text
        TextBox2.Text = TextBox3.Text
        TextBox3.Text = tempstring
        tempstring = TextBox12.Text
        TextBox12.Text = TextBox13.Text
        TextBox13.Text = tempstring
        item = ComboBox1.SelectedIndex
        If (TextBox2.Text <> "") Then
            ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        Else
            ComboBox1.SelectedIndex = -1
        End If
        If (TextBox3.Text <> "") Then
            ComboBox2.SelectedIndex = item
        Else
            ComboBox2.SelectedIndex = -1
        End If
    End Sub

    Private Function CheckForUpdate(ByVal background As Boolean) As String
        Dim versioninfo, currentversion, returnstring As String
        Dim str As String = ""
        Dim currpos As Integer
        Dim client As WebClient = New WebClient()
        ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)

        returnstring = ""

        Try
            Dim data As Stream = client.OpenRead("https://github.com/benchen71/epub-metadata-editor")
            Dim reader As StreamReader = New StreamReader(data)

            str = reader.ReadLine()
            Do While Not str Is Nothing
                If InStr(str, "Version: ") Then
                    GoTo foundversion
                End If
                str = reader.ReadLine()
            Loop
            If Not background Then
                DialogResult = MsgBox("Check for version failed: web page missing latest version info.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            End If
            Return returnstring

foundversion:
            currpos = InStr(str, "Version: ")
            versioninfo = Mid(str, currpos + 9)
            currentversion = Mid(My.Application.Info.Version.ToString, 1, Len(My.Application.Info.Version.ToString) - 2)
            Dim oldVersion As New Version(currentversion)
            Dim newVersion As New Version(versioninfo)
            If Version.op_GreaterThan(newVersion, oldVersion) Then
                Dialog3.Label4.Text = "Current installed version: " + currentversion + Chr(10) + "Latest available version: " + versioninfo
                If Not background Then
                    DialogResult = Dialog3.ShowDialog
                Else
                    returnstring = "Update available!"
                    updateinfo = "Current installed version: " + currentversion + Chr(10) + "Latest available version: " + versioninfo
                End If
            ElseIf Version.Equals(newVersion, oldVersion) Then
                If Not background Then
                    DialogResult = MsgBox("You have the latest version of EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Else
                    returnstring = "Versions are the same!"
                End If
            ElseIf Version.op_LessThan(newVersion, oldVersion) Then
                returnstring = "Future version!"
            End If
            Return returnstring
        Catch ex As TimeoutException
            ' Timeout
            If Not background Then
                DialogResult = MsgBox("Check for version failed: timeout.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            End If
            GoTo skip
        Catch ex As WebException
            ' 404 error
            If Not background Then
                DialogResult = MsgBox("Check for version failed: web page missing.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            End If
            GoTo skip
        End Try
skip:
        Return returnstring
    End Function

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim result As String
        result = CheckForUpdate(False)
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim viewerfilename, inidirectory, inifilename As String
        Dim result As Windows.Forms.DialogResult
redo:
        ' get current template from ini file
        inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
        inifilename = inidirectory + "\EPubMetadataEditor.ini"
        If System.IO.File.Exists(inifilename) = False Then
            If System.IO.Directory.Exists(inidirectory) = False Then
                System.IO.Directory.CreateDirectory(inidirectory)
            End If
            Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
            Dim s As New StreamWriter(fs)

            ' look for ini file in old location
            Dim tempinifile = Application.StartupPath() + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(tempinifile) = True Then
                Dim tempobjIniFile As New IniFile(tempinifile)
                viewerfilename = tempobjIniFile.GetString("Viewer", "Path", "(none)")
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + viewerfilename + Chr(34))
            Else
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
            End If
            s.WriteLine("[Renamer]")
            s.WriteLine("Template=" + Chr(34) + "(none)" + Chr(34))
            s.Close()
        End If

        Dim objIniFile As New IniFile(inifilename)
        Dim template = objIniFile.GetString("Renamer", "Template", "(none)")
        Form4.ComboBox1.Items.Clear()
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template1", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template2", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template3", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template4", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template5", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If

        ' get current filename
        Form4.TextBox3.Text = System.IO.Path.GetFileNameWithoutExtension(OpenFileDialog1.FileName)
        If ((My.Computer.Keyboard.ShiftKeyDown) And (Form4.ComboBox1.Text <> "")) Then
            Form4.UpdateFilename()
            result = Windows.Forms.DialogResult.OK
        Else
            result = Form4.ShowDialog()
        End If
        If result = Windows.Forms.DialogResult.OK Then
            ' rename file
            Dim destFileName = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + Form4.TextBox2.Text + ".epub"
            Dim directoryName = System.IO.Path.GetDirectoryName(destFileName)
            If destFileName <> OpenFileDialog1.FileName Then
                If (Not Directory.Exists(directoryName)) Then
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destFileName))
                End If
                If destFileName.ToString.ToUpper = OpenFileDialog1.FileName.ToString.ToUpper Then
                    ' filenames differ only in case, so need to use a temp file
                    System.IO.File.Copy(OpenFileDialog1.FileName, "tempepub.epub")
                    System.IO.File.Delete(OpenFileDialog1.FileName)
                    System.IO.File.Copy("tempepub.epub", destFileName)
                    System.IO.File.Delete("tempepub.epub")
                    OpenFileDialog1.FileName = destFileName
                Else
                    If (File.Exists(destFileName)) Then
                        DialogResult = MsgBox("The following file already exists:" + Chr(10) + destFileName + Chr(10) + "Overwrite?", MsgBoxStyle.YesNo, "EPUB Metadata Editor")
                        If DialogResult = Windows.Forms.DialogResult.No Then
                            GoTo exitwithoutsaving
                        End If
                    End If
                    System.IO.File.Copy(OpenFileDialog1.FileName, destFileName, True)
                    System.IO.File.Delete(OpenFileDialog1.FileName)
                    OpenFileDialog1.FileName = destFileName
                End If
                If projectchanged Then
                    Me.Text = "*" + CaptionString
                Else
                    Me.Text = CaptionString
                End If
            End If
exitwithoutsaving:
            If Not ((My.Computer.Keyboard.ShiftKeyDown) And (Form4.ComboBox1.Text <> "")) Then
                template = objIniFile.GetString("Renamer", "Template", "(none)")
                If Form4.ComboBox1.Text <> template Then
                    ' update ini file, cycling through existing history, adding new template as most recent
                    template = objIniFile.GetString("Renamer", "Template4", "(none)")
                    objIniFile.WriteString("Renamer", "Template5", Chr(34) + template + Chr(34))
                    template = objIniFile.GetString("Renamer", "Template3", "(none)")
                    objIniFile.WriteString("Renamer", "Template4", Chr(34) + template + Chr(34))
                    template = objIniFile.GetString("Renamer", "Template2", "(none)")
                    objIniFile.WriteString("Renamer", "Template3", Chr(34) + template + Chr(34))
                    template = objIniFile.GetString("Renamer", "Template1", "(none)")
                    objIniFile.WriteString("Renamer", "Template2", Chr(34) + template + Chr(34))
                    template = objIniFile.GetString("Renamer", "Template", "(none)")
                    objIniFile.WriteString("Renamer", "Template1", Chr(34) + template + Chr(34))
                    objIniFile.WriteString("Renamer", "Template", Chr(34) + Form4.ComboBox1.Text + Chr(34))
                End If
            End If

            ' update form caption and file selector (just in case rename action has created a subfolder and moved the file into it)
            searchResults = Directory.GetFiles(IO.Path.GetDirectoryName(OpenFileDialog1.FileName), "*.epub", SearchOption.TopDirectoryOnly)
            refreshfilelist = True
            ComboBox3.Items.Clear()
            Dim fi As String
            For Each fi In searchResults
                ComboBox3.Items.Add(fi.Substring(fi.LastIndexOf("\") + 1, fi.Length - fi.LastIndexOf("\") - 1))
            Next
            Dim x As Integer = 0
            While (searchResults(x) <> OpenFileDialog1.FileName)
                x = x + 1
            End While
            currentfilenumber = x + 1 'searchResults is zero based
            ComboBox3.SelectedIndex = x
            ComboBox3.Enabled = True
            CaptionString = IO.Path.GetFileName(OpenFileDialog1.FileName) + " [" + currentfilenumber.ToString + "/" + searchResults.Length.ToString + "] - EPUB Metadata Editor"
            If Mid(Me.Text, 1, 1) = "*" Then
                Me.Text = "*" + CaptionString
            Else
                Me.Text = CaptionString
            End If
        End If
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        ' get current template from ini file
        Dim viewerfilename, inidirectory, inifilename As String
        inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
        inifilename = inidirectory + "\EPubMetadataEditor.ini"
        If System.IO.File.Exists(inifilename) = False Then
            If System.IO.Directory.Exists(inidirectory) = False Then
                System.IO.Directory.CreateDirectory(inidirectory)
            End If
            Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
            Dim s As New StreamWriter(fs)

            ' look for ini file in old location
            Dim tempinifile = Application.StartupPath() + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(tempinifile) = True Then
                Dim tempobjIniFile As New IniFile(tempinifile)
                viewerfilename = tempobjIniFile.GetString("Viewer", "Path", "(none)")
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + viewerfilename + Chr(34))
            Else
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
            End If
            s.WriteLine("[Renamer]")
            s.WriteLine("Template=" + Chr(34) + "(none)" + Chr(34))
            s.Close()
        End If

        Dim objIniFile As New IniFile(inifilename)
        Dim template = objIniFile.GetString("Renamer", "Template", "(none)")
        Form4.ComboBox1.Items.Clear()
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        Else
            Form4.ComboBox1.Items.Add("")
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template1", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template2", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template3", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template4", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Renamer", "Template5", "(none)")
        If template <> "(none)" Then
            Form4.ComboBox1.Items.Add(template)
            Form4.ComboBox1.SelectedIndex = 0
        End If

        ' change form
        Form4.TextBox2.Text = ""
        Form4.TextBox2.Enabled = False
        Form4.Button7.Enabled = False
        Form4.Button3.Text = "OK"

        If Form4.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ' check to see if current scheme has changed
            template = objIniFile.GetString("Renamer", "Template", "(none)")
            If Form4.ComboBox1.Text <> template Then
                ' update ini file, cycling through existing history, adding new template as most recent
                template = objIniFile.GetString("Renamer", "Template4", "(none)")
                objIniFile.WriteString("Renamer", "Template5", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Renamer", "Template3", "(none)")
                objIniFile.WriteString("Renamer", "Template4", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Renamer", "Template2", "(none)")
                objIniFile.WriteString("Renamer", "Template3", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Renamer", "Template1", "(none)")
                objIniFile.WriteString("Renamer", "Template2", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Renamer", "Template", "(none)")
                objIniFile.WriteString("Renamer", "Template1", Chr(34) + template + Chr(34))
                objIniFile.WriteString("Renamer", "Template", Chr(34) + Form4.ComboBox1.Text + Chr(34))
            End If
        End If

        ' put form back to normal
        Form4.TextBox2.Enabled = True
        Form4.Button7.Enabled = True
        Form4.Button3.Text = "Rename"
    End Sub

    Private Sub wait(ByVal interval As Integer)
        ' Loops for a specificied period of time (milliseconds)
        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
            ' Allows UI to remain responsive
            Application.DoEvents()
        Loop
        sw.Stop()
    End Sub
    Private Sub asyncwait(ByVal interval As Integer)
        ' Loops for a specificied period of time (milliseconds)
        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
        Loop
        sw.Stop()
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        ' get current template from ini file
        Dim viewerfilename, inidirectory, inifilename As String
        inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
        inifilename = inidirectory + "\EPubMetadataEditor.ini"
        If System.IO.File.Exists(inifilename) = False Then
            If System.IO.Directory.Exists(inidirectory) = False Then
                System.IO.Directory.CreateDirectory(inidirectory)
            End If
            Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
            Dim s As New StreamWriter(fs)

            ' look for ini file in old location
            Dim tempinifile = Application.StartupPath() + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(tempinifile) = True Then
                Dim tempobjIniFile As New IniFile(tempinifile)
                viewerfilename = tempobjIniFile.GetString("Viewer", "Path", "(none)")
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + viewerfilename + Chr(34))
            Else
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
            End If
            s.WriteLine("[Renamer]")
            s.WriteLine("Template=" + Chr(34) + "(none)" + Chr(34))
            s.Close()
        End If

        Dim objIniFile As New IniFile(inifilename)
        Dim template = objIniFile.GetString("Renamer", "Template", "(none)")
        If template = "(none)" Then
            ' change form
            Form4.TextBox2.Text = ""
            Form4.TextBox2.Enabled = False
            Form4.Button7.Enabled = False
            Form4.Button3.Text = "OK"

            If Form4.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ' update ini file
                objIniFile.WriteString("Renamer", "Template", Chr(34) + Form4.ComboBox1.Text + Chr(34))
            Else
                ' cancelled action
                Exit Sub
            End If

            ' put form back to normal
            Form4.TextBox2.Enabled = True
            Form4.Button7.Enabled = True
            Form4.Button3.Text = "Rename"
        End If

        template = objIniFile.GetString("Renamer", "Template", "(none)")
        If template <> "(none)" Then
            ' process all files, applying template
            Dim filenum, x As Integer
            Dim metadatafile As String

            ClearInterface()
            tempdirectory = System.IO.Path.GetTempPath
            ebookdirectory = tempdirectory + "EPUB"

            filenum = ListBox1.Items.Count
            ProgressBar1.Maximum = filenum - 1
            ProgressBar1.Visible = True

            For x = 1 To filenum
                ChDir(tempdirectory)
                ProgressBar1.Value = x - 1
                ProgressBar1.Update()
                Application.DoEvents()

                ' open file
                'Unzip epub to temp directory

                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        'delete contents of temp directory
                        DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                    Catch
                        wait(500)
                        'try again
                        DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                    End Try
                Else
                    MkDir(ebookdirectory)
                End If
                ChDir(ebookdirectory)

                Try
                    Dim zip As ZipStorer
                    zip = ZipStorer.Open(ListBox1.Items(x - 1).ToString, FileAccess.Read)
                    Dim dir = zip.ReadCentralDir()
                    Dim item As ZipStorer.ZipFileEntry
                    For Each item In dir
                        zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
                    Next
                    zip.Close()
                Catch ex1 As Exception
                    Console.Error.WriteLine("exception: {0}", ex1.ToString)
                    DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Exit Sub

                End Try

                'Search for .opf file
                searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

                'Open .opf file into RichTextBox
                If searchResults.Length < 1 Then
                    DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Return
                Else
                    opffile = searchResults(0)
                    If InStr(opffile, "_MACOSX") Then
                        If searchResults.Length > 1 Then
                            opffile = searchResults(1)
                        End If
                    End If
                    opfdirectory = Path.GetDirectoryName(opffile)
                    RichTextBox1.Text = LoadUnicodeFile(opffile)
                End If

                'Extract metadata into textboxes (but no need to extract cover)
                metadatafile = LoadUnicodeFile(opffile)
                ExtractMetadata(metadatafile, False)

                'Convert metadata into new filename
                Dim currpos, endpos, temppos, field, nextchar, insertText, newFileName
                newFileName = ""
                currpos = 0
                While (currpos < Len(template))
                    currpos = currpos + 1

                    ' look for field marker
                    If (Mid(template, currpos, 1) = "%") Then
                        If (Mid(template, currpos + 1, 1) = "%") Then
                            ' found '%%' (replace with '%')
                            newFileName = newFileName + "%"
                            currpos = currpos + 1
                        Else
                            ' look for end field marker
                            endpos = InStr(currpos + 1, template, "%")
                            If (endpos <> 0) Then
                                ' end field marker found
                                field = Mid(template, currpos + 1, endpos - currpos - 1)
                                insertText = ""
                                If field = "Creator" Then
                                    insertText = TextBox2.Text
                                ElseIf field = "CREATOR" Then
                                    insertText = TextBox2.Text.ToUpper
                                ElseIf field = "CreatorFileAs" Then
                                    insertText = TextBox12.Text
                                ElseIf field = "CREATORFILEAS" Then
                                    insertText = TextBox12.Text.ToUpper
                                ElseIf ((field = "CreatorSurnameOnly") Or (field = "CREATORSURNAMEONLY")) Then
                                    insertText = TextBox2.Text
                                    If InStr(insertText, " ") <> 0 Then
                                        temppos = Len(insertText)
                                        nextchar = Mid(insertText, temppos, 1)
                                        While (nextchar <> " ")
                                            If temppos = 1 Then
                                                GoTo errortext
                                            End If
                                            temppos = temppos - 1
                                            nextchar = Mid(insertText, temppos, 1)
                                        End While
                                        insertText = Mid(insertText, temppos + 1)
                                        If (Mid(TextBox2.Text, temppos - 1, 1) = ",") Then
                                            insertText = TextBox2.Text
                                            temppos = 1
                                            nextchar = Mid(insertText, temppos, 1)
                                            While (nextchar <> ",")
                                                If temppos = Len(insertText) Then
                                                    GoTo errortext
                                                End If
                                                temppos = temppos + 1
                                                nextchar = Mid(insertText, temppos, 1)
                                            End While
                                            insertText = Mid(insertText, 1, temppos - 1)
                                        End If
                                    End If
                                    If (field = "CREATORSURNAMEONLY") Then
                                        insertText = insertText.ToUpper
                                    End If
                                ElseIf ((field = "CreatorFirstInitial") Or (field = "CREATORFIRSTINITIAL")) Then
                                    insertText = TextBox2.Text
                                    If InStr(insertText, " ") <> 0 Then
                                        temppos = Len(insertText)
                                        nextchar = Mid(insertText, temppos, 1)
                                        While (nextchar <> " ")
                                            If temppos = 1 Then
                                                GoTo errortext
                                            End If
                                            temppos = temppos - 1
                                            nextchar = Mid(insertText, temppos, 1)
                                        End While
                                        insertText = Mid(insertText, temppos + 1)
                                        If (Mid(TextBox2.Text, temppos - 1, 1) = ",") Then
                                            insertText = TextBox2.Text
                                            temppos = 1
                                            nextchar = Mid(insertText, temppos, 1)
                                            While (nextchar <> ",")
                                                If temppos = Len(insertText) Then
                                                    GoTo errortext
                                                End If
                                                temppos = temppos + 1
                                                nextchar = Mid(insertText, temppos, 1)
                                            End While
                                            insertText = Mid(insertText, 1, temppos - 1)
                                        End If
                                    End If
                                    If Len(insertText) > 1 Then
                                        insertText = Mid(insertText, 1, 1)
                                    End If
                                    If (field = "CREATORFIRSTINITIAL") Then
                                        insertText = insertText.ToUpper
                                    End If
                                ElseIf field = "Title" Then
                                    insertText = TextBox1.Text
                                ElseIf field = "TITLE" Then
                                    insertText = TextBox1.Text.ToUpper
                                ElseIf field = "TitleFileAs" Then
                                    insertText = TextBox16.Text
                                ElseIf field = "TITLEFILEAS" Then
                                    insertText = TextBox16.Text.ToUpper
                                ElseIf field = "Series" Then
                                    insertText = TextBox15.Text
                                ElseIf field = "SERIES" Then
                                    insertText = TextBox15.Text.ToUpper
                                ElseIf field = "SeriesIndex" Then
                                    insertText = TextBox14.Text
                                ElseIf field = "SERIESINDEX" Then
                                    insertText = TextBox14.Text.ToUpper
                                ElseIf field = "Date" Then
                                    insertText = TextBox6.Text
                                ElseIf field = "DATE" Then
                                    insertText = TextBox6.Text.ToUpper
                                Else
                                    insertText = ""
                                End If
errortext:
                                newFileName = newFileName + insertText
                                currpos = endpos
                            End If
                        End If
                    Else
                        newFileName = newFileName + Mid(template, currpos, 1)
                    End If
                End While

                'Replace illegal characters
                Dim Letter As String
                Dim pos As Integer = 0
                Dim charactersDisallowed As String = "/:*?<>|" + Chr(34)
                While pos < Len(newFileName)
                    Letter = newFileName.Substring(pos, 1)
                    If charactersDisallowed.Contains(Letter) Then
                        newFileName = newFileName.Replace(Letter, "-")
                    End If
                    pos = pos + 1
                End While

                'Replace html characters
                pos = 1
                While pos < Len(newFileName)
                    If (Mid(newFileName, pos, 5) = "&amp;") Then
                        newFileName = newFileName.Replace("&amp;", "&")
                    ElseIf (Mid(newFileName, pos, 6) = "&quot;") Then
                        newFileName = newFileName.Replace("&quot;", "'")
                    End If
                    pos = pos + 1
                End While

                'Rename file
                newFileName = System.IO.Path.GetDirectoryName(ListBox1.Items(x - 1)) + "\" + newFileName + ".epub"
                Dim directoryName = System.IO.Path.GetDirectoryName(newFileName)
                If newFileName <> ListBox1.Items(x - 1) Then
                    If (Not Directory.Exists(directoryName)) Then
                        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(newFileName))
                    End If
                    If newFileName.ToString.ToUpper = ListBox1.Items(x - 1).ToString.ToUpper Then
                        ' filenames differ only in case, so need to use a temp file
                        System.IO.File.Copy(ListBox1.Items(x - 1), "tempepub.epub")
                        System.IO.File.Delete(ListBox1.Items(x - 1))
                        System.IO.File.Copy("tempepub.epub", newFileName)
                        System.IO.File.Delete("tempepub.epub")
                        ListBox1.Items(x - 1) = newFileName
                    Else
                        If (File.Exists(newFileName)) Then
                            DialogResult = MsgBox("The following file already exists:" + Chr(10) + newFileName + Chr(10) + "Overwrite?", MsgBoxStyle.YesNo, "EPUB Metadata Editor")
                            If DialogResult = Windows.Forms.DialogResult.No Then
                                GoTo exitwithoutsaving
                            End If
                        End If
                        System.IO.File.Copy(ListBox1.Items(x - 1), newFileName, True)
                        System.IO.File.Delete(ListBox1.Items(x - 1))
                        ListBox1.Items(x - 1) = newFileName
                    End If
                End If
exitwithoutsaving:
                ClearInterface()
            Next

            ListBox1.Sorted = True
            ListBox1.Sorted = False
            ProgressBar1.Value = 0
            ProgressBar1.Update()
            ProgressBar1.Visible = False
            projectchanged = False
            Button3.Enabled = False
            Me.Text = "EPUB Metadata Editor"
            ClearInterface()
            DialogResult = MsgBox("All done!", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")

            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                'delete contents of temp directory
                DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
            End If
        End If
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        OpenFileDialog5.Filter = "(X)HTML files|*.htm;*.html;*.xhtml"
        OpenFileDialog5.FilterIndex = 1
        OpenFileDialog5.FileName = ""
        OpenFileDialog5.InitialDirectory = ebookdirectory
        If OpenFileDialog5.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myuri As New Uri(OpenFileDialog5.FileName)
            Form6.WebBrowser1.Url = myuri
            Form6.Text = "EPUB Metadata Editor - " + IO.Path.GetFileName(OpenFileDialog5.FileName)
            Button33.Enabled = False
            Form6.Show()
        End If
    End Sub

    Private Sub ProcessTOCNCX(ByVal file As String)
        Dim tempstring As String
        Dim startpos, endpos, currpos, doclength, depth, textpos, nodenums As Integer
        Dim currentrootnode(5) As Integer
        Dim uplevel, downlevel As Boolean

        RichTextBox1.Text = LoadUnicodeFile(file)

        'Search for doctitle
        tempstring = LCase(RichTextBox1.Text)
        startpos = InStr(tempstring, "<doctitle>")
        If startpos <> 0 Then
            currpos = InStr(startpos, tempstring, "<text>")
            If currpos <> 0 Then
                endpos = InStr(currpos, tempstring, "</text>")
                If endpos <> 0 Then
                    Form3.Text = Mid(RichTextBox1.Text, currpos + 6, endpos - currpos - 6)
                End If
            End If
        End If

        'Search for author
        startpos = InStr(tempstring, "<docauthor>")
        If startpos <> 0 Then
            currpos = InStr(startpos, tempstring, "<text>")
            If currpos <> 0 Then
                endpos = InStr(currpos, tempstring, "</text>")
                If endpos <> 0 Then
                    Form3.Text = Form3.Text + " by " + Mid(RichTextBox1.Text, currpos + 6, endpos - currpos - 6)
                End If
            End If
        End If

        Dim NewNode() As VistaNode
        nodenums = 0
        Form3.VistaTreeView1.Nodes.Clear()
        doclength = Len(tempstring)
        depth = 0
        uplevel = False
        currentrootnode(0) = -1
        currentrootnode(1) = -1
        currentrootnode(2) = -1
        currentrootnode(3) = -1
        currentrootnode(4) = -1
        currentrootnode(5) = -1

        startpos = InStr(tempstring, "<navmap>")
        currpos = startpos
        If startpos <> 0 Then
            While currpos <= doclength
                currpos = currpos + 1
                If Mid(tempstring, currpos, 9) = "<navpoint" Then
                    If uplevel = True Then
                        depth = depth + 1
                    End If
                    uplevel = True
                    downlevel = False
                    currentrootnode(depth) = currentrootnode(depth) + 1
                    ' Look for node label
                    textpos = InStr(currpos, tempstring, "<text>")
                    If textpos <> 0 Then
                        endpos = InStr(textpos, tempstring, "</text>")
                        If endpos <> 0 Then
                            If depth = 0 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 1 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 2 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes(currentrootnode(1)).Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 3 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes(currentrootnode(1)).Nodes(currentrootnode(2)).Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 4 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes(currentrootnode(1)).Nodes(currentrootnode(2)).Nodes(currentrootnode(3)).Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 5 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes(currentrootnode(1)).Nodes(currentrootnode(2)).Nodes(currentrootnode(3)).Nodes(currentrootnode(4)).Nodes.Add(NewNode(nodenums))
                            End If
                            If depth = 6 Then
                                nodenums = nodenums + 1
                                ReDim NewNode(nodenums)
                                NewNode(nodenums) = New VistaNode
                                NewNode(nodenums).Text = Mid(RichTextBox1.Text, textpos + 6, endpos - textpos - 6)
                                Form3.VistaTreeView1.Nodes(currentrootnode(0)).Nodes(currentrootnode(1)).Nodes(currentrootnode(2)).Nodes(currentrootnode(3)).Nodes(currentrootnode(4)).Nodes(currentrootnode(5)).Nodes.Add(NewNode(nodenums))
                            End If
                        End If
                    End If
                End If

                If Mid(tempstring, currpos, 11) = "</navpoint>" Then
                    If downlevel = True Then
                        currentrootnode(depth) = -1
                        depth = depth - 1
                    End If
                    downlevel = True
                    uplevel = False
                End If
            End While
        End If

        Dim currentNode As TreeNode = Form3.VistaTreeView1.GetNodeAt(0, 0)
        Form3.VistaTreeView1.ExpandAll()
        If Not currentNode Is Nothing Then currentNode.EnsureVisible()
        Form3.ToolStripMenuItem1.Checked = True
        Button23.Enabled = False
    End Sub

    Private Sub LinkLabel5_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        ProcessTOCNCX(tocncxfile)
        Form3.Show()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        fileeditorreturn = False
        filecontents = ""
        Form2.Button3.Visible = False
        Form2.Button9.Visible = False
        Form2.RichTextBox1.Text = LoadUnicodeFile(tocncxfile)
        Form2.ShowDialog()
        If fileeditorreturn = True Then
            RichTextBox1.Text = filecontents
            SaveUnicodeFile(tocncxfile, RichTextBox1.Text)
            projectchanged = True
            Button3.Enabled = True
            Me.Text = "*" + CaptionString
        End If
    End Sub

    Private Sub ContextMenuStrip2_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        If (DirectCast(Me.ActiveControl, RichTextBox).CanUndo) Then
            ContextMenuStrip2.Items(0).Enabled = True
        Else
            ContextMenuStrip2.Items(0).Enabled = False
        End If

        If (DirectCast(Me.ActiveControl, RichTextBox).SelectedText.Length = 0) Then
            ContextMenuStrip2.Items(2).Enabled = False
            ContextMenuStrip2.Items(3).Enabled = False
        Else
            ContextMenuStrip2.Items(2).Enabled = True
            ContextMenuStrip2.Items(3).Enabled = True
        End If

        If (Clipboard.ContainsText()) Then
            ContextMenuStrip2.Items(4).Enabled = True
        Else
            ContextMenuStrip2.Items(4).Enabled = False
        End If

        If (DirectCast(Me.ActiveControl, RichTextBox).Text.Length = 0) Then
            ContextMenuStrip2.Items(6).Enabled = False
        Else
            ContextMenuStrip2.Items(6).Enabled = True
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
    Private Function GetMimeType(ByVal fileName As String)
        Dim mimeType As String = "application/unknown"
        Dim ext As String = Path.GetExtension(fileName).ToLower()
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext)
        If (regKey.GetValue("Content Type") <> Nothing) Then
            mimeType = regKey.GetValue("Content Type").ToString()
        End If
        Return mimeType
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim result As String
        result = CheckForUpdate(True)
        If result = "Update available!" Then
            BackgroundWorker1.ReportProgress(100)
        ElseIf result = "Versions are the same!" Then
            BackgroundWorker1.ReportProgress(50)
        ElseIf result = "Future version!" Then
            BackgroundWorker1.ReportProgress(1)
        Else
            BackgroundWorker1.ReportProgress(0)
        End If
    End Sub

    Private Sub LinkLabel6_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        DialogResult = Dialog3.ShowDialog
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If e.ProgressPercentage = 100 Then
            LinkLabel6.Visible = True
            LinkLabel2.Visible = False
            Dialog3.Label4.Text = updateinfo
        ElseIf e.ProgressPercentage = 50 Then
            LinkLabel2.Text = ""
            LinkLabel6.Visible = False
            LinkLabel2.Visible = True
        ElseIf e.ProgressPercentage = 1 Then
            LinkLabel2.Text = "Developer!"
            LinkLabel6.Visible = False
            LinkLabel2.Visible = True
        Else
            LinkLabel6.Visible = False
            LinkLabel2.Visible = True
        End If
    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If (Clipboard.ContainsImage()) Then
            ContextMenuStrip1.Items(4).Enabled = True
        Else
            ContextMenuStrip1.Items(4).Enabled = False
        End If
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        Dim metadatafile, newrelativecoverimagefile As String
        Dim startpos, endpos, pos, temppos As Integer

        RichTextBox1.Text = LoadUnicodeFile(opffile)
        metadatafile = LoadUnicodeFile(opffile)
        newrelativecoverimagefile = relativecoverimagefile

        If fixcovermetadata Then
            ' Add cover image information to metadata
            ' e.g. <meta content="cover.jpg" name="cover" />
            pos = InStr(metadatafile, "</metadata>")
            metadatafile = Mid(metadatafile, 1, pos - 1) + "<meta content=" + Chr(34) + coverimagefilename + Chr(34) + " name=" + Chr(34) + "cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, pos)
            fixcovermetadata = False
        End If

        If fixcovermanifest Then
            ' Add cover image information to manifest
            ' e.g. <item href="Images/cover.jpg" id="cover" media-type="image/jpeg" />

            ' First, get relative cover image file location
            startpos = InStr(metadatafile, "<manifest")
            pos = InStr(startpos, metadatafile, coverimagefilename)
            If pos <> 0 Then
                startpos = InStrRev(metadatafile, "<", pos)
                temppos = InStrRev(metadatafile, "href=", pos)
                If temppos < startpos Then
                    temppos = InStr(pos, metadatafile, "href=")
                End If
                endpos = InStr(temppos + 6, metadatafile, Chr(34))
                newrelativecoverimagefile = Mid(metadatafile, temppos + 6, endpos - temppos - 6)
            End If

            pos = InStr(metadatafile, "</manifest>")
            metadatafile = Mid(metadatafile, 1, pos - 1) + "<item href=" + Chr(34) + newrelativecoverimagefile + Chr(34) + " id=" + Chr(34) + "cover" + Chr(34) + " media-type=" + Chr(34) + "image/jpeg" + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, pos)
            fixcovermanifest = False
        End If

        RichTextBox1.Text = metadatafile
        SaveUnicodeFile(opffile, metadatafile)

        Button35.Visible = False
        Label27.Visible = False
        If ((Button27.Visible = False) And (Button1.Visible = False)) Then
            Button42.Visible = False
        End If

        projectchanged = True
        Button3.Enabled = True
        Me.Text = "*" + CaptionString
    End Sub

    Private Sub UseExistingImageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UseExistingImageToolStripMenuItem.Click
        Dim result As DialogResult
        Dim ImageDirectory, FileNameOnly, RelativeLocation, metadatafile As String
        Dim startpos, endpos As Integer
        OpenFileDialog6.InitialDirectory = opfdirectory
        OpenFileDialog6.Filter = "Image files (*.jpg, *.jpeg, *.png) | *.jpg; *.jpeg; *.png"
        result = OpenFileDialog6.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            ImageDirectory = Path.GetDirectoryName(OpenFileDialog6.FileName)
            If InStr(ImageDirectory, opfdirectory) = 0 Then
                DialogResult = MsgBox("ERROR: You can only select an image that is already in the EPUB file." + Chr(10) + "If you want to select an image that is not already in the EPUB file, use 'Add image...' or 'Change image...'", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Else
                RelativeLocation = ImageDirectory.Replace(opfdirectory, "")
                FileNameOnly = Path.GetFileName(OpenFileDialog6.FileName)
                If Mid(RelativeLocation, 1, 1) = "\" Then
                    RelativeLocation = Mid(RelativeLocation, 2)
                End If
                RelativeLocation = RelativeLocation.Replace("\", "/")

                RichTextBox1.Text = LoadUnicodeFile(opffile)
                metadatafile = LoadUnicodeFile(opffile)

                If (InStr(metadatafile, "<guide>") = 0) Then
                    endpos = InStr(metadatafile, "</package>")
                    If endpos <> 0 Then
                        metadatafile = Mid(metadatafile, 1, endpos - 1) + "<guide>" + Chr(13) + Chr(10) + Chr(9) + "<reference href=" + Chr(34) + RelativeLocation + "/" + FileNameOnly + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + "</guide>" + Chr(13) + Chr(10) + Mid(metadatafile, endpos)
                    End If
                Else
                    startpos = InStr(metadatafile, "<guide>")
                    endpos = InStr(startpos, metadatafile, "type=" + Chr(34) + "cover")
                    If endpos = 0 Then
                        metadatafile = Mid(metadatafile, 1, startpos + 7) + Chr(9) + "<reference href=" + Chr(34) + RelativeLocation + "/" + FileNameOnly + Chr(34) + " type=" + Chr(34) + "cover" + Chr(34) + " title=" + Chr(34) + "Cover" + Chr(34) + "/>" + Chr(13) + Chr(10) + Mid(metadatafile, startpos + 8)
                    Else
                        While (Mid(metadatafile, endpos, 5) <> "href=")
                            endpos = endpos - 1
                        End While
                        startpos = endpos
                        endpos = InStr(startpos + 7, metadatafile, Chr(34))
                        metadatafile = Mid(metadatafile, 1, startpos + 5) + RelativeLocation + "/" + FileNameOnly + Mid(metadatafile, endpos)
                    End If
                End If

                RichTextBox1.Text = metadatafile
                SaveUnicodeFile(opffile, metadatafile)

                SaveImageAsToolStripMenuItem.Enabled = True
                AddImageToolStripMenuItem.Enabled = False
                ChangeImageToolStripMenuItem.Enabled = True

                projectchanged = True
                Button3.Enabled = True
                Button27.Visible = True
                Label23.Visible = True
                Button42.Visible = True
                Me.Text = "*" + CaptionString

                ' Need to update metadata
                RichTextBox1.Text = LoadUnicodeFile(opffile)
                metadatafile = LoadUnicodeFile(opffile)
                ExtractMetadata(metadatafile, True)
            End If
        End If
    End Sub


    Private Sub LinkLabel7_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=KC9T4JCJ2MPZG")
    End Sub
    Private Function XMLOutput(ByVal InputString As String) As String
        Dim x As Integer
        Dim OutputString, nextchar As String
        OutputString = ""
        For x = 1 To Len(InputString)
            nextchar = Mid(InputString, x, 1)
            If ((nextchar = "&") Or (nextchar = "<") Or (nextchar = ">") Or (nextchar = Chr(34)) Or (nextchar = "'") Or (nextchar = Chr(10))) Then
                If (nextchar = "&") Then
                    OutputString = OutputString + "&amp;"
                ElseIf (nextchar = "<") Then
                    OutputString = OutputString + "&lt;"
                ElseIf (nextchar = ">") Then
                    OutputString = OutputString + "&gt;"
                ElseIf (nextchar = Chr(34)) Then
                    OutputString = OutputString + "&quot;"
                ElseIf (nextchar = "'") Then
                    OutputString = OutputString + "&apos;"
                ElseIf (nextchar = Chr(10)) Then
                    OutputString = OutputString + "&#10;"
                End If
            Else
                OutputString = OutputString + nextchar
            End If
        Next
        Return OutputString
    End Function
    Private Function XMLInput(ByVal InputString As String) As String
        Dim x, length As Integer
        Dim OutputString, nextchars As String
        Dim DidSomething As Boolean
        OutputString = ""
        length = Len(InputString)
        x = 1
        While (x <= length)
            DidSomething = False
            If (x + 3 <= length) Then
                nextchars = Mid(InputString, x, 4)
                If (nextchars = "&lt;") Then
                    OutputString = OutputString + "<"
                    x = x + 4
                    DidSomething = True
                ElseIf (nextchars = "&gt;") Then
                    OutputString = OutputString + ">"
                    x = x + 4
                    DidSomething = True
                End If
            End If
            If (x + 4 <= length) Then
                nextchars = Mid(InputString, x, 5)
                If (nextchars = "&amp;") Then
                    OutputString = OutputString + "&"
                    x = x + 5
                    DidSomething = True
                ElseIf (nextchars = "&#10;") Then
                    OutputString = OutputString + Chr(10)
                    x = x + 5
                    DidSomething = True
                End If
            End If
            If (x + 5 <= length) Then
                nextchars = Mid(InputString, x, 6)
                If (nextchars = "&quot;") Then
                    OutputString = OutputString + Chr(34)
                    x = x + 6
                    DidSomething = True
                ElseIf (nextchars = "&apos;") Then
                    OutputString = OutputString + "'"
                    x = x + 6
                    DidSomething = True
                End If
            End If
            If Not DidSomething Then
                OutputString = OutputString + Mid(InputString, x, 1)
                x = x + 1
            End If
        End While
        Return OutputString
    End Function

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        If (WebBrowser1.Visible = False) Then
            WebBrowser1.DocumentText = "<head><style type=" + Chr(34) + "text/css" + Chr(34) + ">" + "body {margin:0;padding:0}" + "</style></head>" + TextBox4.Text.Replace(Chr(10), "<br>")
            WebBrowser1.Visible = True
            Button38.Text = "E"
            ToolTip1.SetToolTip(Button38, "Edit Description")
        Else
            WebBrowser1.Visible = False
            Button38.Text = "OK"
            ToolTip1.SetToolTip(Button38, "Show Changes")
        End If

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If refreshfilelist Then
            refreshfilelist = False
        Else
            If (DealWithPreviousFile() = "proceed") Then
                OpenFileDialog1.FileName = searchResults(ComboBox3.SelectedIndex)
                OpenEPUB()
                Button3.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        ' move file up
        'Make sure our item is not the first one on the list.
        If ListBox1.SelectedIndex > 0 Then
            Dim I = ListBox1.SelectedIndex - 1
            ListBox1.Items.Insert(I, ListBox1.SelectedItem)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ListBox1.SelectedIndex = I
        End If
    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        ' move file down
        'Make sure our item is not the last one on the list.
        If ListBox1.SelectedIndex < ListBox1.Items.Count - 1 Then
            'Insert places items above the index you supply, since we want
            'to move it down the list we have to do + 2
            Dim I = ListBox1.SelectedIndex + 2
            ListBox1.Items.Insert(I, ListBox1.SelectedItem)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ListBox1.SelectedIndex = I - 1
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItem = Nothing Then
            Exit Sub
        End If

        Dim metadatafile As String
        Button39.Enabled = True
        Button40.Enabled = True
        Button45.Enabled = True

        ' Load EPUB metadata
        ' open file
        'Unzip epub to temp directory
        tempdirectory = System.IO.Path.GetTempPath
        ebookdirectory = tempdirectory + "EPUB"

        ChDir(tempdirectory)

        If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
            Try
                'delete contents of temp directory
                DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
            Catch
                wait(500)
                'try again
                DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
            End Try
        Else
            MkDir(ebookdirectory)
        End If
        ChDir(ebookdirectory)

        Try
            Dim zip As ZipStorer
            zip = ZipStorer.Open(ListBox1.SelectedItem.ToString, FileAccess.Read)
            Dim dir = zip.ReadCentralDir()
            Dim item As ZipStorer.ZipFileEntry
            For Each item In dir
                zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
            Next
            zip.Close()
        Catch ex1 As Exception
            Console.Error.WriteLine("exception: {0}", ex1.ToString)
            DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "The ebook " + ListBox1.SelectedItem + " cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Exit Sub
            Exit Sub
        End Try

        'Search for .opf file
        searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

        'Open .opf file into RichTextBox
        If searchResults.Length < 1 Then
            DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "The ebook " + ListBox1.SelectedItem + " is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Return
        Else
            opffile = searchResults(0)
            If InStr(opffile, "_MACOSX") Then
                If searchResults.Length > 1 Then
                    opffile = searchResults(1)
                End If
            End If
            opfdirectory = Path.GetDirectoryName(opffile)
            RichTextBox1.Text = LoadUnicodeFile(opffile)
        End If

        'Extract metadata into textboxes (but no need to extract cover)
        ClearInterface()
        metadatafile = LoadUnicodeFile(opffile)
        ExtractMetadata(metadatafile, True)

        projectchanged = False
        CaptionString = "EPUB Metadata Editor"
        Me.Text = CaptionString
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        ListBox1.Sorted = True
        ListBox1.Sorted = False
    End Sub

    Private Sub DisableInterface()
        TextBox1.Enabled = False
        TextBox16.Enabled = False
        TextBox2.Enabled = False
        TextBox12.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox13.Enabled = False
        ComboBox2.Enabled = False
        TextBox15.Enabled = False
        TextBox14.Enabled = False
        WebBrowser1.Visible = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox17.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        Button21.Enabled = False
        Button25.Enabled = False
        Button22.Enabled = False
        Button18.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button15.Enabled = False
        Button14.Enabled = False
        Button13.Enabled = False
        Button38.Enabled = False
        Button28.Enabled = False
        Button29.Enabled = False
        Button35.Enabled = False
        Button1.Enabled = False
        Button27.Enabled = False
        Label27.Visible = False
        Label4.Visible = False
        Label23.Visible = False
        CheckBox5.Enabled = False
        PictureBox1.Enabled = False
        GroupBox1.Visible = False
        ComboBox3.Enabled = False
        Button3.Enabled = False
        Button8.Enabled = False
        Button23.Enabled = False
        LinkLabel3.Enabled = False
        LinkLabel5.Enabled = False
        Button42.Visible = False
        Button35.Visible = False
        Button1.Visible = False
        Button27.Visible = False
        Button19.Enabled = False
        Button20.Enabled = False
        Button34.Enabled = False
        Button24.Enabled = False
        Button26.Enabled = False
        Button33.Enabled = False
        CheckBox5.Visible = False
    End Sub
    Private Function TitleCase(ByVal stringtext As String) As String
        ' Capitalise according to The U.S. Government Printing Office Style Manual (http://www.gpoaccess.gov/stylemanual/browse.html)

        Dim words() As String = Split(WordsNotToCapitalise, ",")
        Dim result() As String = Split(stringtext, " ")
        Dim wordnum, num, x, y As Integer
        Dim inthelist As Boolean
        Dim returnstring As String
        wordnum = words.Length
        num = result.Length
        returnstring = ""
        If num > 0 Then
            ' Capitalise first word
            result(0) = StrConv(result(0), VbStrConv.ProperCase)

            ' Captilise if not in the list of words or after a colon
            For x = 1 To num - 2
                inthelist = False
                For y = 0 To wordnum - 1
                    If StrConv(result(x), VbStrConv.Lowercase) = words(y) Then
                        inthelist = True
                        Exit For
                    End If
                Next
                If (inthelist And (Mid(result(x - 1), result(x - 1).Length, 1) <> ":")) Then
                    result(x) = StrConv(result(x), VbStrConv.Lowercase)
                Else
                    result(x) = StrConv(result(x), VbStrConv.ProperCase)
                End If
            Next

            ' Capitalise last word
            result(num - 1) = StrConv(result(num - 1), VbStrConv.ProperCase)

            returnstring = result(0)
            For x = 1 To num - 1
                returnstring = returnstring + " " + result(x)
            Next
        Else
            returnstring = StrConv(stringtext, VbStrConv.ProperCase)
        End If
        Return returnstring
    End Function

    Private Sub LinkLabel8_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Dim NewWordsNotToCapitalise As String
        Dim inidirectory, inifilename As String
        Dialog2.TextBox1.Text = WordsNotToCapitalise
        Dialog2.ShowDialog()
        NewWordsNotToCapitalise = Dialog2.TextBox1.Text
        If ((Dialog2.DialogResult = Windows.Forms.DialogResult.OK) And (WordsNotToCapitalise <> NewWordsNotToCapitalise)) Then
            WordsNotToCapitalise = NewWordsNotToCapitalise

            inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
            inifilename = inidirectory + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(inifilename) = False Then
                If System.IO.Directory.Exists(inidirectory) = False Then
                    System.IO.Directory.CreateDirectory(inidirectory)
                End If
                Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
                Dim s As New StreamWriter(fs)
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
                s.Close()
            End If

            Dim objIniFile As New IniFile(inifilename)
            objIniFile.WriteString("Editor", "Words", WordsNotToCapitalise)
        End If
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        Dim metadatafile As String

        metadatafile = LoadUnicodeFile(opffile)
        metadatafile = CleanOPF(metadatafile)
        metadatafile = Regularise(metadatafile)
        SaveUnicodeFile(opffile, metadatafile)

        If Button35.Visible Then
            Button35_Click(sender, e)
        End If
        If Button1.Visible Then
            Button1_Click(sender, e)
        End If
        If Button27.Visible Then
            Button27_Click(sender, e)
        End If
        SaveEpub(OpenFileDialog1.FileName, False)
        Button42.Visible = False
    End Sub

    Private Sub LinkLabel9_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        Form7.ShowDialog()
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        Form8.ShowDialog()
    End Sub
    Public Sub FindFilesInCurrentFolder()
        Dim filenum, x As Integer

        tempdirectory = System.IO.Path.GetTempPath
        ebookdirectory = tempdirectory + "EPUB"

        filenum = ComboBox3.Items.Count
        ProgressBar1.Maximum = filenum - 1
        ProgressBar1.Visible = True
        ListBox1.Items.Clear()
        ClearInterface()

        For x = 1 To filenum
            ChDir(tempdirectory)
            ProgressBar1.Value = x - 1
            ProgressBar1.Update()
            Application.DoEvents()

            ' open file
            'Unzip epub to temp directory

            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                Try
                    'delete contents of temp directory
                    DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                Catch
                    wait(500)
                    'try again
                    DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                End Try
            Else
                MkDir(ebookdirectory)
            End If
            ChDir(ebookdirectory)

            Try
                Dim zip As ZipStorer
                zip = ZipStorer.Open(Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + ComboBox3.Items(x - 1).ToString, FileAccess.Read)
                Dim dir = zip.ReadCentralDir()
                Dim item As ZipStorer.ZipFileEntry
                For Each item In dir
                    zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
                Next
                zip.Close()
            Catch ex1 As Exception
                Console.Error.WriteLine("exception: {0}", ex1.ToString)
                DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "The ebook " + ComboBox3.Items(x - 1) + " cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Exit Sub
            End Try

            'Search for .opf file
            searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

            'Open .opf file into RichTextBox
            If searchResults.Length < 1 Then
                DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "The ebook " + ComboBox3.Items(x - 1) + " is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                Return
            Else
                opffile = searchResults(0)
                If InStr(opffile, "_MACOSX") Then
                    If searchResults.Length > 1 Then
                        opffile = searchResults(1)
                    End If
                End If
                opfdirectory = Path.GetDirectoryName(opffile)
                RichTextBox1.Text = LoadUnicodeFile(opffile)
            End If

            If Form8.RadioButton1.Checked Then
                'Process .opf file to determine EPUB version
                Dim opffiletext As String
                Dim packagepos, endpos, versionpos As Integer
                opffiletext = LoadUnicodeFile(opffile)
                packagepos = InStr(opffiletext, "<package")
                If packagepos <> 0 Then
                    endpos = InStr(packagepos, opffiletext, ">")
                    versionpos = InStr(packagepos, opffiletext, "version=")
                    If versionpos < endpos Then
                        versioninfo = Mid(opffiletext, versionpos + 9, 3)
                    End If
                End If

                If versioninfo = "3.0" Then
                    ListBox1.Items.Add(Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + ComboBox3.Items(x - 1))
                    Button10.Enabled = True
                    Button32.Enabled = True
                    Button41.Enabled = True
                End If
            ElseIf Form8.RadioButton2.Checked Then
                'Look for << and >> in OPF file
                Dim opffiletext As String
                opffiletext = LoadUnicodeFile(opffile)
                ' delete whitespace
                opffiletext = opffiletext.Replace(Chr(13), "")
                opffiletext = opffiletext.Replace(Chr(10), "")
                opffiletext = opffiletext.Replace(Chr(9), "")
                While (opffiletext.Contains("> "))
                    opffiletext = opffiletext.Replace("> ", ">")
                End While
                While (opffiletext.Contains("< "))
                    opffiletext = opffiletext.Replace("< ", "<")
                End While

                If ((opffiletext.Contains(">>")) Or (opffiletext.Contains("<<"))) Then
                    ListBox1.Items.Add(Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + ComboBox3.Items(x - 1))
                    If Form8.CheckBox1.Checked Then
                        'fix file
                        opffiletext = opffiletext.Replace(">>", ">")
                        opffiletext = opffiletext.Replace("<<", "<")
                        opffiletext = Regularise(opffiletext)

                        'save opf file
                        SaveUnicodeFile(opffile, opffiletext)

                        'save EPUB file
                        Dim EPUBfilename As String
                        EPUBfilename = Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + ComboBox3.Items(x - 1)
                        Dim fi As New FileInfo(EPUBfilename)

                        'Zip temp directory (after deleting original file)
                        fi.Delete()

                        'Delete mimetype file
                        Dim temporarydirectory = CurDir()
                        ChDir(ebookdirectory)
                        IO.File.Delete("mimetype")
                        ChDir(temporarydirectory)

                        Dim zip As ZipStorer
                        zip = ZipStorer.Create(EPUBfilename, "")
                        Dim mimetype As New MemoryStream(System.Text.Encoding.UTF8.GetBytes("application/epub+zip"))
                        zip.AddStream(ZipStorer.Compression.Store, "mimetype", mimetype, DateTime.Now, "")
                        mimetype.Close()
                        Dim dir = Directory.GetDirectories(ebookdirectory)
                        Dim item As String
                        For Each item In dir
                            zip.AddDirectory(ZipStorer.Compression.Deflate, item, "", "")
                        Next
                        Dim files = Directory.GetFiles(ebookdirectory)
                        For Each item In files
                            zip.AddFile(ZipStorer.Compression.Deflate, item, Path.GetFileName(item), "")
                        Next
                        zip.Close()
                    End If
                    Button10.Enabled = True
                    Button32.Enabled = True
                    Button41.Enabled = True
                End If
            ElseIf Form8.RadioButton3.Checked Then
                'Look for text in OPF file
                Dim opffiletext As String
                opffiletext = LoadUnicodeFile(opffile)
                If (opffiletext.Contains(Form8.TextBox1.Text)) Then
                    ListBox1.Items.Add(Path.GetDirectoryName(OpenFileDialog1.FileName) + "\" + ComboBox3.Items(x - 1))
                    Button10.Enabled = True
                    Button32.Enabled = True
                    Button41.Enabled = True
                End If
            End If

            Application.DoEvents()

        Next
        ProgressBar1.Value = 0
        ProgressBar1.Update()
        ProgressBar1.Visible = False
        projectchanged = False
        Button3.Enabled = False
        ClearInterface()
        DialogResult = MsgBox("All done!", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")

        If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
            'delete contents of temp directory
            DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
        End If

        projectchanged = False
        CaptionString = "EPUB Metadata Editor"
        Me.Text = CaptionString
    End Sub
    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        Process.Start(ebookdirectory)
    End Sub

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        If ListBox1.SelectedIndex >= 0 Then
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            ClearInterface()
        End If
    End Sub

    Private Sub LinkLabel10_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel10.LinkClicked
        Form9.ShowDialog()
    End Sub

    Private Sub LinkLabel11_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel11.LinkClicked
        Form10.ShowDialog()
    End Sub
    Public Function countString(ByVal inputString As String, ByVal stringToBeSearchedInsideTheInputString As String) As Integer
        Return System.Text.RegularExpressions.Regex.Split(inputString, stringToBeSearchedInsideTheInputString).Length - 1
    End Function

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        Dim filenum, x As Integer
        Dim metadatafile As String
        Dim filesfailed As Boolean
        Dim filelist As String

        ' get current template from ini file
        Dim viewerfilename, inidirectory, inifilename As String
        inidirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\EPubMetadataEditor"
        inifilename = inidirectory + "\EPubMetadataEditor.ini"
        If System.IO.File.Exists(inifilename) = False Then
            If System.IO.Directory.Exists(inidirectory) = False Then
                System.IO.Directory.CreateDirectory(inidirectory)
            End If
            Dim fs As New FileStream(inifilename, FileMode.Create, FileAccess.Write)
            Dim s As New StreamWriter(fs)

            ' look for ini file in old location
            Dim tempinifile = Application.StartupPath() + "\EPubMetadataEditor.ini"
            If System.IO.File.Exists(tempinifile) = True Then
                Dim tempobjIniFile As New IniFile(tempinifile)
                viewerfilename = tempobjIniFile.GetString("Viewer", "Path", "(none)")
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + viewerfilename + Chr(34))
            Else
                s.WriteLine("[Viewer]")
                s.WriteLine("Path=" + Chr(34) + "(none)" + Chr(34))
            End If
            s.WriteLine("[Extractor]")
            s.WriteLine("Template=" + Chr(34) + "(none)" + Chr(34))
            s.Close()
        End If

        Dim objIniFile As New IniFile(inifilename)
        Dim template = objIniFile.GetString("Extractor", "Template", "(none)")
        Form11.ComboBox1.Items.Clear()
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        Else
            Form11.ComboBox1.Items.Add("")
            Form11.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Extractor", "Template1", "(none)")
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Extractor", "Template2", "(none)")
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Extractor", "Template3", "(none)")
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Extractor", "Template4", "(none)")
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        End If
        template = objIniFile.GetString("Extractor", "Template5", "(none)")
        If template <> "(none)" Then
            Form11.ComboBox1.Items.Add(template)
            Form11.ComboBox1.SelectedIndex = 0
        End If

        If Form11.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ' check to see if current scheme has changed
            template = objIniFile.GetString("Extractor", "Template", "(none)")
            If Form11.ComboBox1.Text <> template Then
                ' update ini file, cycling through existing history, adding new template as most recent
                template = objIniFile.GetString("Extractor", "Template4", "(none)")
                objIniFile.WriteString("Extractor", "Template5", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Extractor", "Template3", "(none)")
                objIniFile.WriteString("Extractor", "Template4", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Extractor", "Template2", "(none)")
                objIniFile.WriteString("Extractor", "Template3", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Extractor", "Template1", "(none)")
                objIniFile.WriteString("Extractor", "Template2", Chr(34) + template + Chr(34))
                template = objIniFile.GetString("Extractor", "Template", "(none)")
                objIniFile.WriteString("Extractor", "Template1", Chr(34) + template + Chr(34))
                objIniFile.WriteString("Extractor", "Template", Chr(34) + Form11.ComboBox1.Text + Chr(34))
            End If

            ClearInterface()
            tempdirectory = System.IO.Path.GetTempPath
            ebookdirectory = tempdirectory + "EPUB"

            filenum = ListBox1.Items.Count
            ProgressBar1.Maximum = filenum - 1
            ProgressBar1.Visible = True
            filesfailed = False
            filelist = "The extract procedure failed on the following files:"

            For x = 1 To filenum
                ChDir(tempdirectory)
                ProgressBar1.Value = x - 1
                ProgressBar1.Update()
                Application.DoEvents()

                ' open file
                'Unzip epub to temp directory

                If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                    Try
                        'delete contents of temp directory
                        DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                    Catch
                        wait(500)
                        'try again
                        DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
                    End Try
                Else
                    MkDir(ebookdirectory)
                End If
                ChDir(ebookdirectory)

                Try
                    Dim zip As ZipStorer
                    zip = ZipStorer.Open(ListBox1.Items(x - 1).ToString, FileAccess.Read)
                    Dim dir = zip.ReadCentralDir()
                    Dim item As ZipStorer.ZipFileEntry
                    For Each item In dir
                        zip.ExtractFile(item, ebookdirectory + "\" + item.FilenameInZip)
                    Next
                    zip.Close()
                Catch ex1 As Exception
                    Console.Error.WriteLine("exception: {0}", ex1.ToString)
                    DialogResult = MsgBox("ERROR: Problem with unzipping file." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " cannot be opened by the ZIP library used by EPUB Metadata Editor.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Exit Sub
                End Try

                'Search for .opf file
                searchResults = Directory.GetFiles(ebookdirectory, "*.opf", SearchOption.AllDirectories)

                'Open .opf file into RichTextBox
                If searchResults.Length < 1 Then
                    DialogResult = MsgBox("ERROR: Metadata not found." + Chr(10) + "The ebook " + ListBox1.Items(x - 1) + " is malformed.", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
                    Return
                Else
                    opffile = searchResults(0)
                    If InStr(opffile, "_MACOSX") Then
                        If searchResults.Length > 1 Then
                            opffile = searchResults(1)
                        End If
                    End If
                    opfdirectory = Path.GetDirectoryName(opffile)
                    RichTextBox1.Text = LoadUnicodeFile(opffile)
                End If

                'Extract metadata into textboxes
                metadatafile = LoadUnicodeFile(opffile)

                ' No need to extract cover
                ExtractMetadata(metadatafile, False)

                WebBrowser1.Visible = False

                ' extract the metadata from the filename
                Dim currposTemplate, currposFilename, endpos, endMetadata, slashloc As Integer
                Dim currentField, currentSearchText, currentFilename, currentMetadata As String
                Dim foundfield As Boolean
                currentSearchText = ""
                currentMetadata = ""
                currentField = ""
                currposTemplate = 0
                currposFilename = 1
                foundfield = False

                Try
                    ' get location of last "\"
                    slashloc = InStrRev(ListBox1.Items(x - 1), "\")

                    ' include folders if Template contains "\" characters
                    Dim numslash As Integer = Len(Form11.ComboBox1.Text) - Len(Strings.Replace(Form11.ComboBox1.Text, "\", ""))
                    While numslash > 0
                        slashloc = InStrRev(ListBox1.Items(x - 1), "\", slashloc - 1)
                        numslash = numslash - 1
                    End While

                    ' get filename
                    currentFilename = Mid(ListBox1.Items(x - 1), slashloc + 1)
                    currentFilename = Mid(currentFilename, 1, Len(currentFilename) - 5)

                    ' parse template
                    While (currposTemplate < Len(Form11.ComboBox1.Text))
                        currposTemplate = currposTemplate + 1

                        ' look for field marker
                        If (Mid(Form11.ComboBox1.Text, currposTemplate, 1) = "%") Then
                            If (Mid(Form11.ComboBox1.Text, currposTemplate + 1, 1) = "%") Then
                                ' found '%%' (replace with '%')
                                currentSearchText = currentSearchText + "%"
                                currposTemplate = currposTemplate + 1
                            Else
                                If foundfield Then
                                    ' new field found so use currentSearchText to extract metadata for currentField
                                    endMetadata = InStr(currposFilename, currentFilename, currentSearchText) - 1
                                    currentMetadata = Mid(currentFilename, currposFilename, endMetadata - currposFilename + 1)

                                    ' update metadata from filename
                                    If currentField = "Creator" Then
                                        TextBox2.Text = currentMetadata
                                    ElseIf currentField = "CreatorFileAs" Then
                                        TextBox12.Text = currentMetadata
                                    ElseIf currentField = "Title" Then
                                        TextBox1.Text = currentMetadata
                                    ElseIf currentField = "TitleFileAs" Then
                                        TextBox16.Text = currentMetadata
                                    ElseIf currentField = "Series" Then
                                        TextBox15.Text = currentMetadata
                                    ElseIf currentField = "SeriesIndex" Then
                                        TextBox14.Text = currentMetadata
                                    ElseIf currentField = "Date" Then
                                        TextBox6.Text = currentMetadata
                                    End If
                                    currposFilename = endMetadata + Len(currentSearchText) + 1
                                    currentSearchText = ""
                                End If

                                ' look for end field marker
                                endpos = InStr(currposTemplate + 1, Form11.ComboBox1.Text, "%")
                                If (endpos <> 0) Then
                                    ' end field marker found
                                    currentField = Mid(Form11.ComboBox1.Text, currposTemplate + 1, endpos - currposTemplate - 1)
                                    foundfield = True
                                    currposTemplate = endpos
                                End If
                            End If
                        Else
                            currentSearchText = currentSearchText + Mid(Form11.ComboBox1.Text, currposTemplate, 1)
                        End If
                    End While
                    ' Get metadata from filename for last field
                    currentMetadata = Mid(currentFilename, currposFilename)

                    ' update metadata for last field
                    If currentField = "Creator" Then
                        TextBox2.Text = currentMetadata
                    ElseIf currentField = "CreatorFileAs" Then
                        TextBox12.Text = currentMetadata
                    ElseIf currentField = "Title" Then
                        TextBox1.Text = currentMetadata
                    ElseIf currentField = "TitleFileAs" Then
                        TextBox16.Text = currentMetadata
                    ElseIf currentField = "Series" Then
                        TextBox15.Text = currentMetadata
                    ElseIf currentField = "SeriesIndex" Then
                        TextBox14.Text = currentMetadata
                    ElseIf currentField = "Date" Then
                        TextBox6.Text = currentMetadata
                    End If

                    Application.DoEvents()

                    ' Only save file if we made it this far
                    SaveEpub(ListBox1.Items(x - 1), False)
                Catch
                    ' Extract failed, so add filename to list
                    filesfailed = True
                    filelist = filelist + Chr(10) + ListBox1.Items(x - 1)
                End Try

                ClearInterface()

            Next
            ProgressBar1.Value = 0
            ProgressBar1.Update()
            ProgressBar1.Visible = False
            projectchanged = False
            Button3.Enabled = False
            ClearInterface()
            If filesfailed Then
                DialogResult = MsgBox("All done!" + Chr(10) + filelist, MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            Else
                DialogResult = MsgBox("All done!", MsgBoxStyle.OkOnly, "EPUB Metadata Editor")
            End If


            If (My.Computer.FileSystem.DirectoryExists(ebookdirectory)) Then
                'delete contents of temp directory
                DeleteDirContents(New IO.DirectoryInfo(ebookdirectory))
            End If

            projectchanged = False
            CaptionString = "EPUB Metadata Editor"
            Me.Text = CaptionString
        End If
    End Sub
End Class
