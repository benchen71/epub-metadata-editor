' ZipStorer, by Jaime Olivares
' Website: http://github.com/jaime-olivares/zipstorer
' Version: 3.5.0 (May 20, 2019)
' Code updated by Ben Chenoweth

Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.IO.Compression

' <summary>
' Unique class for compression/decompression file. Represents a Zip file.
' </summary>
Public Class ZipStorer
    Implements IDisposable

    ' <summary>
    ' Compression method enumeration
    ' </summary>
    Public Enum Compression As UShort
        ' <summary>Uncompressed storage</summary> 
        Store = 0
        ' <summary>Deflate compression method</summary>
        Deflate = 8
    End Enum

    ' <summary>
    ' Represents an entry in Zip file directory
    ' </summary>
    Public Structure ZipFileEntry

        ' <summary>Compression method</summary>
        Public Method As Compression
        ' <summary>Full path and filename as stored in Zip</summary>
        Public FilenameInZip As String
        ' <summary>Original file size</summary>
        Public FileSize As UInteger
        ' <summary>Compressed file size</summary>
        Public CompressedSize As UInteger
        ' <summary>Offset of header information inside Zip storage</summary>
        Public HeaderOffset As UInteger
        ' <summary>Offset of file inside Zip storage</summary>
        Public FileOffset As UInteger
        ' <summary>Size of header information</summary>
        Public HeaderSize As UInteger
        ' <summary>32-bit checksum of entire file</summary>
        Public Crc32 As UInteger
        ' <summary>Last modification time of file</summary>
        Public ModifyTime As DateTime
        ' <summary>Creation time of file</summary>
        Public CreationTime As DateTime
        ' <summary>Last access time of file</summary>
        Public AccessTime As DateTime
        ' <summary>User comment for file</summary>
        Public Comment As String
        ' <summary>True if UTF8 encoding for filename and comments, false if default (CP 437)</summary>
        Public EncodeUTF8 As Boolean

        ' <summary>Overriden method</summary>
        ' <returns>Filename in Zip</returns>
        Public Overrides Function ToString() As String
            Return Me.FilenameInZip
        End Function
    End Structure

#Region "Public fields"
    ' <summary>True if UTF8 encoding for filename and comments, false if default (CP 437)</summary>
    Public EncodeUTF8 As Boolean = False
    ' <summary>Force deflate algotithm even if it inflates the stored file. Off by default.</summary>
    Public ForceDeflating As Boolean = False
#End Region

#Region "Private fields"
    ' List of files to store
    Private Files As New List(Of ZipFileEntry)()
    ' Filename of storage file
    Private FileName As String
    ' Stream object of storage file
    Private ZipFileStream As Stream
    ' General comment
    Private Comment As String = ""
    ' Central dir image
    Private CentralDirImage As Byte() = Nothing
    ' Existing files in zip
    Private ExistingFiles As UShort = 0
    ' File access for Open method
    Private Access As FileAccess
    ' Leave the stream open after the ZipStorer object is disposed
    Private leaveOpen As Boolean
    ' Static CRC32 Table
    Private Shared CrcTable As UInt32() = Nothing
    ' Default filename encoder
    Private Shared DefaultEncoding As Encoding = Encoding.GetEncoding(437)
#End Region

#Region "Public methods"
    ' Static constructor. Just invoked once in order to create the CRC32 lookup table.
    Shared Sub New()
        ' Generate CRC32 table
        CrcTable = New UInt32(255) {}
        For i As Integer = 0 To CrcTable.Length - 1
            Dim c As UInt32 = CType(i, UInt32)
            For j As Integer = 0 To 7
                If (c And 1) <> 0 Then
                    c = 3988292384 Xor (c >> 1)
                Else
                    c >>= 1
                End If
            Next
            CrcTable(i) = c
        Next
    End Sub

    ' <summary>
    ' Method to create a new storage file
    ' </summary>
    ' <param name="_filename">Full path of Zip file to create</param>
    ' <param name="_comment">General comment for Zip file</param>
    ' <param name="_leaveOpen">true to leave the stream open after the ZipStorer object is disposed; otherwise, false (default).</param>
    ' <returns>A valid ZipStorer object</returns>
    Public Shared Function Create(ByVal _filename As String, ByVal _comment As String) As ZipStorer
        Dim stream As Stream = New FileStream(_filename, FileMode.Create, FileAccess.ReadWrite)
        Dim zip As ZipStorer = Create(stream, _comment)

        zip.Comment = _comment
        zip.FileName = _filename

        Return zip
    End Function

    ' <summary>
    ' Method to create a new zip storage in a stream
    ' </summary>
    ' <param name="_stream"></param>
    ' <param name="_comment"></param>
    ' <returns>A valid ZipStorer object</returns>
    Public Shared Function Create(ByVal _stream As Stream, ByVal _comment As String, Optional ByVal _leaveOpen As Boolean = False) As ZipStorer
        Dim zip As New ZipStorer()

        zip.Comment = _comment
        zip.ZipFileStream = _stream
        zip.Access = FileAccess.Write
        zip.leaveOpen = _leaveOpen

        Return zip
    End Function

    ' <summary>
    ' Method to open an existing storage file
    ' </summary>
    ' <param name="_filename">Full path of Zip file to open</param>
    ' <param name="_access">File access mode as used in FileStream constructor</param>
    ' <returns>A valid ZipStorer object</returns>
    Public Shared Function Open(ByVal _filename As String, ByVal _access As FileAccess) As ZipStorer
        Dim stream As Stream = CType(New FileStream(_filename, FileMode.Open, If(_access = FileAccess.Read, FileAccess.Read, FileAccess.ReadWrite)), Stream)
        Dim zip As ZipStorer = Open(stream, _access)

        zip.FileName = _filename

        Return zip
    End Function

    ' <summary>
    ' Method to open an existing storage from stream
    ' </summary>
    ' <param name="_stream">Already opened stream with zip contents</param>
    ' <param name="_access">File access mode for stream operations</param>
    ' <returns>A valid ZipStorer object</returns>
    Public Shared Function Open(ByVal _stream As Stream, ByVal _access As FileAccess, Optional ByVal _leaveOpen As Boolean = False) As ZipStorer
        If Not _stream.CanSeek AndAlso _access <> FileAccess.Read Then Throw New InvalidOperationException("Stream cannot seek")

        Dim zip As New ZipStorer()
        'zip.FileName = _filename
        zip.ZipFileStream = _stream
        zip.Access = _access
        zip.leaveOpen = _leaveOpen

        If zip.ReadFileInfo() Then Return zip

        If (Not _leaveOpen) Then zip.Close()

        Throw New System.IO.InvalidDataException()
    End Function

    ' <summary>
    ' Add full contents of a file into the Zip storage
    ' </summary>
    ' <param name="_method">Compression method</param>
    ' <param name="_pathname">Full path of file to add to Zip storage</param>
    ' <param name="_filenameInZip">Filename and path as desired in Zip directory</param>
    ' <param name="_comment">Comment for stored file</param>        
    Public Sub AddFile(ByVal _method As Compression, ByVal _pathname As String, ByVal _filenameInZip As String, ByVal _comment As String)
        If Access = FileAccess.Read Then Throw New InvalidOperationException("Writing is not alowed")

        Dim stream As New FileStream(_pathname, FileMode.Open, FileAccess.Read)
        AddStream(_method, _filenameInZip, stream, File.GetLastWriteTime(_pathname), _comment)
        stream.Close()
    End Sub

    ''' <summary>
    ''' Add full contents of a directory into the Zip storage
    ''' </summary>
    ''' <param name="_method">Compression method</param>
    ''' <param name="_pathname">Full path of directory to add to Zip storage</param>
    ''' <param name="_pathnameInZip">Path name as desired in Zip directory</param>
    ''' <param name="_comment">Comment for stored directory</param>
    Public Sub AddDirectory(ByVal _method As Compression, ByVal _pathname As String, ByVal _pathnameInZip As String, ByVal _comment As String)
        If Access = FileAccess.Read Then
            Throw New InvalidOperationException("Writing is not allowed")
        End If

        Dim foldername As String
        Dim pos As Integer = _pathname.LastIndexOf(Path.DirectorySeparatorChar)
        If pos >= 0 Then
            foldername = _pathname.Remove(0, pos + 1)
        Else
            foldername = _pathname
        End If

        If _pathnameInZip IsNot Nothing AndAlso _pathnameInZip <> "" Then
            foldername = _pathnameInZip & foldername
        End If

        If Not foldername.EndsWith("/") Then
            foldername = foldername & "/"
        End If

        ' Process the list of files found in the directory.
        Dim fileEntries As String() = Directory.GetFiles(_pathname)
        For Each fileName As String In fileEntries
            AddFile(_method, fileName, foldername & Path.GetFileName(fileName), "")
        Next

        ' Recurse into subdirectories of this directory.
        Dim subdirectoryEntries As String() = Directory.GetDirectories(_pathname)
        For Each subdirectory As String In subdirectoryEntries
            AddDirectory(_method, subdirectory, foldername, "")
        Next
    End Sub

    ' <summary>
    ' Add full contents of a stream into the Zip storage
    ' </summary>
    ' <param name="_method">Compression method</param>
    ' <param name="_filenameInZip">Filename and path as desired in Zip directory</param>
    ' <param name="_source">Stream object containing the data to store in Zip</param>
    ' <param name="_modTime">Modification time of the data to store</param>
    ' <param name="_comment">Comment for stored file</param>
    Public Sub AddStream(ByVal _method As Compression, ByVal _filenameInZip As String, ByVal _source As Stream, ByVal _modTime As DateTime, ByVal _comment As String)
        If Access = FileAccess.Read Then Throw New InvalidOperationException("Writing is not alowed")

        Dim offset As Long
        If Me.Files.Count = 0 Then
            offset = 0
        Else
            Dim last As ZipFileEntry = Me.Files(Me.Files.Count - 1)
            offset = last.HeaderOffset + last.HeaderSize
        End If

        ' Prepare the fileinfo
        Dim zfe As New ZipFileEntry()
        zfe.Method = _method
        zfe.EncodeUTF8 = Me.EncodeUTF8
        zfe.FilenameInZip = NormalizedFilename(_filenameInZip)
        zfe.Comment = If(_comment Is Nothing, "", _comment)

        ' Even though we write the header now, it will have to be rewritten, since we don't know compressed size or crc.
        zfe.Crc32 = 0  ' to be updated later
        zfe.HeaderOffset = CType(Me.ZipFileStream.Position, UInteger)  ' offset within file of the start of this local record
        zfe.CreationTime = _modTime
        zfe.ModifyTime = _modTime
        zfe.AccessTime = _modTime

        ' Write local header
        WriteLocalHeader(zfe)
        zfe.FileOffset = CType(Me.ZipFileStream.Position, UInteger)

        ' Write file to zip (store)
        Store(zfe, _source)
        _source.Close()

        Me.UpdateCrcAndSizes(zfe)

        Files.Add(zfe)
    End Sub

    ' <summary>
    ' Updates central directory (if pertinent) and close the Zip storage
    ' </summary>
    ' <remarks>This is a required step, unless automatic dispose is used</remarks>
    Public Sub Close()
        If Me.Access <> FileAccess.Read AndAlso Me.ZipFileStream IsNot Nothing Then
            Dim centralOffset As UInteger = CType(Me.ZipFileStream.Position, UInteger)
            Dim centralSize As UInteger = 0

            If Me.CentralDirImage IsNot Nothing Then Me.ZipFileStream.Write(CentralDirImage, 0, CentralDirImage.Length)

            For i As Integer = 0 To Files.Count - 1
                Dim pos As Long = Me.ZipFileStream.Position
                Me.WriteCentralDirRecord(Files(i))
                centralSize += CType(Me.ZipFileStream.Position - pos, UInteger)
            Next

            If (Me.CentralDirImage IsNot Nothing) Then
                Me.WriteEndRecord(centralSize + CType(CentralDirImage.Length, UInteger), centralOffset)
            Else
                Me.WriteEndRecord(centralSize, centralOffset)
            End If
        End If

        If ((Me.ZipFileStream IsNot Nothing) And (Not Me.leaveOpen)) Then
            Me.ZipFileStream.Flush()
            Me.ZipFileStream.Dispose()
            Me.ZipFileStream = Nothing
        End If
    End Sub

    ' <summary>
    ' Read all the file records in the central directory 
    ' </summary>
    ' <returns>List of all entries in directory</returns>
    Public Function ReadCentralDir() As List(Of ZipFileEntry)
        If Me.CentralDirImage Is Nothing Then Throw New InvalidOperationException("Central directory currently does not exist")

        Dim result As New List(Of ZipFileEntry)
        Dim pointer As Integer = 0

        While pointer < Me.CentralDirImage.Length - 3 ' Need 4 bytes left for ToUInt32
            Dim signature As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer)
            If signature <> &H2014B50UI Then Exit While

            Dim encodeUTF8 As Boolean = (BitConverter.ToUInt16(CentralDirImage, pointer + 8) And &H800UI) <> 0
            Dim method As UShort = BitConverter.ToUInt16(CentralDirImage, pointer + 10)
            Dim modifyTime As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer + 12)
            Dim crc32 As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer + 16)
            Dim comprSize As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer + 20)
            Dim fileSize As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer + 24)
            Dim filenameSize As UShort = BitConverter.ToUInt16(CentralDirImage, pointer + 28)
            Dim extraSize As UShort = BitConverter.ToUInt16(CentralDirImage, pointer + 30)
            Dim commentSize As UShort = BitConverter.ToUInt16(CentralDirImage, pointer + 32)
            Dim headerOffset As UInteger = BitConverter.ToUInt32(CentralDirImage, pointer + 42)
            Dim headerSize As UInteger = CType(46 + filenameSize + extraSize + commentSize, UInteger)

            Dim encoder As Encoding = If(encodeUTF8, Encoding.UTF8, DefaultEncoding)

            Dim zfe As New ZipFileEntry()
            zfe.Method = CType(method, Compression)
            zfe.FilenameInZip = encoder.GetString(CentralDirImage, pointer + 46, filenameSize)
            zfe.FileOffset = GetFileOffset(headerOffset)
            zfe.FileSize = fileSize
            zfe.CompressedSize = comprSize
            zfe.HeaderOffset = headerOffset
            zfe.HeaderSize = headerSize
            zfe.Crc32 = crc32
            zfe.ModifyTime = DosTimeToDateTime(modifyTime)
            zfe.CreationTime = zfe.ModifyTime
            zfe.AccessTime = DateTime.Now

            If commentSize > 0 Then zfe.Comment = encoder.GetString(CentralDirImage, pointer + 46 + filenameSize + extraSize, commentSize)
            If (extraSize > 0) Then ReadExtraInfo(CentralDirImage, pointer + 46 + filenameSize, zfe)

            result.Add(zfe)
            pointer += (46 + filenameSize + extraSize + commentSize)
        End While

        Return result
    End Function

    ' <summary>
    ' Copy the contents of a stored file into a physical file
    ' </summary>
    ' <param name="_zfe">Entry information of file to extract</param>
    ' <param name="_filename">Name of file to store uncompressed data</param>
    ' <returns>True if success, false if not.</returns>
    ' <remarks>Unique compression methods are Store and Deflate</remarks>
    Public Function ExtractFile(ByVal _zfe As ZipFileEntry, ByVal _filename As String) As Boolean

        ' Make sure the parent directory exist
        Dim path As String = System.IO.Path.GetDirectoryName(_filename)

        If Not Directory.Exists(path) Then Directory.CreateDirectory(path)
        ' Check it is directory. If so, do nothing
        If (Directory.Exists(_filename)) Then Return True

        Dim output As Stream = New FileStream(_filename, FileMode.Create, FileAccess.Write)
        Dim result As Boolean = ExtractFile(_zfe, output)
        If result Then output.Close()

        File.SetCreationTime(_filename, _zfe.CreationTime)
        File.SetLastWriteTime(_filename, _zfe.ModifyTime)
        File.SetLastAccessTime(_filename, _zfe.AccessTime)

        Return result
    End Function

    ' <summary>
    ' Copy the contents of a stored file into an opened stream
    ' </summary>
    ' <param name="_zfe">Entry information of file to extract</param>
    ' <param name="_stream">Stream to store the uncompressed data</param>
    ' <returns>True if success, false if not.</returns>
    ' <remarks>Unique compression methods are Store and Deflate</remarks>
    Public Function ExtractFile(ByVal _zfe As ZipFileEntry, ByVal _stream As Stream) As Boolean
        If Not _stream.CanWrite Then Throw New InvalidOperationException("Stream cannot be written")

        ' check signature
        Dim signature(3) As Byte
        Me.ZipFileStream.Seek(_zfe.HeaderOffset, SeekOrigin.Begin)
        Me.ZipFileStream.Read(signature, 0, 4)
        If BitConverter.ToUInt32(signature, 0) <> &H4034B50UI Then Return False

        ' Select input stream for inflating or just reading
        Dim inStream As Stream
        If _zfe.Method = Compression.Store Then
            inStream = Me.ZipFileStream
        ElseIf _zfe.Method = Compression.Deflate Then
            inStream = New DeflateStream(Me.ZipFileStream, CompressionMode.Decompress, True)
        Else
            Return False
        End If

        ' Buffered copy
        Dim buffer(16383) As Byte
        Me.ZipFileStream.Seek(_zfe.FileOffset, SeekOrigin.Begin)
        Dim bytesPending As UInteger = _zfe.FileSize
        While bytesPending > 0
            Dim bytesRead As Integer = inStream.Read(buffer, 0, CType(Math.Min(bytesPending, buffer.Length), Integer))
            _stream.Write(buffer, 0, bytesRead)
            bytesPending -= CType(bytesRead, UInteger)
        End While
        _stream.Flush()

        If _zfe.Method = Compression.Deflate Then inStream.Dispose()
        Return True
    End Function

    ' <summary>
    ' Removes one of many files in storage. It creates a new Zip file.
    ' </summary>
    ' <param name="_zip">Reference to the current Zip object</param>
    ' <param name="_zfes">List of Entries to remove from storage</param>
    ' <returns>True if success, false if not</returns>
    ' <remarks>This method only works for storage of type FileStream</remarks>
    Public Shared Function RemoveEntries(ByRef _zip As ZipStorer, ByVal _zfes As List(Of ZipFileEntry)) As Boolean
        If Not TypeOf _zip.ZipFileStream Is FileStream Then Throw New InvalidOperationException("RemoveEntries is allowed just over streams of type FileStream")

        'Get full list of entries
        Dim fullList As List(Of ZipFileEntry) = _zip.ReadCentralDir()

        'In order to delete we need to create a copy of the zip file excluding the selected items
        Dim tempZipName As String = Path.GetTempFileName()
        Dim tempEntryName As String = Path.GetTempFileName()

        Try
            Dim tempZip As ZipStorer = ZipStorer.Create(tempZipName, String.Empty)

            For Each zfe As ZipFileEntry In fullList
                If Not _zfes.Contains(zfe) Then
                    If _zip.ExtractFile(zfe, tempEntryName) Then
                        tempZip.AddFile(zfe.Method, tempEntryName, zfe.FilenameInZip, zfe.Comment)
                    End If
                End If
            Next
            _zip.Close()
            tempZip.Close()

            File.Delete(_zip.FileName)
            File.Move(tempZipName, _zip.FileName)

            _zip = ZipStorer.Open(_zip.FileName, _zip.Access)

        Catch ex As Exception
            Return False
        Finally
            If (File.Exists(tempZipName)) Then File.Delete(tempZipName)
            If (File.Exists(tempEntryName)) Then File.Delete(tempEntryName)
        End Try

        Return True
    End Function
#End Region

#Region "'Private methods"
    ' Calculate the file offset by reading the corresponding local header
    Private Function GetFileOffset(ByVal _headerOffset As UInteger) As UInteger
        Dim buffer(1) As Byte

        Me.ZipFileStream.Seek(_headerOffset + 26, SeekOrigin.Begin)
        Me.ZipFileStream.Read(buffer, 0, 2)
        Dim filenameSize As UShort = BitConverter.ToUInt16(buffer, 0)
        Me.ZipFileStream.Read(buffer, 0, 2)
        Dim extraSize As UShort = BitConverter.ToUInt16(buffer, 0)

        Return CType(30 + filenameSize + extraSize + _headerOffset, UInteger)
    End Function

    ' Local file header:
    '    local file header signature     4 bytes  (&H04034b50)
    '    version needed to extract       2 bytes
    '    general purpose bit flag        2 bytes
    '    compression method              2 bytes
    '    last mod file time              2 bytes
    '    last mod file date              2 bytes
    '    crc-32                          4 bytes
    '    compressed size                 4 bytes
    '    uncompressed size               4 bytes
    '    filename length                 2 bytes
    '    extra field length              2 bytes

    '    filename (variable size)
    '    extra field (variable size)

    Private Sub WriteLocalHeader(ByRef _zfe As ZipFileEntry)
        Dim pos As Long = Me.ZipFileStream.Position
        Dim encoder As Encoding = If(_zfe.EncodeUTF8, Encoding.UTF8, DefaultEncoding)
        Dim encodedFilename As Byte() = encoder.GetBytes(_zfe.FilenameInZip)
        Dim extraInfo = CreateExtraInfo(_zfe)

        Me.ZipFileStream.Write(New Byte() {80, 75, 3, 4, 20, 0}, 0, 6) ' No extra header
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(If(_zfe.EncodeUTF8, &H800, 0), UShort)), 0, 2) ' filename and comment encoding 
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(_zfe.Method, UShort)), 0, 2)  ' zipping method
        Me.ZipFileStream.Write(BitConverter.GetBytes(DateTimeToDosTime(_zfe.ModifyTime)), 0, 4) ' zipping date and time
        Me.ZipFileStream.Write(New Byte() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, 0, 12) ' unused CRC, un/compressed size, updated later
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(encodedFilename.Length, UShort)), 0, 2) ' filename length
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(extraInfo.Length, UShort)), 0, 2) ' extra length

        Me.ZipFileStream.Write(encodedFilename, 0, encodedFilename.Length)
        Me.ZipFileStream.Write(extraInfo, 0, extraInfo.Length)
        _zfe.HeaderSize = CType(Me.ZipFileStream.Position - pos, UInteger)
    End Sub

    Private Sub WriteCentralDirRecord(ByVal _zfe As ZipFileEntry)
        Dim encoder As Encoding = If(_zfe.EncodeUTF8, Encoding.UTF8, DefaultEncoding)
        Dim encodedFilename As Byte() = encoder.GetBytes(_zfe.FilenameInZip)
        Dim encodedComment As Byte() = encoder.GetBytes(_zfe.Comment)
        Dim extraInfo = CreateExtraInfo(_zfe)

        Me.ZipFileStream.Write(New Byte() {80, 75, 1, 2, 23, &HB, 20, 0}, 0, 8)
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(If(_zfe.EncodeUTF8, &H800, 0), UShort)), 0, 2) ' filename and comment encoding 
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(_zfe.Method, UShort)), 0, 2)  ' zipping method
        Me.ZipFileStream.Write(BitConverter.GetBytes(DateTimeToDosTime(_zfe.ModifyTime)), 0, 4)  ' zipping date and time
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.Crc32), 0, 4) ' file CRC
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.CompressedSize), 0, 4) ' compressed file size
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.FileSize), 0, 4) ' uncompressed file size
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(encodedFilename.Length, UShort)), 0, 2) ' Filename in zip
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(extraInfo.Length, UShort)), 0, 2) ' extra length
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(encodedComment.Length, UShort)), 0, 2)

        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(0, UShort)), 0, 2) ' disk=0
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(0, UShort)), 0, 2) ' file type: binary
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(0, UShort)), 0, 2) ' Internal file attributes
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(&H8100, UShort)), 0, 2) ' External file attributes (normal/readable)
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.HeaderOffset), 0, 4)  ' Offset of header

        Me.ZipFileStream.Write(encodedFilename, 0, encodedFilename.Length)
        Me.ZipFileStream.Write(extraInfo, 0, extraInfo.Length)
        Me.ZipFileStream.Write(encodedComment, 0, encodedComment.Length)
    End Sub

    ' End of central dir record:
    '    end of central dir signature    4 bytes  (&H06054b50)
    '    number of this disk             2 bytes
    '    number of the disk with the
    '    start of the central directory  2 bytes
    '    total number of entries in
    '    the central dir on this disk    2 bytes
    '    total number of entries in
    '    the central dir                 2 bytes
    '    size of the central directory   4 bytes
    '    offset of start of central
    '    directory with respect to
    '    the starting disk number        4 bytes
    '    zipfile comment length          2 bytes
    '    zipfile comment (variable size)

    Private Sub WriteEndRecord(ByVal _size As UInteger, ByVal _offset As UInteger)
        Dim encoder As Encoding = If(Me.EncodeUTF8, Encoding.UTF8, DefaultEncoding)
        Dim encodedComment As Byte() = encoder.GetBytes(Me.Comment)

        Me.ZipFileStream.Write(New Byte() {80, 75, 5, 6, 0, 0, 0, 0}, 0, 8)
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(Files.Count + ExistingFiles, UShort)), 0, 2)
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(Files.Count + ExistingFiles, UShort)), 0, 2)
        Me.ZipFileStream.Write(BitConverter.GetBytes(_size), 0, 4)
        Me.ZipFileStream.Write(BitConverter.GetBytes(_offset), 0, 4)
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(encodedComment.Length, UShort)), 0, 2)
        Me.ZipFileStream.Write(encodedComment, 0, encodedComment.Length)
    End Sub

    ' Copies all source file into storage file
    Private Sub Store(ByRef _zfe As ZipFileEntry, ByVal _source As Stream)
        Dim buffer(16383) As Byte
        Dim bytesRead As Integer
        Dim totalRead As UInteger = 0
        Dim outStream As Stream

        Dim posStart As Long = Me.ZipFileStream.Position
        Dim sourceStart As Long = _source.Position

        If _zfe.Method = Compression.Store Then
            outStream = Me.ZipFileStream
        Else
            outStream = New DeflateStream(Me.ZipFileStream, CompressionMode.Compress, True)
        End If

        _zfe.Crc32 = (0 Xor &HFFFFFFFFUI)

        Do
            bytesRead = _source.Read(buffer, 0, buffer.Length)
            totalRead += CType(bytesRead, UInteger)
            If bytesRead > 0 Then
                outStream.Write(buffer, 0, bytesRead)
                For i As UInteger = 0 To bytesRead - 1
                    _zfe.Crc32 = ZipStorer.CrcTable((_zfe.Crc32 Xor buffer(i)) And &HFF) Xor (_zfe.Crc32 >> 8)
                Next
            End If
        Loop While bytesRead = buffer.Length
        outStream.Flush()

        If _zfe.Method = Compression.Deflate Then outStream.Dispose()

        _zfe.Crc32 = _zfe.Crc32 Xor &HFFFFFFFFUI
        _zfe.FileSize = totalRead
        _zfe.CompressedSize = CType(Me.ZipFileStream.Position - posStart, UInteger)

        ' Verify for real compression
        If _zfe.Method = Compression.Deflate AndAlso Not Me.ForceDeflating AndAlso _source.CanSeek AndAlso _zfe.CompressedSize > _zfe.FileSize Then
            ' Start operation again with Store algorithm
            _zfe.Method = Compression.Store
            Me.ZipFileStream.Position = posStart
            Me.ZipFileStream.SetLength(posStart)
            _source.Position = sourceStart
            Me.Store(_zfe, _source)
        End If
    End Sub

    ' DOS Date and time:
    '    MS-DOS date. The date is a packed value with the following format. Bits Description 
    '        0-4 Day of the month (1–31) 
    '        5-8 Month (1 = January, 2 = February, and so on) 
    '        9-15 Year offset from 1980 (add 1980 to get actual year) 
    '    MS-DOS time. The time is a packed value with the following format. Bits Description 
    '        0-4 Second divided by 2 
    '        5-10 Minute (0–59) 
    '        11-15 Hour (0–23 on a 24-hour clock) 

    Private Function DateTimeToDosTime(ByVal _dt As DateTime) As UInteger
        Return CType((_dt.Second / 2) Or (_dt.Minute << 5) Or (_dt.Hour << 11) Or (_dt.Day << 16) Or (_dt.Month << 21) Or ((_dt.Year - 1980) << 25), UInteger)
    End Function

    Private Function CreateExtraInfo(ByVal _zfe As ZipFileEntry) As Byte()
        Dim buffer(36) As Byte

        BitConverter.GetBytes(CType(&HA, UShort)).CopyTo(buffer, 0) ' NTFS FileTime
        BitConverter.GetBytes(CType(32, UShort)).CopyTo(buffer, 2) ' Length
        BitConverter.GetBytes(CType(1, UShort)).CopyTo(buffer, 8) ' Tag 1
        BitConverter.GetBytes(CType(24, UShort)).CopyTo(buffer, 10) ' Size 1
        BitConverter.GetBytes(_zfe.ModifyTime.ToFileTime()).CopyTo(buffer, 12) ' MTime
        BitConverter.GetBytes(_zfe.AccessTime.ToFileTime()).CopyTo(buffer, 20) ' ATime
        BitConverter.GetBytes(_zfe.CreationTime.ToFileTime()).CopyTo(buffer, 28) ' CTime

        Return buffer
    End Function

    Private Sub ReadExtraInfo(ByVal buffer As Byte(), ByVal offset As Integer, ByVal _zfe As ZipFileEntry)

        If (Buffer.Length < 4) Then
            Return
        End If

        Dim pos As Integer = offset

        While (pos < buffer.Length - 4)
            Dim extraId As UInteger = BitConverter.ToUInt16(buffer, pos)
            Dim length As UInteger = BitConverter.ToUInt16(buffer, pos + 2)

            If (extraId = &HA) Then ' NTFS FileTime

                Dim tag As UInteger = BitConverter.ToUInt16(buffer, pos + 8)
                Dim size As UInteger = BitConverter.ToUInt16(buffer, pos + 10)

                If ((tag = 1) And (size = 24)) Then
                    _zfe.ModifyTime = DateTime.FromFileTime(BitConverter.ToInt64(buffer, pos + 12))
                    _zfe.AccessTime = DateTime.FromFileTime(BitConverter.ToInt64(buffer, pos + 20))
                    _zfe.CreationTime = DateTime.FromFileTime(BitConverter.ToInt64(buffer, pos + 28))
                End If
            End If

            pos = pos + CType(length, Integer) + 4
        End While
    End Sub

    Private Function DosTimeToDateTime(ByVal _dt As UInteger) As DateTime
        Dim year As Integer = CType(_dt >> 25, Integer) + 1980
        Dim month As Integer = CType(_dt >> 21, Integer) And 15
        Dim day As Integer = CType(_dt >> 16, Integer) And 31
        Dim hours As Integer = CType(_dt >> 11, Integer) And 31
        Dim minutes As Integer = CType(_dt >> 5, Integer) And 63
        Dim seconds As Integer = CType(_dt And 31, Integer) * 2

        If seconds >= 60 Then seconds = 59
        If ((year >= 2107) Or (month = 0) Or (day = 0)) Then Return DateTime.Now

        Return New DateTime(year, month, day, hours, minutes, seconds)
    End Function

    '     CRC32 algorithm
    '      The 'magic number' for the CRC is &Hdebb20e3.  
    '      The proper CRC pre and post conditioning
    '      is used, meaning that the CRC register is
    '      pre-conditioned with all ones (a starting value
    '      of &Hffffffff) and the value is post-conditioned by
    '      taking the one's complement of the CRC residual.
    '      If bit 3 of the general purpose flag is set, this
    '      field is set to zero in the local header and the correct
    '      value is put in the data descriptor and in the central
    '      directory.

    Private Sub UpdateCrcAndSizes(ByRef _zfe As ZipFileEntry)
        Dim lastPos As Long = Me.ZipFileStream.Position  ' remember position

        Me.ZipFileStream.Position = _zfe.HeaderOffset + 8
        Me.ZipFileStream.Write(BitConverter.GetBytes(CType(_zfe.Method, UShort)), 0, 2)  ' zipping method

        Me.ZipFileStream.Position = _zfe.HeaderOffset + 14
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.Crc32), 0, 4)  ' Update CRC
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.CompressedSize), 0, 4)  ' Compressed size
        Me.ZipFileStream.Write(BitConverter.GetBytes(_zfe.FileSize), 0, 4)  ' Uncompressed size

        Me.ZipFileStream.Position = lastPos  ' restore position
    End Sub

    ' Replaces backslashes with slashes to store in zip header
    Private Function NormalizedFilename(ByVal _filename As String) As String
        Dim filename As String = _filename.Replace("\"c, "/"c)
        Dim pos As Integer = filename.IndexOf(":"c)
        If pos >= 0 Then filename = filename.Remove(0, pos + 1)

        Return filename.Trim("/"c)
    End Function

    ' Reads the end-of-central-directory record
    Private Function ReadFileInfo() As Boolean

        If (Me.ZipFileStream.Length < 22) Then Return False

        Try
            Me.ZipFileStream.Seek(-17, SeekOrigin.End)
            Dim br As New BinaryReader(Me.ZipFileStream)

            Do

                Me.ZipFileStream.Seek(-5, SeekOrigin.Current)
                Dim sig As UInt32 = br.ReadUInt32()

                If sig = &H6054B50UI Then

                    Me.ZipFileStream.Seek(6, SeekOrigin.Current)

                    Dim entries As UInt16 = br.ReadUInt16()
                    Dim centralSize As Int32 = br.ReadInt32()
                    Dim centralDirOffset As UInt32 = br.ReadUInt32()
                    Dim commentSize As UInt16 = br.ReadUInt16()

                    ' check if comment field is the very last data in file
                    ' the following line has been commented out so that zip files with trailing data can be opened
                    'If (Me.ZipFileStream.Position + commentSize <> Me.ZipFileStream.Length) Then Return False

                    ' Copy entire central directory to a memory buffer
                    Me.ExistingFiles = entries
                    ReDim Me.CentralDirImage(centralSize)
                    Me.ZipFileStream.Seek(centralDirOffset, SeekOrigin.Begin)
                    Me.ZipFileStream.Read(Me.CentralDirImage, 0, centralSize)

                    ' Leave the pointer at the begining of central dir, to append new files
                    Me.ZipFileStream.Seek(centralDirOffset, SeekOrigin.Begin)
                    Return True
                End If
            Loop While (Me.ZipFileStream.Position > 0)

        Catch
        End Try

        Return False
    End Function
#End Region

#Region "IDisposable Members"
    ' <summary>
    ' Closes the Zip file stream
    ' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Me.Close()
    End Sub
#End Region

End Class