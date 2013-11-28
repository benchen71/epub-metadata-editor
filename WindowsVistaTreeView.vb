Imports System.ComponentModel
Imports System.Drawing.Drawing2D

' -----------------------------------------------------------------------------
' Copyright © 2007 Stephen J Whiteley
'
' This source code is provided AS IS, with no warranty expressed or implied,
' including without limitation, warranties of merchantability or
' fitness for a particular purpose or any warranty of title or
' non-infringement. This disclaimer must be passed on whenever the Software
' is distributed in either source form or as a derivative works.
' It may be used for both commercial and non-commercial applications.
'
' -----------------------------------------------------------------------------

Public Class VistaTreeView : Inherits TreeView

    Sub New()
        ' ---------------------------------------------------------------------
        ' New Treeview
        ' ---------------------------------------------------------------------
        '
        MyBase.LineColor = SystemColors.GrayText
        _expanderStyle = TreeExpanderStyle.Arrow

        Me.BackColor = Color.FromArgb(232, 232, 232)
        Call LoadExpanderImages()
        '
        ' Set the style to double buffered and do all the drawing in
        ' the paint event (onPaint) of the treeview.
        '
        ' This reduces the flicker, but is a lot more work
        '
        MyBase.SetStyle(ControlStyles.DoubleBuffer _
             Or ControlStyles.UserPaint _
             Or ControlStyles.AllPaintingInWmPaint, _
             True)
        MyBase.UpdateStyles() ' Update the styles
        '
        ' Set the treeview styles
        Call SetTreeviewStyle()
    End Sub

#Region "-  Enums  -"
    Public Enum TreeExpanderStyle
        PlusMinus = 0
        Arrow = 1
    End Enum
#End Region

#Region "-  Properties  -"

#Region "-  Private Variables  -"
    Private CollapsedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAARpJREFUOE+lk9tKAmEUhfWhfIiew1eyUkkvPCaNdZdHbPCACIqCklZERQZCZhNZeZylS5jRhvlHcAb2zcD61tqLfzsBOGx9BNgZXdwffCKYLCEgFXF2IcN3XoA3nsNJJAtPOK1Ptd5Z+21NdUDwsgxVBRZLFdPZEj9/CyjjOd6VKd6GEzwPfnH/OobryG0OCEilvWK58SIGMLbRmW6ac+fpG1fynRjgX+9sjE0AY1PcfPiCVLAAnMby+s4UGqfWVZDI98QJjqMZ08LoTHG5PUI82xUDPJH0v7YZmyk08U3rA9HMHsBuYbvOFOcaQ4RTt9YJWFix2Ueq8rgpjDszNp0pDl1bAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Private CollapsedArrowString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAANFJREFUOE+l00kKhDAQheHq+x/KlaALxXkeF04o4g1ep0J3Y0OEiILoIvnyK9QLAD26GHhyKzc7joNpmmhZFtq2jfZ9p+M4lGsvgTOyrqtEVKWXQN/3YGQcR1nCiDZg2zbatsUZmef5HlBVFZqmQdd1smQYBn3AsizkeY6yLP8Q7U8wTRNpmv4QLhAl+gUMRFGEJEnA76KE6rrWBwzDQBAE4KdAKMsyKoriHvBBSJTQF9H+B7zZdV3yPI9836cwDCmOY/2CO7PxaJDkJN85TbX2Db5d1YfJcQ3TAAAAAElFTkSuQmCC"
    Private CollapsedIcon As Image
    Private ExpandedString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAO9JREFUOE+lk9sKAVEUhnkoD+E5vJJzuHEshzs1DiGUQ0RRigsKJceRcWZ+tppxmjXKTK2bqe/791p7Lz0AnaaPCbSUDI8mK3iiBbgjebjCOThCGdiDKVh9SZi9nFylWvue9wyVBZ5YEaIIXK4ijqcrtvsLeOGMGX/EeH7AYLJDdyjAYDQpC9yRwk+43d/QAnZstWQGN3prWuC890wdW4IrHZ4W2AJpuWfW52cxuNha0gKLP/E1sNdkBmebC1pg9nFv01aCU/W5ukC6KgrmqjN1AbtnNThentIC9sKUhvf5j3yJ/+6DpkV6bPK/yRJ3A/PE7e2oP8DgAAAAAElFTkSuQmCCAPjCzMoz/hO+xEPvwdYhbS75UGdNtwLNm+LI5h1FwAAAAABJRU5ErkJggg=="
    Private ExpandedArrowString As String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAM9JREFUOE/N0kkKg0AQBdCf+x/KlaALxXkeF04o4g0qXU1sDPQqBhLhI72o17+gH0SEWx8Dd3JrWLa/c/t3Ac/ziOM4Dtm2LWNZFpmmKWMYhsq1tVphmiYw0Pc9tW1LVVVRnueUpilFUURBEEjAdd23tdVhWRZckaZpqCxLiSRJIocFwpfogW3bwMg4juDqXdfRifCwQCCawPd9PbDvO9Z1VQgPMcJ/0QRZloGRMAz1wHEcOJF5njEMA14I6rpGURQSieNYD3z6Hv7oIf1shSf3G9UMQ+Vu/QAAAABJRU5ErkJggg=="
    Private ExpandedIcon As Image
    '
    ' Text Formatting
    Private TextFormat As New StringFormat(StringFormatFlags.NoWrap)
    '
    Private HighlightBorderPen As New Pen(Color.FromArgb(255, 127, 157, 185))
    Private HighlightColor As Color = Color.FromArgb(255, 204, 230, 255)
    Private HighlightBrush As New SolidBrush(HighlightColor)
    Private RootColor As Color = Color.FromArgb(255, 204, 204, 204)
    '
    Private _haveImages As Boolean
    Private _expanderStyle As TreeExpanderStyle
    Private _backgroundimage As Image
    Private _lastKnowTopIndex As String = ""
    '
    Private NodeBrush As Brush
#End Region

#Region "-  Appearance  -"

    Private mBaseColor As Color = Color.FromArgb(0, 5, 10)
    ' The backing color that the rest of 
    ' the Checkbox is drawn. For a glassier 
    ' effect set this property to Transparent.
    <Category("Appearance"), DefaultValue(GetType(Color), "0, 5, 10"), Description("The backing color that the rest of the highlighted area is drawn. For a glassier effect set this property to Transparent.")> _
    Public Property NodeBaseColor() As Color
        Get
            Return mBaseColor
        End Get
        Set(ByVal value As Color)
            mBaseColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mNodeColor As Color = Color.FromArgb(0, 0, 32)
    ' The bottom color of the Checkbox that 
    ' will be drawn over the base color.
    <Category("Appearance"), DefaultValue(GetType(Color), "0, 0, 32"), Description("The bottom color of the highlighted node that will be drawn over the base color.")> _
    Public Property NodeColor() As Color
        Get
            Return mNodeColor
        End Get
        Set(ByVal value As Color)
            mNodeColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mSelectedColor As Color = Color.White
    ' The font color of a selected node.
    <Category("Appearance"), DefaultValue(GetType(Color), "White"), Description("The font color of a selected node.")> _
    Public Property FontColorSelected() As Color
        Get
            Return mSelectedColor
        End Get
        Set(ByVal value As Color)
            mSelectedColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mHotTrackingColor As Color = Color.Silver
    ' The font color of a hovered node if the hottracking is on.
    <Category("Appearance"), DefaultValue(GetType(Color), "Silver"), Description("The font color of a hovered node if the hottracking is on.")> _
    Public Property FontColorHotTracking() As Color
        Get
            Return mHotTrackingColor
        End Get
        Set(ByVal value As Color)
            mHotTrackingColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mFontHotTracking As Font = Me.Font
    ' The hovered node font if the hottracking is on.
    <Category("Appearance"), Description("The hovered node font if the hottracking is on.")> _
    Public Property FontHotTracking() As Font
        Get
            Return mFontHotTracking
        End Get
        Set(ByVal value As Font)
            mFontHotTracking = value
            Me.Invalidate()
        End Set
    End Property

#End Region

#Region "-  Behavior  -"
    <Category("Behavior"), DefaultValue(GetType(TreeExpanderStyle), "Arrow"), Description("Collapsing arrow style.")> _
    Public Property ExpanderStyle() As TreeExpanderStyle
        Get
            Return _expanderStyle
        End Get
        Set(ByVal value As TreeExpanderStyle)
            If _expanderStyle <> value Then
                _expanderStyle = value
                ' Load the expander styles and redraw
                ' the whole treeview
                Call LoadExpanderImages()
                MyBase.Invalidate()
            End If
        End Set
    End Property

#End Region

#Region "-  Image  -"
    ' The background image
    <Category("Image"), DefaultValue(GetType(Image), Nothing), Description("The image displayed on the treeview background."), EditorBrowsable(EditorBrowsableState.Always), Browsable(True), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), Bindable(True)> _
    Public Overrides Property BackgroundImage() As Image
        Get
            Return _backgroundimage
        End Get
        Set(ByVal value As Image)
            _backgroundimage = value
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "-  Shadowed properties  -"

    Public Shadows Property ImageList() As ImageList
        Get
            Return MyBase.ImageList
        End Get
        Set(ByVal value As ImageList)
            MyBase.ImageList = value
            _haveImages = (MyBase.ImageList IsNot Nothing)
        End Set
    End Property

    <Browsable(True), [ReadOnly](False)> _
    Public Shadows Property Font() As Font
        Get
            Return MyBase.Font
        End Get
        Set(ByVal value As Font)
            MyBase.Font = value
            Call SetTreeviewStyle()
        End Set
    End Property

    <Browsable(False), [ReadOnly](True)> _
    Public Shadows ReadOnly Property DrawMode() As TreeViewDrawMode
        Get
            Return MyBase.DrawMode
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
    Public Shadows ReadOnly Property ShowLines() As Boolean
        Get
            Return MyBase.ShowLines
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
    Public Shadows ReadOnly Property ShowRootLines() As Boolean
        Get
            Return MyBase.ShowRootLines
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
    Public Shadows ReadOnly Property LabelEdit() As Boolean
        Get
            Return MyBase.LabelEdit
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
    Public Shadows ReadOnly Property ShowPlusMinus() As Boolean
        Get
            Return MyBase.ShowPlusMinus
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
  Public Shadows ReadOnly Property LineColor() As Color
        Get
            Return MyBase.LineColor
        End Get
    End Property

    <Browsable(False), [ReadOnly](True)> _
  Public Shadows ReadOnly Property CheckBoxes() As Boolean
        Get
            Return MyBase.CheckBoxes
        End Get
    End Property

#End Region

#End Region

#Region "-  Overrided subs  -"
    Protected Overrides Sub InitLayout()
        ' ---------------------------------------------------------------------
        ' New Treeview: Initialize the Layout
        ' ---------------------------------------------------------------------
        MyBase.InitLayout()
        '
        ' Overridden properties
        MyBase.DrawMode = TreeViewDrawMode.OwnerDrawAll
        MyBase.ShowPlusMinus = True
        MyBase.ShowLines = False
        MyBase.ShowRootLines = True
        MyBase.CheckBoxes = False
        MyBase.LineColor = Color.Black
        ''
    End Sub
    Protected Overrides Sub OnAfterCollapse(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        ' ---------------------------------------------------------------------
        ' After collapsing - redo the node count
        ' ---------------------------------------------------------------------
        MyBase.OnAfterCollapse(e)
    End Sub

    Protected Overrides Sub OnAfterExpand(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        ' ---------------------------------------------------------------------
        ' After Expanding - redo the node count
        ' ---------------------------------------------------------------------
        MyBase.OnAfterExpand(e)
    End Sub

    Protected Overrides Sub OnParentFontChanged(ByVal e As System.EventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnParentFontChanged event
        ' ---------------------------------------------------------------------
        MyBase.OnParentFontChanged(e)
        Call SetTreeviewStyle()
    End Sub

    Protected Overrides Sub OnBeforeSelect(ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnBeforeSelect event
        ' ---------------------------------------------------------------------
        MyBase.OnBeforeSelect(e)
        Try
            Dim _node As VistaNode = e.Node
            If _node.Title Then e.Cancel = True
        Catch ex As Exception
        End Try
    End Sub

    Protected Overrides Sub OnBeforeExpand(ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnBeforeExpand event
        ' ---------------------------------------------------------------------
        MyBase.OnBeforeExpand(e)
        MyBase.Invalidate()
    End Sub

    Protected Overrides Sub OnBeforeCollapse(ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnBeforeCollapse event
        ' ---------------------------------------------------------------------
        Try
            Dim _node As VistaNode = e.Node
            If _node.Title Then e.Cancel = True
        Catch ex As Exception
        End Try

        MyBase.OnBeforeCollapse(e)
        MyBase.Invalidate()
    End Sub
    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnPaintBackground event
        ' ---------------------------------------------------------------------
        If _backgroundimage IsNot Nothing Then
            pevent.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Low
            pevent.Graphics.DrawImage(_backgroundimage, New Rectangle(0, 0, Me.Width, Me.Height))
        Else
            Dim b As LinearGradientBrush = New LinearGradientBrush(Point.Empty, New PointF(0, Me.Height), Color.Black, Me.BackColor)
            Dim border As New GraphicsPath
            pevent.Graphics.FillRectangle(b, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
            Dim glowSize As Int32 = Convert.ToInt32(Convert.ToSingle(Me.Height) / 2)
            Dim glow As Rectangle = New Rectangle(0, Me.Height - glowSize - 1, Me.Width - 1, glowSize)
            b = New LinearGradientBrush(New Point(0, glow.Top - 1), New PointF(0, glow.Bottom), Color.FromArgb(0, Color.FromArgb(&H43, &H53, &H7A)), Color.FromArgb(&H43, &H53, &H7A))
            pevent.Graphics.FillRectangle(b, glow)
        End If
        ' If scroll, draw it again
        If Me.TopNode IsNot Nothing Then
            If _lastKnowTopIndex <> Me.TopNode.Name Then
                Me.Invalidate()
                _lastKnowTopIndex = Me.TopNode.Name
            End If
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        ' ---------------------------------------------------------------------
        ' Override the OnPaint event to paint the nodes
        ' ---------------------------------------------------------------------
        '        MyBase.OnPaint(e)
        Call PaintNodes(MyBase.Nodes, e)
    End Sub

#End Region

#Region "-  Private subs  -"
    Private Sub LoadExpanderImages()
        ' ---------------------------------------------------------------------
        ' Loads the included and defined expander images
        ' ---------------------------------------------------------------------
        '
        Dim b() As Byte
        '
        ' Expanded Image
        Select Case _expanderStyle
            Case TreeExpanderStyle.Arrow
                b = Convert.FromBase64String(ExpandedArrowString)
            Case Else
                b = Convert.FromBase64String(ExpandedString)
        End Select
        Dim ms As New System.IO.MemoryStream()
        ms.Write(b, 0, b.Length)
        ExpandedIcon = Image.FromStream(ms)
        '
        ' Collapsed Image
        Select Case _expanderStyle
            Case TreeExpanderStyle.Arrow
                b = Convert.FromBase64String(CollapsedArrowString)
            Case Else
                b = Convert.FromBase64String(CollapsedString)
        End Select
        ms.Position = 0
        ms.Write(b, 0, b.Length)
        CollapsedIcon = Image.FromStream(ms)
        ms.Close()
        ''
    End Sub
    Private Sub SetTreeviewStyle()
        ' ---------------------------------------------------------------------
        ' Sets the various Treeview settings
        ' ---------------------------------------------------------------------
        '
        ' Unselect the node if it's a parent and are not allowed
        ' to select parent nodes
        '
        ' ensure a node is selected
        If MyBase.SelectedNode Is Nothing Then
            For Each n As TreeNode In MyBase.Nodes
                MyBase.SelectedNode = n
                Exit For
            Next
        End If
        '
        Dim g As Graphics = Me.CreateGraphics
        Dim sz As SizeF = g.MeasureString(CollapsedString.Substring(0, 2), Me.Font, 125)
        Dim newHeight As Integer = Convert.ToInt32(Math.Ceiling(sz.Height))
        If newHeight > MyBase.ItemHeight Then MyBase.ItemHeight = newHeight
        g.Dispose()
        '
        NodeBrush = New SolidBrush(Me.ForeColor)
        ''
    End Sub
#End Region

#Region "-  Drawing  -"
    Private Sub PaintNodes(ByVal parentNodes As TreeNodeCollection, ByVal args As PaintEventArgs)
        ' ---------------------------------------------------------------------
        ' Paint the nodes
        ' ---------------------------------------------------------------------
        '
        ' Hottracking - find a node over which is the mouse
        Dim on_item As VistaNode = Nothing
        If Me.HotTracking Then
            Dim LocalMousePosition As Point
            LocalMousePosition = Me.PointToClient(Cursor.Position)
            on_item = Me.GetNodeAt(LocalMousePosition)
            If on_item IsNot Nothing Then
                If on_item.Title Then on_item = Nothing
            End If
        End If

        For Each n As TreeNode In parentNodes
            Dim b As New Rectangle(0, n.Bounds.Y, MyBase.ClientSize.Width, n.Bounds.Height)
            If args.ClipRectangle.IntersectsWith(b) Then
                ' This node intersects with the clip rectangle, so draw it
                b = New Rectangle(n.Bounds.X, n.Bounds.Y, MyBase.ClientSize.Width - n.Bounds.X - 1, n.Bounds.Height)
                Call DrawThisNode(args.Graphics, n, b, n Is on_item)
            End If
            ' Check for children nodes
            ' (Recursively draw the nodes)
            If n.Nodes.Count > 0 Then Call PaintNodes(n.Nodes, args)
        Next
    End Sub

    Private Sub DrawThisNode(ByVal g As Graphics, ByVal node As TreeNode, ByVal rBounds As Rectangle, ByVal HotTracked As Boolean)
        ' ---------------------------------------------------------------------
        ' Draw a specific node
        ' ---------------------------------------------------------------------
        '
        Dim selected As Boolean = node.IsSelected ' convenience
        '
        Dim hasExpander As Boolean ' has an Expand/collaps been drawn for this node
        Dim imageSize As Size ' has an image been drawn for this node
        '
        '
        ' Draw selected node background
        If selected Then DrawNodeBackground(g, New Rectangle(-2, rBounds.Y, MyBase.Width + 2, rBounds.Height))
        '
        ' Draw expander and image
        Dim expRect As Rectangle = New Rectangle(rBounds.X - 17, rBounds.Y, rBounds.Width + 12, rBounds.Height)
        If MyBase.ImageList IsNot Nothing Then
            Select Case MyBase.ImageList.ImageSize.Width
                Case 1 To 16
                    expRect = New Rectangle(rBounds.X - MyBase.ImageList.ImageSize.Width - 20, rBounds.Y, rBounds.Width + MyBase.ImageList.ImageSize.Width + 3, rBounds.Height)
                Case Is > 16
                    expRect = New Rectangle(rBounds.X - 2 * MyBase.ImageList.ImageSize.Width - 4, rBounds.Y, rBounds.Width + 2 * MyBase.ImageList.ImageSize.Width - 13, rBounds.Height)
            End Select
        End If
        hasExpander = DrawNodeExpander(g, node, expRect)
        imageSize = DrawNodeImage(g, node, New Rectangle(expRect.X + 22, rBounds.Y, expRect.Width - 17, rBounds.Height))
        '
        '
        ' Reduce the SIZE of the bounds, used for drawing rectangles
        Dim rBorder As New Rectangle(expRect.X + imageSize.Width + 22, rBounds.Y, expRect.Width - imageSize.Width - 17 + rBounds.X, rBounds.Height - 1)
        '
        ' Draw the actual Text
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Dim rText As RectangleF = DrawNodeText(g, node, rBorder, selected, HotTracked)
        ' Draw any node customization
        Call DrawNodeCustom(g, node, rText)
        ''
    End Sub

    Private Function DrawNodeImage(ByVal g As Graphics, ByVal node As TreeNode, ByVal r As Rectangle) As Size
        ' ---------------------------------------------------------------------
        ' Draw the IMAGE associated with this node
        ' ---------------------------------------------------------------------
        If _haveImages = False Then Return New Size(0, 0)

        Dim n As VistaNode = TryCast(node, VistaNode)
        If n IsNot Nothing Then
            ' No images for titles - this code should be changed, but the drawing text code has to be changed as well
            If n.Title Then Return New Size(0, 0)
        End If
        ' Have images: does this NODE have an image?
        '
        Dim img As Drawing.Image = Nothing ' Define an image object
        '
        If node.IsSelected Then
            ' SELECTED
            If node.SelectedImageIndex >= 0 Then
                img = MyBase.ImageList.Images(node.SelectedImageIndex)
            ElseIf node.SelectedImageKey IsNot Nothing Then
                img = MyBase.ImageList.Images(node.SelectedImageKey)
            End If
        Else
            ' NOT SELECTED
            If node.ImageIndex >= 0 Then
                img = MyBase.ImageList.Images(node.ImageIndex)
            ElseIf node.ImageKey IsNot Nothing Then
                img = MyBase.ImageList.Images(node.ImageKey)
            End If
        End If
        If img Is Nothing Then Return New Size(0, 0)
        ' If an image exists, then draw this image at the correct position
        Dim yPosition As Int32 = r.Y + (MyBase.ItemHeight - img.Height) \ 2
        g.DrawImage(img, r.X, yPosition)
        Return New Size(img.Width + 8, img.Height) ' move the text 8px to the right of the image
    End Function

    Private Function DrawNodeExpander(ByVal g As Graphics, ByVal node As TreeNode, ByVal r As Rectangle) As Boolean
        ' ---------------------------------------------------------------------
        ' Draw the Expander, returning T/F for expander drawn
        ' ---------------------------------------------------------------------
        ' Draw the PLUS/MINUS Expand/Collapse Icon if necessary
        'If (node.Level = 0 And _collapsibleParent = False) Then Return False
        ' Set the position for the Expand/collapse icon
        Dim result As Boolean
        Dim n As VistaNode = TryCast(node, VistaNode)
        If n IsNot Nothing Then
            If n.Title Then Return result
        End If

        If node.IsExpanded Then
            ' Draw the EXPANDED Icon
            Dim yPosition As Int32 = r.Y + (MyBase.ItemHeight - ExpandedIcon.Height) \ 2
            g.DrawImage(ExpandedIcon, r.X, yPosition)
            result = True
        Else
            If node.Nodes.Count > 0 Then
                ' Draw the Collapsed Icon
                Dim yPosition As Int32 = r.Y + (MyBase.ItemHeight - CollapsedIcon.Height) \ 2
                g.DrawImage(CollapsedIcon, r.X, yPosition)
                result = True
            End If
        End If
        Return result ' Has an expander been drawn
    End Function

    Private Sub DrawNodeBackground(ByVal g As Graphics, ByVal rBorder As Rectangle)
        ' ---------------------------------------------------------------------
        ' Draws the background for a Node
        ' ---------------------------------------------------------------------
        DrawBackground(g, rBorder)
        DrawOuterStroke(g, rBorder)
        DrawInnerStroke(g, rBorder)
        DrawHighlight(g, rBorder)
        DrawGlossyEffect(g, rBorder)
    End Sub

    Private Sub DrawBackground(ByVal g As Graphics, ByVal r As Rectangle)
        Dim alpha As Int32 = 127
        Dim sb As SolidBrush = New SolidBrush(Me.NodeBaseColor)
        g.FillRectangle(sb, r)
        SetClip(g)
        g.ResetClip()
        sb = New SolidBrush(Color.FromArgb(alpha, Me.NodeColor))
        g.FillRectangle(sb, r)
    End Sub

    Private Sub DrawOuterStroke(ByVal g As Graphics, ByVal r As Rectangle)
        Dim p As Pen = New Pen(Me.NodeColor)
        g.DrawLine(p, r.Left, r.Top, r.Right - 1, r.Top)
        g.DrawLine(p, r.Left, r.Bottom - 1, r.Right - 1, r.Bottom - 1)
    End Sub

    Private Sub DrawInnerStroke(ByVal g As Graphics, ByVal r As Rectangle)
        r.X += 1
        r.Y += 1
        r.Width -= 2
        r.Height -= 2
        Dim p As Pen = New Pen(Me.HighlightColor)
        g.DrawLine(p, r.Left, r.Top, r.Right - 1, r.Top)
        g.DrawLine(p, r.Left, r.Bottom - 1, r.Right - 1, r.Bottom - 1)
    End Sub

    Private Sub DrawHighlight(ByVal g As Graphics, ByVal rBorder As Rectangle)
        If Not Me.Enabled Then Return
        Dim alpha As Int32 = 150
        Dim r As Rectangle = New Rectangle(rBorder.X, rBorder.Y, rBorder.Width, rBorder.Height / 2)
        Dim lg As LinearGradientBrush = New LinearGradientBrush(r, Color.FromArgb(alpha, Me.HighlightColor), Color.FromArgb(alpha / 3, Me.HighlightColor), 90.1F, True)
        lg.WrapMode = WrapMode.TileFlipXY
        g.FillRectangle(lg, r)
    End Sub
    Private Sub DrawGlossyEffect(ByVal g As Graphics, ByVal rBorder As Rectangle)
        SetClip(g)
        Dim r As Rectangle = New Rectangle(rBorder.X, rBorder.Y, rBorder.Width, rBorder.Height)
        Dim glow As GraphicsPath = CreateTopRoundRectangle(r, 2)
        Dim gl As PathGradientBrush = New PathGradientBrush(glow)
        gl.CenterColor = Color.FromArgb(70, Color.FromArgb(141, 189, 255))
        gl.SurroundColors = New Color() {Color.FromArgb(0, Color.FromArgb(141, 189, 255))}
        g.FillPath(gl, glow)
        g.ResetClip()
    End Sub

    Private Sub SetClip(ByVal g As Graphics)
        Dim r As Rectangle = Me.ClientRectangle
        r.X += 1
        r.Y += 1
        r.Width -= 3
        r.Height -= 3
        g.SetClip(r)
    End Sub

    Private Function DrawNodeText(ByVal g As Graphics, ByVal node As TreeNode, ByVal rBounds As Rectangle, ByVal selected As Boolean, ByVal HotTracked As Boolean) As RectangleF
        ' ---------------------------------------------------------------------
        ' Draw the actual node text
        ' Returns a rectangle where the text was drawn
        ' ---------------------------------------------------------------------
        Dim txt As String = node.Text, title As Boolean = False
        Dim NodeFont As Font
        If HotTracked And Not selected Then
            NodeFont = Me.FontHotTracking
            NodeBrush = New SolidBrush(Me.FontColorHotTracking)
        ElseIf selected Then
            NodeFont = Me.Font
            NodeBrush = New SolidBrush(Me.FontColorSelected)
        Else
            NodeFont = Me.Font
            If node.ForeColor.IsEmpty Then
                NodeBrush = New SolidBrush(Me.ForeColor)
            Else
                NodeBrush = New SolidBrush(node.ForeColor)
            End If
        End If
        Dim n As VistaNode = TryCast(node, VistaNode)
        If n IsNot Nothing Then
            If n.NodeFont IsNot Nothing Then NodeFont = n.NodeFont
            If n.Highlighted Then
                NodeBrush = New SolidBrush(n.HighlightedColor)
                NodeFont = New Font(NodeFont.FontFamily, NodeFont.Size, FontStyle.Bold)
            End If
            title = n.Title
        End If

        Dim sz As SizeF = g.MeasureString(txt, NodeFont, rBounds.Width, TextFormat)
        Dim textBounds As RectangleF
        If title Then
            textBounds = New RectangleF(10, rBounds.Y + Convert.ToInt32((rBounds.Height - sz.Height) / 2), rBounds.Width, Convert.ToInt32(sz.Height))
        Else
            textBounds = New RectangleF(rBounds.X, rBounds.Y + Convert.ToInt32((rBounds.Height - sz.Height) / 2), rBounds.Width, Convert.ToInt32(sz.Height))
        End If
        g.DrawString(node.Text, NodeFont, NodeBrush, textBounds, TextFormat)
        ' Return the rectanglewhere the text was drawn
        textBounds.Width = sz.Width
        Return textBounds
    End Function

    Private Sub DrawNodeCustom(ByVal g As Graphics, ByVal node As TreeNode, ByVal rText As RectangleF)
        ' ---------------------------------------------------------------------
        ' Draw a CUSTOM Node if it's been supplied
        ' ---------------------------------------------------------------------
        Dim n As VistaNode = TryCast(node, VistaNode)
        If n Is Nothing Then Return ' Nothing is customized
        '
        ' Draw the CUSTOM Properties for this node
        Dim br As Brush
        ' Get the Post-text color
        Dim col As Color = n.PostTextColor
        ' If there isn't a color, use the derault color
        Dim NodeFont As Font
        If n.Font Is Nothing Then NodeFont = Me.Font Else NodeFont = n.Font
        If col.IsEmpty Then
            br = NodeBrush
        Else
            br = New SolidBrush(n.PostTextColor)
        End If
        If n.Highlighted Then
            br = New SolidBrush(n.HighlightedColor)
            NodeFont = New Font(Me.Font.FontFamily, Me.Font.Size, FontStyle.Bold)
        End If
        g.DrawString(String.Format(Globalization.CultureInfo.CurrentUICulture, " {0}", n.PostText), NodeFont, br, rText.X + rText.Width, rText.Y)
    End Sub

    Public Function RoundRect(ByVal r As RectangleF, ByVal r1 As Single, ByVal r2 As Single, ByVal r3 As Single, ByVal r4 As Single) As GraphicsPath
        Dim x As Single = r.X, y = r.Y, w = r.Width, h = r.Height
        Dim rr As GraphicsPath = New GraphicsPath()
        rr.AddBezier(x, y + r1, x, y, x + r1, y, x + r1, y)
        rr.AddLine(x + r1, y, x + w - r2, y)
        rr.AddBezier(x + w - r2, y, x + w, y, x + w, y + r2, x + w, y + r2)
        rr.AddLine(x + w, y + r2, x + w, y + h - r3)
        rr.AddBezier(x + w, y + h - r3, x + w, y + h, x + w - r3, y + h, x + w - r3, y + h)
        rr.AddLine(x + w - r3, y + h, x + r4, y + h)
        rr.AddBezier(x + r4, y + h, x, y + h, x, y + h - r4, x, y + h - r4)
        rr.AddLine(x, y + h - r4, x, y + r1)
        Return rr
    End Function

#End Region

#Region "-  private functions  -"
    ' Creates a rectangle rounded on the top
    ' <param name="rectangle">Base rectangle</param>
    ' <param name="radius">Radius of the top corners</param>
    ' <returns>Rounded rectangle (on top) as a GraphicsPath object</returns>
    Private Function CreateTopRoundRectangle(ByVal rectangle As Rectangle, ByVal radius As Int32) As GraphicsPath
        Dim path As GraphicsPath = New GraphicsPath()

        Dim l As Int32 = rectangle.Left
        Dim t As Int32 = rectangle.Top
        Dim w As Int32 = rectangle.Width
        Dim h As Int32 = rectangle.Height
        Dim d As Int32 = radius << 1

        path.AddArc(l, t, d, d, 180, 90) ' topleft
        path.AddLine(l + radius, t, l + w - radius, t) ' top
        path.AddArc(l + w - d, t, d, d, 270, 90) ' topright
        path.AddLine(l + w, t + radius, l + w, t + h) ' right
        path.AddLine(l + w, t + h, l, t + h) ' bottom
        path.AddLine(l, t + h, l, t + radius) ' left
        path.CloseFigure()

        Return path
    End Function

#End Region
End Class

<Serializable()> Public Class VistaNode : Inherits TreeNode

    Private _postText As String
    Private _postTextColor As Color
    Private _postTextShow As Boolean
    Private _font As Font
    Private _highlighted As Boolean
    Private _highlightedColor As Color
    Private _title As Boolean

    Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, _
                    ByVal context As System.Runtime.Serialization.StreamingContext)
        ' ---------------------------------------------------------------------
        ' Constructor for Serialization
        ' ---------------------------------------------------------------------
        Call MyBase.New(info, context)
    End Sub

    Public Sub New()
        _postTextColor = Color.Empty
        _postTextShow = True
        _postText = String.Empty
        _font = Nothing
        _highlighted = False
        _highlightedColor = Color.Empty
        _title = False
    End Sub

    Public Property PostText() As String
        Get
            Return _postText
        End Get
        Set(ByVal value As String)
            _postText = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property PostTextColor() As Color
        Get
            Return _postTextColor
        End Get
        Set(ByVal value As Color)
            _postTextColor = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property Font() As Font
        Get
            Return _font
        End Get
        Set(ByVal value As Font)
            _font = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property HighlightedColor() As Color
        Get
            Return _highlightedColor
        End Get
        Set(ByVal value As Color)
            _highlightedColor = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property Highlighted() As Boolean
        Get
            Return _highlighted
        End Get
        Set(ByVal value As Boolean)
            _highlighted = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property PostTextShow() As Boolean
        Get
            Return _postTextShow
        End Get
        Set(ByVal value As Boolean)
            _postTextShow = value
            Call ForceUpdate()
        End Set
    End Property

    Public Property Title() As Boolean
        Get
            Return _title
        End Get
        Set(ByVal value As Boolean)
            _title = value
        End Set
    End Property

    Private Sub ForceUpdate()
        ' ---------------------------------------------------------------------
        ' Force a change in the node text to update the
        ' node in the parent treeview
        ' Rather crude, but appears to work
        ' ---------------------------------------------------------------------
        MyBase.Text = MyBase.Text
    End Sub
End Class