Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.IO
Imports System.Linq
Imports ClosedXML.Excel

'================ Theme =================
Module Theme
    Public ReadOnly Navy As Color = ColorTranslator.FromHtml("#0C2E4E")
    Public ReadOnly PrimaryBlue As Color = ColorTranslator.FromHtml("#0066A6")
    Public ReadOnly CardBg As Color = Color.FromArgb(18, 46, 78)
    Public ReadOnly CardBgAlt As Color = Color.FromArgb(22, 54, 91)
    Public ReadOnly TextMain As Color = Color.White
    Public ReadOnly TextMuted As Color = Color.FromArgb(190, 220, 240)
End Module

'================ Data model =================
Public Class Person
    Public Property Id As Integer
    Public Property FullName As String
    Public Property Email As String
    Public Property PhotoPath As String
End Class

'================ UI helpers =================
Friend Class GradientHeader
    Inherits Panel
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Using lg As New LinearGradientBrush(Me.ClientRectangle, Navy, PrimaryBlue, 0.0F)
            e.Graphics.FillRectangle(lg, Me.ClientRectangle)
        End Using
    End Sub
End Class

Friend Class SmoothListView
    Inherits ListView
    Public Sub New()
        Me.DoubleBuffered = True
        Me.OwnerDraw = True
        Me.HeaderStyle = ColumnHeaderStyle.None
        Me.View = View.List
    End Sub
End Class

'================ Main Form =================
Partial Public Class Form1
    ' Header
    Private ReadOnly header As New GradientHeader() With {.Dock = DockStyle.Top, .Height = 96}
    Private ReadOnly titleLbl As New Label() With {.Text = "FlightAssist Directory", .AutoSize = True, .ForeColor = TextMain, .Font = New Font("Segoe UI Semibold", 18, FontStyle.Bold)}
    Private ReadOnly statusLbl As New Label() With {.AutoSize = True, .ForeColor = Color.FromArgb(220, 240, 255)}
    Private ReadOnly logo As New PictureBox() With {.SizeMode = PictureBoxSizeMode.Zoom, .Width = 52, .Height = 52}

    ' Search
    Private ReadOnly txtSearch As New TextBox() With {.Width = 360}
    Private ReadOnly debounce As New System.Windows.Forms.Timer() With {.Interval = 1000}

    ' List
    Private ReadOnly lv As New SmoothListView()
    Private ReadOnly imgs As New ImageList() With {.ImageSize = New Size(48, 48), .ColorDepth = ColorDepth.Depth32Bit}
    Private RowHeight As Integer = 64

    ' Data
    Private ReadOnly people As New List(Of Person)()

    ' Paths
    Private ReadOnly fixedRoot As String = "D:\work\Flightassist"
    Private ReadOnly appRoot As String = If(Directory.Exists(fixedRoot), fixedRoot, AppDomain.CurrentDomain.BaseDirectory)

    Private ReadOnly excelPath As String = Path.Combine(appRoot, "data\people.xlsx")
    Private ReadOnly photosDir As String = Path.Combine(appRoot, "photos")
    Private ReadOnly assetsLogo As String = Path.Combine(appRoot, "assets\logo.png")

    Public Sub New()
        InitializeComponent()

        Me.Text = "Directory — Excel data (Styled)"
        Me.BackColor = CardBg
        Me.Size = New Size(900, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        BuildHeader()
        BuildList()
        BindEvents()

        ' أنشئ المجلدات تحت appRoot (مش BaseDirectory)
        Directory.CreateDirectory(Path.Combine(appRoot, "data"))
        Directory.CreateDirectory(photosDir)

        ' Load logo if exists
        Try
            If File.Exists(assetsLogo) Then
                logo.Image = Image.FromFile(assetsLogo)
            End If
        Catch
        End Try

        Try
            If Not File.Exists(excelPath) Then
                Throw New FileNotFoundException("Excel not found", excelPath)
            End If
            LoadFromExcel(excelPath)
            statusLbl.Text = $"Loaded: {people.Count} people"
        Catch ex As Exception
            MessageBox.Show("Excel load error: " & ex.Message, "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SeedPeople()
            statusLbl.Text = $"Loaded seed: {people.Count}"
        End Try

        PerformSearch("")
    End Sub

    Private Sub BuildHeader()
        Controls.Add(header)
        titleLbl.Location = New Point(84, 20)
        header.Controls.Add(titleLbl)
        logo.Location = New Point(20, 18)
        header.Controls.Add(logo)
        txtSearch.Location = New Point(22, 60)
        txtSearch.Width = 340
        txtSearch.BorderStyle = BorderStyle.FixedSingle
        header.Controls.Add(txtSearch)
        statusLbl.Location = New Point(380, 64)
        header.Controls.Add(statusLbl)
    End Sub

    Private Sub BuildList()
        lv.Location = New Point(16, header.Bottom + 12)
        lv.Size = New Size(ClientSize.Width - 32, ClientSize.Height - header.Height - 28)
        lv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        lv.SmallImageList = imgs
        Controls.Add(lv)
        AddHandler lv.DrawItem, AddressOf Lv_DrawItem
        AddHandler lv.Resize, Sub() lv.Invalidate()
    End Sub

    Private Sub BindEvents()
        AddHandler txtSearch.TextChanged,
            Sub()
                debounce.Stop()
                debounce.Start()
            End Sub
        AddHandler debounce.Tick, AddressOf DebouncedSearch
    End Sub

    Private Sub DebouncedSearch(sender As Object, e As EventArgs)
        debounce.Stop()
        PerformSearch(txtSearch.Text.Trim())
    End Sub

    '================ Excel =================
    Private Sub LoadFromExcel(path As String)
        people.Clear()
        Using wb As New XLWorkbook(path)
            Dim ws = wb.Worksheet("People") ' اسم الشيت لازم People
            Dim r As Integer = 2 ' الصف الأول للعناوين
            While Not ws.Cell(r, 1).IsEmpty()
                Dim p As New Person With {
                    .Id = ws.Cell(r, 1).GetValue(Of Integer)(),
                    .FullName = ws.Cell(r, 2).GetString(),
                    .Email = ws.Cell(r, 3).GetString(),
                    .PhotoPath = ws.Cell(r, 4).GetString()
                }
                people.Add(p)
                r += 1
            End While
        End Using
    End Sub

    '================ Search =================
    Private Sub PerformSearch(query As String)
        lv.BeginUpdate()
        lv.Items.Clear()
        imgs.Images.Clear()

        Dim res As IEnumerable(Of Person)
        If String.IsNullOrWhiteSpace(query) Then
            res = Enumerable.Empty(Of Person)()
        ElseIf Not IsQueryAllowed(query) Then
            res = Enumerable.Empty(Of Person)()
        Else
            Dim q = query.ToLowerInvariant()
            res = people.
                Where(Function(p) p.FullName.ToLower().Contains(q) _
                               OrElse p.Email.ToLower().Contains(q) _
                               OrElse p.Id.ToString().Contains(q)).
                OrderBy(Function(p) p.FullName).
                Take(6)
        End If

        For Each p In res
            Dim key As String = p.Id.ToString()
            imgs.Images.Add(key, GetThumb(p))
            Dim it As New ListViewItem With {.Text = p.FullName, .ImageKey = key, .Tag = p}
            lv.Items.Add(it)
        Next

        lv.EndUpdate()
        lv.Invalidate()
    End Sub

    Private Function IsQueryAllowed(q As String) As Boolean
        If q.Length < 2 OrElse q.Length > 30 Then Return False
        If q.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Length > 4 Then Return False
        If Not Regex.IsMatch(q, "^[\p{L}\p{Nd}\s@._+\-]+$") Then Return False
        Return True
    End Function

    '================ Owner-draw =================
    Private Sub Lv_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Dim bounds As New Rectangle(6, e.Bounds.Y + 4, lv.ClientSize.Width - 12, RowHeight - 8)
        Dim bg = If(e.ItemIndex Mod 2 = 0, CardBg, CardBgAlt)
        Using b As New SolidBrush(bg)
            e.Graphics.FillRectangle(b, bounds)
        End Using

        Dim p = TryCast(e.Item.Tag, Person)
        If p Is Nothing Then Return

        ' Thumbnail (circle)
        Dim img = imgs.Images(e.Item.ImageKey)
        Dim avatarRect As New Rectangle(bounds.X + 10, bounds.Y + 8, 48, 48)
        Using gp As New GraphicsPath()
            gp.AddEllipse(avatarRect)
            e.Graphics.SetClip(gp)
            e.Graphics.DrawImage(img, avatarRect)
            e.Graphics.ResetClip()
            Using pn As New Pen(Color.FromArgb(80, Color.White), 1.5F)
                e.Graphics.DrawEllipse(pn, avatarRect)
            End Using
        End Using

        ' Texts
        Dim namePt As New Point(avatarRect.Right + 12, bounds.Y + 10)
        Dim mailPt As New Point(avatarRect.Right + 12, bounds.Y + 32)
        Using nameFont As New Font("Segoe UI Semibold", 11, FontStyle.Bold),
              mailFont As New Font("Segoe UI", 9)
            TextRenderer.DrawText(e.Graphics, p.FullName, nameFont, namePt, TextMain, TextFormatFlags.NoPadding)
            TextRenderer.DrawText(e.Graphics, p.Email,   mailFont, mailPt,  TextMuted, TextFormatFlags.NoPadding)
        End Using
    End Sub

    '================ Thumbs / Avatars =================
    Private Function GetThumb(p As Person) As Image
        Const size As Integer = 48
        If p Is Nothing Then Return MakeAvatar("?", size, 0)
        Try
            Dim inPath As String = p.PhotoPath
            If Not String.IsNullOrWhiteSpace(inPath) Then
                If Not Path.IsPathRooted(inPath) Then inPath = Path.Combine(appRoot, inPath)
                Dim photoAbs = Path.GetFullPath(inPath)
                If File.Exists(photoAbs) Then
                    Using fs As New FileStream(photoAbs, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Using rawImg As Image = Image.FromStream(fs, False, False)
                            Using tmp As New Bitmap(rawImg)
                                Return ResizeToThumb(tmp, size, size)
                            End Using
                        End Using
                    End Using
                End If
            End If
        Catch
        End Try
        Return MakeAvatar(GetInitials(p.FullName), size, p.Id)
    End Function

    Private Function ResizeToThumb(src As Image, w As Integer, h As Integer) As Image
        Dim bmp As New Bitmap(w, h)
        Using g = Graphics.FromImage(bmp)
            g.CompositingQuality = CompositingQuality.HighSpeed
            g.InterpolationMode = InterpolationMode.HighQualityBilinear
            g.SmoothingMode = SmoothingMode.HighSpeed
            g.DrawImage(src, New Rectangle(0, 0, w, h))
        End Using
        Return bmp
    End Function

    Private Function MakeAvatar(initials As String, size As Integer, key As Integer) As Image
        Dim bmp As New Bitmap(size, size)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            Dim rnd As New Random(key Xor &H2EA3F2)
            Dim c As Color = Color.FromArgb(255, 90 + rnd.Next(110), 90 + rnd.Next(110), 90 + rnd.Next(110))
            Using b As New SolidBrush(c)
                g.FillEllipse(b, 0, 0, size - 1, size - 1)
            End Using
            Using pen As New Pen(Color.FromArgb(230, Color.White), 2)
                g.DrawEllipse(pen, 1, 1, size - 3, size - 3)
            End Using
            Using f As New Font("Segoe UI Semibold", If(initials.Length <= 2, size * 0.42F, size * 0.34F), FontStyle.Bold, GraphicsUnit.Pixel),
                  br As New SolidBrush(Color.White)
                Dim sz = g.MeasureString(initials, f)
                Dim pt As New PointF((size - sz.Width) / 2.0F, (size - sz.Height) / 2.0F - 2)
                g.DrawString(initials.ToUpperInvariant(), f, br, pt)
            End Using
        End Using
        Return bmp
    End Function

    Private Function GetInitials(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "?"
        Dim parts = name.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Take(3)
        Return New String(parts.Select(Function(s) s(0)).ToArray())
    End Function

    ' Seed fallback
    Private Sub SeedPeople()
        people.Clear()
        people.AddRange({
            New Person With {.Id = 6, .FullName = "Mohamed Imran", .Email = "mohamed.imran@acmecorp.com"},
            New Person With {.Id = 24, .FullName = "Sherine Salem", .Email = "sherine.s@acmecorp.com"},
            New Person With {.Id = 31, .FullName = "Mostafa Fathy", .Email = "mostafa.f@acmecorp.com"},
            New Person With {.Id = 42, .FullName = "Iman Mohamed", .Email = "iman.m@acmecorp.com"}
        })
    End Sub
End Class

'================ Startup (لو مش عامل Startup Form في Properties) =================
Module Startup
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form1())
    End Sub
End Module
