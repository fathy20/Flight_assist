Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions
Imports System.IO
Imports ClosedXML.Excel

'===================== Theme =====================
Module Theme
    Public ReadOnly Navy As Color = ColorTranslator.FromHtml("#0C2E4E")
    Public ReadOnly PrimaryBlue As Color = ColorTranslator.FromHtml("#0066A6")
    Public ReadOnly CardBg As Color = Color.FromArgb(18, 46, 78)
    Public ReadOnly CardBgAlt As Color = Color.FromArgb(22, 54, 91)
    Public ReadOnly TextMain As Color = Color.White
    Public ReadOnly TextMuted As Color = Color.FromArgb(190, 220, 240)
End Module

'==================== Model ======================
Public Class Person
    Public Property Id As Integer
    Public Property FullName As String
    Public Property Email As String
    Public Property PhotoPath As String
End Class

'================== Header (Gradient) ================
Friend Class GradientHeader
    Inherits Panel
    Public Sub New()
        Me.DoubleBuffered = True
        Me.Height = 96
        Me.Dock = DockStyle.Top
        Me.Padding = New Padding(16, 12, 16, 12)
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                    ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or
                    ControlStyles.SupportsTransparentBackColor, True)
        Me.BackColor = Color.Transparent
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Using lg As New LinearGradientBrush(Me.ClientRectangle, Theme.Navy, Theme.PrimaryBlue, 0.0F)
            e.Graphics.FillRectangle(lg, Me.ClientRectangle)
        End Using
        Using p As New Pen(Color.FromArgb(80, Color.White), 2)
            e.Graphics.DrawLine(p, 0, Me.Height - 1, Me.Width, Me.Height - 1)
        End Using
    End Sub
End Class

'================== Smoother ListView ==================
Friend Class SmoothListView
    Inherits ListView
    Public Sub New()
        Me.DoubleBuffered = True
        Me.OwnerDraw = True
        Me.HeaderStyle = ColumnHeaderStyle.None
        Me.View = View.List
        Me.BorderStyle = BorderStyle.None
    End Sub
End Class

'======================= Main Form ======================
Public Class Form1
    Inherits Form

    ' Header UI
    Private ReadOnly header As New GradientHeader()
    Private ReadOnly titleLbl As New Label() With {.AutoSize = True, .ForeColor = Theme.TextMain,
                                                   .Font = New Font("Segoe UI Semibold", 18, FontStyle.Bold),
                                                   .Text = "FlightAssist Directory",
                                                   .BackColor = Color.Transparent}
    Private ReadOnly statusLbl As New Label() With {.AutoSize = True, .ForeColor = Color.FromArgb(220, 240, 255),
                                                    .BackColor = Color.Transparent}
    Private ReadOnly logo As New PictureBox() With {.SizeMode = PictureBoxSizeMode.Zoom, .Size = New Size(56, 56)}

    ' Search
    Private ReadOnly txtSearch As New TextBox() With {.Width = 360, .BorderStyle = BorderStyle.FixedSingle}
    Private ReadOnly debounce As New Timer() With {.Interval = 600} ' 0.6s

    ' List
    Private ReadOnly lv As New SmoothListView()
    Private ReadOnly imgs As New ImageList() With {.ImageSize = New Size(48, 48), .ColorDepth = ColorDepth.Depth32Bit}
    Private RowHeight As Integer = 64

    ' Data
    Private ReadOnly people As New List(Of Person)()
    Private ReadOnly thumbCache As New Dictionary(Of String, Image)(StringComparer.OrdinalIgnoreCase)

    ' Package paths
    Private ReadOnly packageRoot As String = FindPackageRoot()
    Private ReadOnly excelPath As String = Path.Combine(packageRoot, "data\people.xlsx")
    Private ReadOnly photosDir As String = Path.Combine(packageRoot, "photos")
    Private ReadOnly logoPath As String = Path.Combine(packageRoot, "assets\logo.png")

    Public Sub New()
        Me.Text = "FlightAssist Directory"
        Me.BackColor = Theme.CardBg
        Me.Size = New Size(980, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        BuildHeader()
        BuildList()
        BindEvents()

        ' Logo
        Try
            If File.Exists(logoPath) Then logo.Image = Image.FromFile(logoPath)
        Catch
        End Try
        If logo.Image Is Nothing Then
            ' فfallback بسيط
            Dim bmp As New Bitmap(56, 56)
            Using g = Graphics.FromImage(bmp)
                g.SmoothingMode = SmoothingMode.AntiAlias
                Using br As New SolidBrush(Color.White)
                    g.FillEllipse(br, 2, 2, 52, 52)
                End Using
            End Using
            logo.Image = bmp
        End If

        ' تأمين ملف الإكسيل (لو ناقص/تالف يتبني نظيف)
        EnsurePeopleExcel(excelPath)

        ' Load data
        Try
            LoadFromExcel(excelPath)
            statusLbl.Text = $"Loaded: {people.Count} people"
        Catch ex As Exception
            MessageBox.Show("Excel load error: " & ex.Message & vbCrLf &
                            "سيتم تشغيل البرنامج ببيانات تجريبية.", "Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SeedPeople()
            statusLbl.Text = $"Loaded seed: {people.Count}"
        End Try

        PerformSearch("") ' بداية
    End Sub

    '----------------- Layout -----------------
    Private Sub BuildHeader()
        Controls.Add(header)

        logo.Location = New Point(header.Padding.Left, header.Padding.Top - 2)
        header.Controls.Add(logo)

        titleLbl.Location = New Point(logo.Right + 10, logo.Top + 6)
        header.Controls.Add(titleLbl)

        txtSearch.Location = New Point(logo.Left, logo.Bottom + 6)
        txtSearch.Width = Math.Max(360, Me.ClientSize.Width \ 3)
        txtSearch.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        header.Controls.Add(txtSearch)

        statusLbl.Location = New Point(txtSearch.Right + 12, txtSearch.Top + 2)
        statusLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        header.Controls.Add(statusLbl)

        AddHandler header.Resize, Sub()
                                      txtSearch.Width = Math.Max(360, Me.ClientSize.Width \ 3)
                                      statusLbl.Left = txtSearch.Right + 12
                                  End Sub
    End Sub

    Private Sub BuildList()
        lv.SmallImageList = imgs
        lv.Dock = DockStyle.Fill
        Controls.Add(lv)
        lv.BringToFront()
        AddHandler lv.DrawItem, AddressOf Lv_DrawItem
    End Sub

    Private Sub BindEvents()
        AddHandler txtSearch.TextChanged,
            Sub(sender As Object, e As EventArgs)
                debounce.Stop()
                debounce.Start()
            End Sub
        AddHandler debounce.Tick, AddressOf DebouncedSearch
    End Sub

    Private Sub DebouncedSearch(sender As Object, e As EventArgs)
        debounce.Stop()
        PerformSearch(txtSearch.Text.Trim())
    End Sub

    '---------------- Excel I/O ----------------
    Private Sub EnsurePeopleExcel(path As String)
        Try
            Dim ok As Boolean = False
            If File.Exists(path) Then
                Try
                    Using wb As New XLWorkbook(path)
                        Dim ws = wb.Worksheet("People")
                        ok = (ws IsNot Nothing)
                    End Using
                Catch
                    ok = False
                End Try
            End If
            If Not ok Then
                Dim dir = System.IO.Path.GetDirectoryName(path)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                CreateCleanPeopleExcel(path)
            End If
        Catch ex As Exception
            ' آخر حل: نبني ملف نظيف
            Try
                CreateCleanPeopleExcel(path)
            Catch
            End Try
        End Try
    End Sub

    Private Sub CreateCleanPeopleExcel(path As String)
        Using wb As New XLWorkbook()
            Dim ws = wb.AddWorksheet("People")
            ' Header
            ws.Cell(1, 1).Value = "Id"
            ws.Cell(1, 2).Value = "FullName"
            ws.Cell(1, 3).Value = "Email"
            ws.Cell(1, 4).Value = "PhotoPath"
            ' Samples
            ws.Cell(2, 1).Value = 6 : ws.Cell(2, 2).Value = "Mohamed Imran"
            ws.Cell(2, 3).Value = "mohamed.imran@acmecorp.com"
            ws.Cell(2, 4).Value = "photos\AA1LaybT.jpg"

            ws.Cell(3, 1).Value = 24 : ws.Cell(3, 2).Value = "Sherine Salem"
            ws.Cell(3, 3).Value = "sherine.s@acmecorp.com"
            ws.Cell(3, 4).Value = "photos\R.jpg"

            ws.Cell(4, 1).Value = 42 : ws.Cell(4, 2).Value = "Iman Mohamed"
            ws.Cell(4, 3).Value = "iman.m@acmecorp.com"
            ws.Cell(4, 4).Value = "photos\woman.jpg"

            ws.Cell(5, 1).Value = 31 : ws.Cell(5, 2).Value = "Mostafa Fathy"
            ws.Cell(5, 3).Value = "mostafa.f@acmecorp.com"
            ws.Cell(5, 4).Value = ""

            ws.Columns().AdjustToContents()
            wb.SaveAs(path)
        End Using
    End Sub

    Private Sub LoadFromExcel(path As String)
        people.Clear()
        Using wb As New XLWorkbook(path)
            Dim ws = wb.Worksheet("People")
            Dim r As Integer = 2
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

    '---------------- Search (priority) --------------
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
            Dim q = query.Trim()
            Dim qLower = q.ToLowerInvariant()

            res = people.
                Select(Function(p)
                           Dim nameLower = If(p.FullName, "").ToLowerInvariant()
                           Dim emailLower = If(p.Email, "").ToLowerInvariant()
                           Dim idExact = If(Integer.TryParse(q, Nothing) AndAlso p.Id.ToString() = q, 0, 1)
                           Dim nameStarts = If(nameLower.StartsWith(qLower), 0, 1)
                           Dim emailStarts = If(emailLower.StartsWith(qLower), 0, 1)
                           Dim wordStarts = If(nameLower.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).
                                               Any(Function(w) w.StartsWith(qLower)), 0, 1)
                           Dim containsHit = If(nameLower.Contains(qLower) OrElse emailLower.Contains(qLower), 0, 1)
                           Return New With {.P = p, .S = New Integer() {idExact, nameStarts, emailStarts, wordStarts, containsHit}}
                       End Function).
                OrderBy(Function(x) x.S(0)).
                ThenBy(Function(x) x.S(1)).
                ThenBy(Function(x) x.S(2)).
                ThenBy(Function(x) x.S(3)).
                ThenBy(Function(x) x.S(4)).
                ThenBy(Function(x) x.P.FullName).
                Select(Function(x) x.P).
                Take(12)   ' خليه 6 لو عايز نتائج أقل
        End If

        For Each p In res
            Dim key As String = p.Id.ToString()
            imgs.Images.Add(key, GetThumb(p))
            lv.Items.Add(New ListViewItem With {.Text = p.FullName, .ImageKey = key, .Tag = p})
        Next

        statusLbl.Text = $"Loaded: {people.Count} • Showing: {lv.Items.Count}"
        lv.EndUpdate()
        lv.Invalidate()
    End Sub

    Private Function IsQueryAllowed(q As String) As Boolean
        If q.Length < 2 OrElse q.Length > 30 Then Return False  ' اكتب 2 حروف على الأقل
        If q.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Length > 4 Then Return False
        If Not Regex.IsMatch(q, "^[\p{L}\p{Nd}\s@._+\-]+$") Then Return False
        Return True
    End Function

    '---------------- Owner-Draw Rows ----------------
    Private Sub Lv_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Dim bounds As New Rectangle(6, e.Bounds.Y + 4, lv.ClientSize.Width - 12, RowHeight - 8)
        Dim bg = If(e.ItemIndex Mod 2 = 0, Theme.CardBg, Theme.CardBgAlt)
        Using b As New SolidBrush(bg)
            e.Graphics.FillRectangle(b, bounds)
        End Using

        Dim p = TryCast(e.Item.Tag, Person)
        If p Is Nothing Then Return

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

        Dim namePt As New Point(avatarRect.Right + 12, bounds.Y + 10)
        Dim mailPt As New Point(avatarRect.Right + 12, bounds.Y + 32)
        Using nameFont As New Font("Segoe UI Semibold", 11, FontStyle.Bold),
              mailFont As New Font("Segoe UI", 9)
            TextRenderer.DrawText(e.Graphics, p.FullName, nameFont, namePt, Theme.TextMain, TextFormatFlags.NoPadding)
            TextRenderer.DrawText(e.Graphics, p.Email, mailFont, mailPt, Theme.TextMuted, TextFormatFlags.NoPadding)
        End Using
    End Sub

    '---------------- Thumbs / Avatars ----------------
    Private Function GetThumb(p As Person) As Image
        Const size As Integer = 48
        If p Is Nothing Then Return MakeAvatar("?", size, 0)

        Dim cacheKey As String = If(String.IsNullOrWhiteSpace(p.PhotoPath),
                                    $"__avatar_{p.Id}",
                                    Path.GetFullPath(Path.Combine(packageRoot, p.PhotoPath)))
        If thumbCache.ContainsKey(cacheKey) Then Return thumbCache(cacheKey)

        Try
            Dim inPath As String = p.PhotoPath
            If Not String.IsNullOrWhiteSpace(inPath) Then
                If Not Path.IsPathRooted(inPath) Then inPath = Path.Combine(packageRoot, inPath)
                Dim abs = Path.GetFullPath(inPath)
                If File.Exists(abs) Then
                    Using fs As New FileStream(abs, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Using rawImg As Image = Image.FromStream(fs, False, False)
                            Using tmp As New Bitmap(rawImg)
                                Dim th = ResizeToThumb(tmp, size, size)
                                thumbCache(cacheKey) = th
                                Return th
                            End Using
                        End Using
                    End Using
                End If
            End If
        Catch
        End Try

        Dim av = MakeAvatar(GetInitials(p.FullName), size, p.Id)
        thumbCache(cacheKey) = av
        Return av
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
            Using f As New Font("Segoe UI Semibold",
                                 If(initials.Length <= 2, size * 0.42F, size * 0.34F),
                                 FontStyle.Bold, GraphicsUnit.Pixel),
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

    '---------------- Seed fallback ----------------
    Private Sub SeedPeople()
        people.Clear()
        people.AddRange({
            New Person With {.Id = 6, .FullName = "Mohamed Imran", .Email = "mohamed.imran@acmecorp.com"},
            New Person With {.Id = 24, .FullName = "Sherine Salem", .Email = "sherine.s@acmecorp.com"},
            New Person With {.Id = 31, .FullName = "Mostafa Fathy", .Email = "mostafa.f@acmecorp.com"},
            New Person With {.Id = 42, .FullName = "Iman Mohamed", .Email = "iman.m@acmecorp.com"}
        })
    End Sub

    '------------- Locate package folder -------------
    Private Shared Function FindPackageRoot() As String
        Dim dir As DirectoryInfo = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory)
        For i = 0 To 6
            Dim cand = Path.Combine(dir.FullName, "package")
            If Directory.Exists(cand) AndAlso
               Directory.Exists(Path.Combine(cand, "assets")) AndAlso
               Directory.Exists(Path.Combine(cand, "data")) AndAlso
               Directory.Exists(Path.Combine(cand, "photos")) Then
                Return cand
            End If
            If dir.Parent Is Nothing Then Exit For
            dir = dir.Parent
        Next
        Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "package")
    End Function
End Class
