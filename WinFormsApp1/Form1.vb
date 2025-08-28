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
        Using lg As New LinearGradientBrush(Me.ClientRectangle, Theme.Navy, Theme.PrimaryBlue, 0.0F)
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
    Private ReadOnly appLogoBox As New PictureBox() With {.SizeMode = PictureBoxSizeMode.Zoom}
    Private ReadOnly titleLbl As New Label() With {.Text = "Members", .AutoSize = True, .ForeColor = Theme.TextMain, .Font = New Font("Segoe UI Semibold", 18, FontStyle.Bold)}
    Private ReadOnly statusLbl As New Label() With {.AutoSize = True, .ForeColor = Color.FromArgb(220, 240, 255)}

    ' Search
    Private ReadOnly txtSearch As New TextBox() With {.Width = 360}
    Private ReadOnly debounce As New System.Windows.Forms.Timer() With {.Interval = 500}

    ' List
    Private ReadOnly lv As New SmoothListView()
    Private ReadOnly noResultsLbl As New Label() With {
        .Text = "No results found.",
        .Font = New Font("Segoe UI", 12),
        .ForeColor = Theme.TextMuted,
        .BackColor = Theme.CardBg,
        .Visible = False,
        .AutoSize = False,
        .TextAlign = ContentAlignment.MiddleCenter
    }
    Private ReadOnly imgs As New ImageList() With {.ImageSize = New Size(48, 48), .ColorDepth = ColorDepth.Depth32Bit}
    Private RowHeight As Integer = 64

    ' Data
    Private ReadOnly people As New List(Of Person)()

    ' Paths
    Private ReadOnly projectRoot As String = FindProjectRoot(AppDomain.CurrentDomain.BaseDirectory)
    Private ReadOnly packageRoot As String = Path.Combine(projectRoot, "package")
    Private ReadOnly excelPath As String = Path.Combine(packageRoot, "data\people.xlsx")
    Private ReadOnly photosDir As String = Path.Combine(packageRoot, "photos")
    Private ReadOnly assetsDir As String = Path.Combine(packageRoot, "assets")

    Public Sub New()
        InitializeComponent()

        ' <<< تم تعديل هذا السطر ليصبح اسم الشركة فقط
        Me.Text = "Flight Assist"

        Me.BackColor = Theme.CardBg
        Me.Size = New Size(880, 560)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        ' This code block loads the window icon from logo.ico
        Try
            Dim iconPath As String = Path.Combine(assetsDir, "logo.ico")
            If File.Exists(iconPath) Then
                Me.Icon = New Icon(iconPath)
            Else
                Console.WriteLine($"Warning: logo.ico not found at {iconPath}. Default icon will be used for the form.")
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading form icon: {ex.Message}")
        End Try

        BuildHeader()
        BuildList()
        BindEvents()

        ' Ensure directories exist
        Directory.CreateDirectory(Path.Combine(packageRoot, "data"))
        Directory.CreateDirectory(photosDir)

        Try
            If Not File.Exists(excelPath) Then
                Throw New FileNotFoundException("Excel file not found", excelPath)
            End If
            LoadFromExcel(excelPath)
            EnsurePhotosAndPaths(excelPath)
            statusLbl.Text = $"Loaded: {people.Count} people"
        Catch ex As Exception
            MessageBox.Show("Excel load error: " & ex.Message & vbCrLf &
                            "The program will run with sample data.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SeedPeople()
            statusLbl.Text = $"Loaded: {people.Count} (seed)"
        End Try

        PerformSearch("")
    End Sub

    Private Shared Function FindProjectRoot(startPath As String) As String
        Dim dir = New DirectoryInfo(startPath)
        ' Search for the solution file (.sln) to identify the project root
        While dir IsNot Nothing AndAlso Not dir.GetFiles("*.sln").Any()
            dir = dir.Parent
        End While
        Return If(dir IsNot Nothing, dir.FullName, startPath)
    End Function

    Private Sub BuildHeader()
        Controls.Add(header)

        ' Setup the logo inside the header using logo.png
        appLogoBox.Size = New Size(40, 40)
        appLogoBox.Location = New Point(20, 18)
        Try
            Dim logoPngPath As String = Path.Combine(assetsDir, "logo.png")
            If File.Exists(logoPngPath) Then
                appLogoBox.Image = Image.FromFile(logoPngPath)
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading header logo: {ex.Message}")
        End Try
        header.Controls.Add(appLogoBox)

        ' Adjust title position to make room for the logo
        titleLbl.Location = New Point(appLogoBox.Right + 12, 20)
        header.Controls.Add(titleLbl)

        ' Adjust search box position
        txtSearch.PlaceholderText = "Search by name, email, or ID..."
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
        lv.BackColor = Theme.CardBg
        Controls.Add(lv)

        noResultsLbl.Bounds = lv.Bounds
        noResultsLbl.Anchor = lv.Anchor
        Controls.Add(noResultsLbl)
        noResultsLbl.BringToFront()

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

    Private Sub EnsurePhotosAndPaths(excelFile As String)
        Dim updated As Boolean = False
        Using wb As New XLWorkbook(excelFile)
            Dim ws = wb.Worksheet("People")
            Dim r As Integer = 2
            While Not ws.Cell(r, 1).IsEmpty()
                Dim id As Integer = ws.Cell(r, 1).GetValue(Of Integer)()
                Dim fullName As String = ws.Cell(r, 2).GetString()
                Dim curPath As String = ws.Cell(r, 4).GetString()

                Dim needsAvatar As Boolean = String.IsNullOrWhiteSpace(curPath)
                If Not needsAvatar Then
                    Dim absPath = If(Path.IsPathRooted(curPath), curPath, Path.Combine(packageRoot, curPath))
                    If Not File.Exists(absPath) Then needsAvatar = True
                End If

                If needsAvatar Then
                    Dim relPath As String = $"photos\emp_{id}.jpg"
                    Dim destPath As String = Path.Combine(packageRoot, relPath)
                    Directory.CreateDirectory(Path.GetDirectoryName(destPath))
                    Using bmp = MakeAvatar(GetInitials(fullName), 256, id)
                        SaveJpeg(destPath, bmp, 85L)
                    End Using
                    ws.Cell(r, 4).Value = relPath
                    updated = True

                    Dim p = people.FirstOrDefault(Function(person) person.Id = id)
                    If p IsNot Nothing Then p.PhotoPath = relPath
                End If
                r += 1
            End While
            If updated Then wb.Save()
        End Using
    End Sub

    Private Sub PerformSearch(query As String)
        lv.BeginUpdate()
        lv.Items.Clear()
        imgs.Images.Clear()

        Dim results As IEnumerable(Of Person)

        If String.IsNullOrWhiteSpace(query) Then
            results = people.OrderBy(Function(p) p.FullName)
        ElseIf Not IsQueryAllowed(query) Then
            results = Enumerable.Empty(Of Person)()
        Else
            Dim q = query.ToLowerInvariant()
            results = people.
                Where(Function(p) p.FullName.ToLower().Contains(q) _
                               OrElse p.Email.ToLower().Contains(q) _
                               OrElse p.Id.ToString().Contains(q)).
                OrderBy(Function(p) p.FullName)
        End If

        For Each p In results
            Dim key As String = p.Id.ToString()
            imgs.Images.Add(key, GetThumb(p))
            Dim it As New ListViewItem With {.Text = p.FullName, .ImageKey = key, .Tag = p}
            lv.Items.Add(it)
        Next

        If lv.Items.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(query) Then
            noResultsLbl.Text = $"No results found for '{query}'"
            noResultsLbl.Visible = True
        Else
            noResultsLbl.Visible = False
        End If

        lv.EndUpdate()
        lv.Invalidate()
    End Sub

    Private Function IsQueryAllowed(q As String) As Boolean
        If q.Length < 1 OrElse q.Length > 30 Then Return False
        If q.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Length > 4 Then Return False
        If Not Regex.IsMatch(q, "^[\p{L}\p{Nd}\s@._+\-]+$") Then Return False
        Return True
    End Function

    Private Sub Lv_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Dim bounds As New Rectangle(6, e.Bounds.Y + 4, lv.ClientSize.Width - 12, RowHeight - 8)
        Dim bg = If(e.ItemIndex Mod 2 = 0, Theme.CardBg, Theme.CardBgAlt)
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
            TextRenderer.DrawText(e.Graphics, p.FullName, nameFont, namePt, Theme.TextMain, TextFormatFlags.NoPadding)
            TextRenderer.DrawText(e.Graphics, p.Email, mailFont, mailPt, Theme.TextMuted, TextFormatFlags.NoPadding)
        End Using

        ' ID right
        Dim idStr = $"ID {p.Id}"
        Using small As New Font("Segoe UI", 9, FontStyle.Regular)
            Dim sz = TextRenderer.MeasureText(idStr, small)
            Dim idPt As New Point(bounds.Right - sz.Width - 12, bounds.Y + (RowHeight - sz.Height) \ 2)
            TextRenderer.DrawText(e.Graphics, idStr, small, idPt, Theme.TextMuted, TextFormatFlags.NoPadding)
        End Using
    End Sub

    Private Function GetThumb(p As Person) As Image
        Const size As Integer = 48
        If p Is Nothing Then Return MakeAvatar("?", size, 0)

        Try
            Dim photoPath As String = p.PhotoPath
            If Not String.IsNullOrWhiteSpace(photoPath) Then
                Dim fullPath = If(Path.IsPathRooted(photoPath), photoPath, Path.Combine(packageRoot, photoPath))
                If File.Exists(fullPath) Then
                    ' Use a memory stream to avoid locking the file
                    Dim fileBytes = File.ReadAllBytes(fullPath)
                    Using ms As New MemoryStream(fileBytes)
                        Using rawImg As Image = Image.FromStream(ms)
                            Return ResizeToThumb(rawImg, size, size)
                        End Using
                    End Using
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading photo for {p.FullName}: {ex.Message}")
        End Try

        ' Fallback to avatar if photo not found or failed to load
        Return MakeAvatar(GetInitials(p.FullName), size, p.Id)
    End Function

    Private Function ResizeToThumb(src As Image, w As Integer, h As Integer) As Image
        Dim bmp As New Bitmap(w, h)
        Using g = Graphics.FromImage(bmp)
            g.CompositingQuality = CompositingQuality.HighQuality
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.SmoothingMode = SmoothingMode.AntiAlias
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

    Private Sub SaveJpeg(dest As String, img As Image, quality As Long)
        Dim enc = ImageCodecInfo.GetImageEncoders().First(Function(c) c.MimeType = "image/jpeg")
        Dim ep As New EncoderParameters(1)
        ep.Param(0) = New EncoderParameter(Encoder.Quality, quality)
        img.Save(dest, enc, ep)
    End Sub

    Private Function GetInitials(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "?"
        Dim parts = name.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Take(3)
        Return New String(parts.Select(Function(s) s(0)).ToArray())
    End Function

    Private Sub SeedPeople()
        people.Clear()
        people.AddRange({
            New Person With {.Id = 6, .FullName = "Mohamed Imran", .Email = "mohamed.imran@acmecorp.com"},
            New Person With {.Id = 24, .FullName = "Sherine Salem", .Email = "sherine.s@acmecorp.com"},
            New Person With {.Id = 31, .FullName = "Mostafa Fathy", .Email = "mostafa.f@acmecorp.com"},
            New Person With {.Id = 42, .FullName = "Iman Mohamed", .Email = "iman.m@acmecorp.com"},
            New Person With {.Id = 57, .FullName = "Karim Nabil", .Email = "karim.n@acmecorp.com"}
        })
    End Sub
End Class