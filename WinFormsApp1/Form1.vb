Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.IO
Imports System.Linq
Imports ClosedXML.Excel
Imports System.Runtime.InteropServices

' Helper module for DPI awareness.
Module DPIHelper
    ' P/Invoke to set the process as DPI aware.
    ' This makes the application look crisp on high-resolution screens.
    <DllImport("user32.dll")>
    Private Function SetProcessDPIAware() As Boolean
    End Function

    Sub EnableHighDPI()
        SetProcessDPIAware()
    End Sub
End Module

'================ Theme =================
' Centralized color theme for easy UI adjustments.
Module Theme
    Public ReadOnly Navy As Color = ColorTranslator.FromHtml("#0C2E4E")
    Public ReadOnly PrimaryBlue As Color = ColorTranslator.FromHtml("#0066A6")
    Public ReadOnly CardBg As Color = Color.White  ' Changed from a dark blue for a lighter, modern look.
    Public ReadOnly CardBgAlt As Color = Color.White ' Also changed to white to simplify the list view.
    Public ReadOnly TextMain As Color = Color.Blue ' Changed from White to have better contrast on the new white background.
    Public ReadOnly TextMuted As Color = Color.Black ' Changed from a light blue for better readability.
End Module

'================ Data model =================
' Simple class to hold person data.
Public Class Person
    Public Property Id As Integer
    Public Property FullName As String
    Public Property Email As String
    Public Property PhotoPath As String
End Class

'================ UI helpers =================
' A custom panel for the header.
Friend Class GradientHeader
    Inherits Panel
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        ' Switched from a gradient to a solid white background for a cleaner design.
        Using b As New SolidBrush(Color.White)
            e.Graphics.FillRectangle(b, Me.ClientRectangle)
        End Using
    End Sub
End Class

' Custom ListView with double buffering enabled to prevent flickering during redraw.
Friend Class SmoothListView
    Inherits ListView
    Public Sub New()
        Me.DoubleBuffered = True
        Me.OwnerDraw = True ' We'll be handling the drawing of items ourselves.
        Me.HeaderStyle = ColumnHeaderStyle.None ' No column headers needed.
        Me.View = View.List
    End Sub
End Class

'================ Main Form =================
Partial Public Class Form1
    ' Header controls
    Private ReadOnly header As New GradientHeader() With {.Dock = DockStyle.Top, .Height = 96}
    Private ReadOnly appLogoBox As New PictureBox() With {.SizeMode = PictureBoxSizeMode.Zoom}
    Private ReadOnly titleLbl As New Label() With {.Text = "Employees", .AutoSize = True, .ForeColor = Color.Black, .Font = New Font("Segoe UI Semibold", 18, FontStyle.Bold)}
    Private ReadOnly statusLbl As New Label() With {.AutoSize = True, .ForeColor = Color.Gray}

    ' Search controls
    Private ReadOnly txtSearch As New TextBox() With {.Width = 360}
    Private ReadOnly debounce As New System.Windows.Forms.Timer() With {.Interval = 500} ' Delays search to avoid firing on every keystroke.

    ' List controls
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

    ' Data
    Private ReadOnly people As New List(Of Person)()

    ' Paths
    Private ReadOnly projectRoot As String = FindProjectRoot(AppDomain.CurrentDomain.BaseDirectory)
    Private ReadOnly packageRoot As String = Path.Combine(projectRoot, "package")
    Private ReadOnly excelPath As String = Path.Combine(packageRoot, "data\people.xlsx")
    Private ReadOnly photosDir As String = Path.Combine(packageRoot, "photos")
    Private ReadOnly assetsDir As String = Path.Combine(packageRoot, "assets")

    Public Sub New()
        ' Enable DPI Awareness for proper scaling on high-resolution screens.
        If Environment.OSVersion.Version.Major >= 6 Then
            DPIHelper.EnableHighDPI()
        End If

        InitializeComponent()

        Me.Text = "Flight Assist"
        Me.BackColor = Color.White

        ' Use percentages of the screen size for the initial form dimensions for better adaptability.
        Dim screenSize = Screen.PrimaryScreen.WorkingArea.Size
        Me.Size = New Size(Math.Min(880, CInt(screenSize.Width * 0.7)),
                             Math.Min(560, CInt(screenSize.Height * 0.8)))

        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        ' Enable automatic scaling based on DPI.
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)

        ' Make the window resizable.
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(600, 400)

        ' This code block loads the window icon from favicon.ico
        Try
            Dim iconPath As String = Path.Combine(assetsDir, "favicon.ico")
            If File.Exists(iconPath) Then
                Me.Icon = New Icon(iconPath)
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading form icon: {ex.Message}")
        End Try

        BuildHeader()
        BuildList()
        BindEvents()

        ' Ensure the necessary directories exist before trying to access them.
        Directory.CreateDirectory(Path.Combine(projectRoot, "data"))
        Directory.CreateDirectory(photosDir)

        Try
            If Not File.Exists(excelPath) Then
                Throw New FileNotFoundException("Excel file not found", excelPath)
            End If
            LoadFromExcel(excelPath)
            ' EnsurePhotosAndPaths(excelPath) ' This method is currently empty, keeping the call just in case.
            statusLbl.Text = $"Loaded: {people.Count} people"
        Catch ex As Exception
            ' If Excel fails to load, fall back to sample data so the app can still run.
            MessageBox.Show("Excel load error: " & ex.Message & vbCrLf &
                            "The program will run with sample data.", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            SeedPeople()
            statusLbl.Text = $"Loaded: {people.Count} (seed)"
        End Try

        PerformSearch("") ' Initial load of all people.
    End Sub

    ' Helper to find the project's root directory by looking for the .sln file.
    ' This makes pathing to assets more reliable.
    Private Shared Function FindProjectRoot(startPath As String) As String
        Dim dir = New DirectoryInfo(startPath)
        ' Search for the solution file (.sln) to identify the project root
        While dir IsNot Nothing AndAlso Not dir.GetFiles("*.sln").Any()
            dir = dir.Parent
        End While
        Return If(dir IsNot Nothing, dir.FullName, startPath)
    End Function

    Private Sub BuildHeader()
        ' Make the header height responsive to the form's size.
        header.Height = Math.Max(96, CInt(Me.Height * 0.15))
        Controls.Add(header)

        ' Scale the logo based on the header's height.
        Dim logoSize As Integer = Math.Min(40, header.Height - 20)
        appLogoBox.Size = New Size(logoSize, logoSize)
        appLogoBox.Location = New Point(20, (header.Height - logoSize) \ 2)

        Try
            Dim logoPngPath As String = Path.Combine(assetsDir, "logo.png")
            If File.Exists(logoPngPath) Then
                appLogoBox.Image = Image.FromFile(logoPngPath)
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading header logo: {ex.Message}")
        End Try
        header.Controls.Add(appLogoBox)

        ' Position the title label relatively.
        titleLbl.Location = New Point(appLogoBox.Right + 12, 16) ' Moved Y position up a bit.
        ' Scale the font size based on the form's width.
        Dim fontSize As Single = Math.Max(14, Math.Min(20, Me.Width / 50))
        titleLbl.Font = New Font("Segoe UI Semibold", fontSize, FontStyle.Bold)
        header.Controls.Add(titleLbl)

        ' Make the search box width responsive.
        txtSearch.PlaceholderText = "Search by name, email, or ID..."
        txtSearch.Location = New Point(22, header.Height - 30)
        txtSearch.Width = Math.Min(340, Me.Width - 200)
        txtSearch.BorderStyle = BorderStyle.FixedSingle
        txtSearch.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        header.Controls.Add(txtSearch)

        ' Position the status label.
        statusLbl.Location = New Point(txtSearch.Right + 20, header.Height - 26)
        statusLbl.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        header.Controls.Add(statusLbl)
    End Sub

    Private Sub BuildList()
        lv.Location = New Point(16, header.Bottom + 12)
        lv.Size = New Size(ClientSize.Width - 32, ClientSize.Height - header.Height - 28)
        lv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        lv.SmallImageList = imgs
        lv.BackColor = Color.White
        Controls.Add(lv)

        ' Position the "No Results" label to perfectly cover the list view when needed.
        noResultsLbl.Bounds = lv.Bounds
        noResultsLbl.Anchor = lv.Anchor
        noResultsLbl.BackColor = Color.White
        noResultsLbl.ForeColor = Color.Black
        Controls.Add(noResultsLbl)
        noResultsLbl.BringToFront()

        AddHandler lv.DrawItem, AddressOf Lv_DrawItem
        AddHandler lv.Resize, Sub()
                                  ' Redraw the list on resize.
                                  lv.Invalidate()
                                  ' Update image size dynamically based on the list view's width.
                                  Dim newImageSize As Integer = Math.Max(32, Math.Min(64, lv.Width \ 15))
                                  If imgs.ImageSize.Width <> newImageSize Then
                                      imgs.ImageSize = New Size(newImageSize, newImageSize)
                                      PerformSearch(txtSearch.Text) ' Reload images to reflect the new size.
                                  End If
                              End Sub

        ' Add a handler to rebuild the header when the form is resized.
        AddHandler Me.Resize, Sub()
                                  ' This ensures controls inside the header reposition correctly.
                                  BuildHeader()
                              End Sub
    End Sub

    Private Sub BindEvents()
        ' Use a debouncer for the search text box to prevent lagging.
        AddHandler txtSearch.TextChanged,
            Sub()
                debounce.Stop()
                debounce.Start()
            End Sub
        AddHandler debounce.Tick, AddressOf DebouncedSearch
    End Sub

    ' This runs after the user stops typing for a moment.
    Private Sub DebouncedSearch(sender As Object, e As EventArgs)
        debounce.Stop()
        PerformSearch(txtSearch.Text.Trim())
    End Sub

    Private Sub LoadFromExcel(path As String)
        people.Clear()
        Using wb As New XLWorkbook(path)
            Dim ws = wb.Worksheet("People")
            Dim r As Integer = 2 ' Start from the second row to skip headers.
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
        ' No avatar generation, no Excel update.
        ' This method is now empty because we always use profile.png for missing photos.
    End Sub

    Private Sub PerformSearch(query As String)
        lv.BeginUpdate() ' Prevents flickering while updating the list.
        lv.Items.Clear()
        imgs.Images.Clear()

        Dim results As IEnumerable(Of Person)

        If String.IsNullOrWhiteSpace(query) Then
            ' If search is empty, show all people.
            results = people.OrderBy(Function(p) p.FullName)
            statusLbl.Text = $"Loaded: {people.Count} people"
        ElseIf Not IsQueryAllowed(query) Then
            ' Basic query validation failed.
            results = Enumerable.Empty(Of Person)()
            statusLbl.Text = "No matches"
        Else
            ' Filter people based on the query.
            Dim q = query.ToLowerInvariant()
            results = people.
                Where(Function(p) p.FullName.ToLower().StartsWith(q) _
                              OrElse p.Email.ToLower().StartsWith(q) _
                              OrElse p.Id.ToString().StartsWith(q)).
                OrderBy(Function(p) p.FullName)

            ' Show count of matches in status label
            Dim matchCount = results.Count()
            If matchCount > 0 Then
                statusLbl.Text = $"Found: {matchCount} " & If(matchCount = 1, "person", "people")
            Else
                statusLbl.Text = "No matches"
            End If
        End If

        ' Populate the ListView with search results.
        For Each p In results
            Dim key As String = p.Id.ToString()
            imgs.Images.Add(key, GetThumb(p))
            Dim it As New ListViewItem With {.Text = p.FullName, .ImageKey = key, .Tag = p}
            lv.Items.Add(it)
        Next

        ' Show or hide the "No results" message.
        If lv.Items.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(query) Then
            noResultsLbl.Text = $"No results found for '{query}'"
            noResultsLbl.Visible = True
        Else
            noResultsLbl.Visible = False
        End If

        lv.EndUpdate()
        lv.Invalidate() ' Force a repaint.
    End Sub

    ' Basic validation to prevent overly long or complex queries.
    Private Function IsQueryAllowed(q As String) As Boolean
        If q.Length < 1 OrElse q.Length > 30 Then Return False
        If q.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Length > 4 Then Return False
        ' Allow letters, numbers, and common email/path characters.
        If Not Regex.IsMatch(q, "^[\p{L}\p{Nd}\s@._+\-]+$") Then Return False
        Return True
    End Function

    ' Custom drawing logic for each item in the ListView.
    Private Sub Lv_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        e.Graphics.SmoothingMode = SmoothingMode.HighQuality
        Dim bounds As New Rectangle(12, e.Bounds.Y + 6, lv.ClientSize.Width - 24, RowHeight - 12)
        Dim bg = If(e.ItemIndex Mod 2 = 0, Theme.CardBg, Theme.CardBgAlt)
        Dim radius As Integer = 18 ' For rounded corners.

        ' Create a margin between each employee card for better visual separation.
        Dim marginTop As Integer = 8
        Dim marginBottom As Integer = 8
        Dim cardRect As New Rectangle(bounds.X, bounds.Y + marginTop, bounds.Width, bounds.Height - marginTop - marginBottom)

        ' Draw the rounded rectangle card shape.
        Using path As New GraphicsPath()
            path.AddArc(cardRect.X, cardRect.Y, radius, radius, 180, 90)
            path.AddArc(cardRect.Right - radius, cardRect.Y, radius, radius, 270, 90)
            path.AddArc(cardRect.Right - radius, cardRect.Bottom - radius, radius, radius, 0, 90)
            path.AddArc(cardRect.X, cardRect.Bottom - radius, radius, radius, 90, 90)
            path.CloseFigure()

            ' subtle top gradient.
            Using topGrad As New LinearGradientBrush(
                New Rectangle(bounds.X, bounds.Y, bounds.Width, bounds.Height \ 2),
                Color.FromArgb(60, Theme.PrimaryBlue),
                Color.Transparent,
                LinearGradientMode.Vertical)
                e.Graphics.FillPath(topGrad, path)
            End Using

            ' subtle bottom gradient.
            Using bottomGrad As New LinearGradientBrush(
                New Rectangle(bounds.X, bounds.Y + bounds.Height \ 2, bounds.Width, bounds.Height \ 2),
                Color.Transparent,
                Color.FromArgb(40, Theme.Navy),
                LinearGradientMode.Vertical)
                e.Graphics.FillPath(bottomGrad, path)
            End Using

            ' Card background color.
            Using b As New SolidBrush(bg)
                e.Graphics.FillPath(b, path)
            End Using
        End Using

        ' Highlight the selected item.
        If (e.State And ListViewItemStates.Selected) = ListViewItemStates.Selected Then
            Using b As New SolidBrush(Color.FromArgb(60, Theme.PrimaryBlue))
                e.Graphics.FillRectangle(b, bounds)
            End Using
        End If

        Dim p = TryCast(e.Item.Tag, Person)
        If p Is Nothing Then Return

        ' Draw the circular avatar.
        Dim img = imgs.Images(e.Item.ImageKey)
        Dim avatarRect As New Rectangle(bounds.X + 16, bounds.Y + 8, 56, 56)
        Using gp As New GraphicsPath()
            gp.AddEllipse(avatarRect)
            e.Graphics.SetClip(gp) ' Clip the drawing region to the circle.
            e.Graphics.DrawImage(img, avatarRect)
            e.Graphics.ResetClip()
            ' Add a subtle border to the avatar.
            Using pn As New Pen(Color.FromArgb(120, Color.White), 2.5F)
                e.Graphics.DrawEllipse(pn, avatarRect)
            End Using
        End Using

        ' Draw the person's name.
        Dim namePt As New Point(avatarRect.Right + 18, bounds.Y + 14)
        Using nameFont As New Font("Segoe UI Semibold", 13, FontStyle.Bold)
            TextRenderer.DrawText(e.Graphics, p.FullName, nameFont, namePt, Theme.TextMain, TextFormatFlags.NoPadding)
        End Using

        ' Only show the email when a search is active to keep the main list clean.
        If Not String.IsNullOrWhiteSpace(txtSearch.Text) Then
            Dim mailPt As New Point(avatarRect.Right + 18, bounds.Y + 38)
            Using mailFont As New Font("Segoe UI", 10)
                TextRenderer.DrawText(e.Graphics, p.Email, mailFont, mailPt, Theme.TextMuted, TextFormatFlags.NoPadding)
            End Using
        End If

        ' Draw the ID, right-aligned.
        Dim idStr = $"ID {p.Id}"
        Using small As New Font("Segoe UI", 10, FontStyle.Regular)
            Dim sz = TextRenderer.MeasureText(idStr, small)
            Dim idPt As New Point(bounds.Right - sz.Width - 18, bounds.Y + (RowHeight - sz.Height) \ 2)
            TextRenderer.DrawText(e.Graphics, idStr, small, idPt, Theme.TextMuted, TextFormatFlags.NoPadding)
        End Using
    End Sub

    Private Function GetThumb(p As Person) As Image
        Const size As Integer = 48
        If p Is Nothing Then Return MakeAvatar("?", size, 0)

        Try
            Dim photoPath As String = p.PhotoPath
            If Not String.IsNullOrWhiteSpace(photoPath) Then
                ' Handle both absolute and relative paths.
                Dim fullPath = If(Path.IsPathRooted(photoPath), photoPath, Path.Combine(packageRoot, photoPath))
                If File.Exists(fullPath) Then
                    ' Load the image from the specified path.
                    Dim fileBytes = File.ReadAllBytes(fullPath)
                    Using ms As New MemoryStream(fileBytes)
                        Using rawImg As Image = Image.FromStream(ms)
                            Return ResizeToThumb(rawImg, size, size)
                        End Using
                    End Using
                End If
            End If
        Catch ex As Exception
            ' Log error but don't crash the app.
            Console.WriteLine($"Error loading photo for {p.FullName}: {ex.Message}")
        End Try

        ' If the specified photo doesn't exist, fall back to a default profile picture.
        Dim defaultProfilePath As String = Path.Combine(photosDir, "profile.png")
        If File.Exists(defaultProfilePath) Then
            Dim fileBytes = File.ReadAllBytes(defaultProfilePath)
            Using ms As New MemoryStream(fileBytes)
                Using rawImg As Image = Image.FromStream(ms)
                    Return ResizeToThumb(rawImg, size, size)
                End Using
            End Using
        End If

        ' Final fallback: return an empty bitmap if no images are found.
        Return New Bitmap(size, size)
    End Function

    ' High-quality image resizing function.
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

    ' Generates a colored circle with initials, used as a placeholder avatar.
    Private Function MakeAvatar(initials As String, size As Integer, key As Integer) As Image
        Dim bmp As New Bitmap(size, size)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            ' Use a seeded random to get consistent colors for the same person.
            Dim rnd As New Random(key Xor &H2EA3F2)
            Dim c As Color = Color.FromArgb(255, 90 + rnd.Next(110), 90 + rnd.Next(110), 90 + rnd.Next(110))
            Using b As New SolidBrush(c)
                g.FillEllipse(b, 0, 0, size - 1, size - 1)
            End Using
            Using pen As New Pen(Color.FromArgb(230, Color.White), 2)
                g.DrawEllipse(pen, 1, 1, size - 3, size - 3)
            End Using
            ' Dynamically adjust font size based on the number of initials.
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

    ' Extracts initials from a full name (e.g., "John Fitzgerald Kennedy" -> "JFK").
    Private Function GetInitials(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "?"
        Dim parts = name.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Take(3)
        Return New String(parts.Select(Function(s) s(0)).ToArray())
    End Function

    ' Sample data used if the Excel file can't be loaded.
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

    ' Changed from a fixed variable to a calculated property for dynamic row heights.
    Private ReadOnly Property RowHeight As Integer
        Get
            ' Relative height based on the main window size for better scaling.
            Return Math.Max(64, CInt(Me.Height / 8))
        End Get
    End Property
End Class