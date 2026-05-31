Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.Encodings.Web
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Globalization
Imports System.Windows.Forms

Public Class Form1
    Private Enum PdfGroupingMode
        AllHearings
        Day
        Week
        Month
        Year
    End Enum

    Private Enum PdfWeekRangeMode
        Automatic
        MondayToSunday
    End Enum

    Private Class PdfDateRange
        Public Property StartDate As Date
        Public Property EndDate As Date
    End Class

    Private Class PdfGroup
        Public Property Title As String
        Public Property Records As List(Of HearingRecord)
    End Class

    Private ReadOnly repository As New ExcelHearingRepository()
    Private ReadOnly webView As New WebView2()
    Private ReadOnly detailsPanel As New Panel()
    Private ReadOnly detailsTitleLabel As New Label()
    Private ReadOnly noValueLabel As New Label()
    Private ReadOnly nameValueLabel As New Label()
    Private ReadOnly courtValueLabel As New Label()
    Private ReadOnly hearing1ValueLabel As New Label()
    Private ReadOnly hearing2ValueLabel As New Label()
    Private ReadOnly dateValueLabel As New Label()
    Private ReadOnly workbookValueLabel As New Label()
    Private ReadOnly searchTextBox As New TextBox()
    Private ReadOnly hearingListBox As New ListBox()
    Private ReadOnly dateListTitleLabel As New Label()
    Private hearings As New List(Of HearingRecord)()
    Private selectedDate As Date = Date.Today
    Private dateWasClicked As Boolean = False
    ' Persists history log entries across ReloadCalendar() calls (keyed by Excel row Id)
    Private _historyCache As New Dictionary(Of Integer, List(Of String))()
    Private _lastExportedIds As New List(Of Integer)()
    Private ReadOnly _progressBar As New ProgressBar()
    Private ReadOnly _statusLabel As New Label()
    Private _progressPanel As Panel
    Private ReadOnly otherHearingsListBox As New ListBox()
    Private ReadOnly otherHearingsTitleLabel As New Label()
    Private _isUpdatingDisplay As Boolean = False

    Private Shared ReadOnly JsonOptions As New JsonSerializerOptions With {
        .Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        .PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    }

    Public Sub New()
        InitializeComponent()
        BuildLayout()
    End Sub

    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Await webView.EnsureCoreWebView2Async()
            AddHandler webView.CoreWebView2.WebMessageReceived, AddressOf CalendarMessageReceived

            Dim calendarPath = Path.Combine(AppContext.BaseDirectory, "Calendar", "calendar.html")
            webView.CoreWebView2.Navigate(New Uri(calendarPath).AbsoluteUri)
        Catch ex As Exception
            ShowError("Startup Error", "The calendar could not be initialized.", ex)
        End Try
    End Sub

    Private Sub BuildLayout()
        BackColor = Color.FromArgb(245, 243, 255)
        Font = New Font("Segoe UI", 10.0F)

        Dim mainLayout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1,
            .Padding = New Padding(0)
        }
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 370.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        webView.Dock = DockStyle.Fill
        webView.DefaultBackgroundColor = Color.FromArgb(245, 243, 255)

        ' Sidebar
        detailsPanel.Dock = DockStyle.Fill
        detailsPanel.BackColor = Color.FromArgb(250, 249, 255)
        detailsPanel.Padding = New Padding(0)

        ' Header strip
        Dim headerStrip As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .BackColor = Color.FromArgb(76, 29, 149)
        }
        Dim headerLbl As New Label With {
            .Text = "BJMP  HEARING PANEL",
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(254, 243, 199),
            .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Padding = New Padding(0, 0, 0, 2)
        }
        Dim goldBar As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 4,
            .BackColor = Color.FromArgb(251, 191, 36)
        }
        headerStrip.Controls.Add(headerLbl)
        headerStrip.Controls.Add(goldBar)

        ' Inner scroll panel
        Dim innerPanel As New Panel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .Padding = New Padding(16, 12, 16, 8)
        }

        ' Search section
        Dim searchSection As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 72,
            .Padding = New Padding(0, 0, 0, 10)
        }
        Dim searchLabel As New Label With {
            .Text = "SEARCH PDL",
            .Dock = DockStyle.Top,
            .Height = 20,
            .ForeColor = Color.FromArgb(107, 33, 168),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        }
        searchTextBox.Dock = DockStyle.Top
        searchTextBox.Height = 34
        searchTextBox.PlaceholderText = "Type name to search..."
        searchTextBox.BorderStyle = BorderStyle.FixedSingle
        searchTextBox.Font = New Font("Segoe UI", 10.0F)
        AddHandler searchTextBox.TextChanged, Sub() RefreshSideList()
        searchSection.Controls.Add(searchTextBox)
        searchSection.Controls.Add(searchLabel)

        ' Divider
        Dim div1 As New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(221, 214, 254)}

        ' Hearings list section
        Dim listSection As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 200,
            .Padding = New Padding(0, 8, 0, 8)
        }
        dateListTitleLabel.Dock = DockStyle.Top
        dateListTitleLabel.Height = 22
        dateListTitleLabel.Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        dateListTitleLabel.ForeColor = Color.FromArgb(107, 33, 168)
        dateListTitleLabel.Text = "SCHEDULED HEARINGS"

        hearingListBox.Dock = DockStyle.Fill
        hearingListBox.BorderStyle = BorderStyle.FixedSingle
        hearingListBox.Font = New Font("Segoe UI", 9.5F)
        hearingListBox.FormattingEnabled = True
        hearingListBox.BackColor = Color.White
        hearingListBox.ItemHeight = 22
        AddHandler hearingListBox.Format, AddressOf FormatHearingListItem
        AddHandler hearingListBox.SelectedIndexChanged,
            Sub()
                Dim item = TryCast(hearingListBox.SelectedItem, HearingRecord)
                If item IsNot Nothing Then
                    DisplayHearing(item)
                End If
            End Sub
        AddHandler hearingListBox.DoubleClick,
            Sub()
                Dim item = TryCast(hearingListBox.SelectedItem, HearingRecord)
                If item IsNot Nothing Then
                    ShowHearingDetailPopup(item)
                End If
            End Sub
        listSection.Controls.Add(hearingListBox)
        listSection.Controls.Add(dateListTitleLabel)

        ' Divider
        Dim div2 As New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(221, 214, 254)}

        ' Details card section
        Dim detailsCardLabel As New Label With {
            .Text = "SELECTED HEARING INFO",
            .Dock = DockStyle.Top,
            .Height = 22,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(107, 33, 168),
            .Padding = New Padding(0, 8, 0, 0)
        }

        detailsTitleLabel.Dock = DockStyle.Top
        detailsTitleLabel.AutoSize = True
        detailsTitleLabel.Font = New Font("Segoe UI Semibold", 12.0F)
        detailsTitleLabel.ForeColor = Color.FromArgb(76, 29, 149)
        detailsTitleLabel.Text = "-  No hearing selected"
        detailsTitleLabel.Padding = New Padding(0, 4, 0, 8)

        Dim detailsCard As New Panel With {
            .Dock = DockStyle.Top,
            .BackColor = Color.White,
            .Padding = New Padding(12, 10, 12, 10),
            .AutoSize = True
        }
        AddHandler detailsCard.Paint, Sub(s, ev)
                                          Dim rect = New Rectangle(0, 0, detailsCard.Width - 1, detailsCard.Height - 1)
                                          ev.Graphics.DrawRectangle(New System.Drawing.Pen(Color.FromArgb(221, 214, 254)), rect)
                                      End Sub

        Dim detailsTable As New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .ColumnCount = 2,
            .AutoSize = True
        }
        detailsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
        detailsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

        AddDetailRow(detailsTable, "Case No.", noValueLabel)
        AddDetailRow(detailsTable, "BR / Court", courtValueLabel)
        AddDetailRow(detailsTable, "Hearing", hearing1ValueLabel)
        AddDetailRow(detailsTable, "Hearing", hearing2ValueLabel)
        AddDetailRow(detailsTable, "Next Hearing", dateValueLabel)

        workbookValueLabel.Text = repository.WorkbookPath
        Dim fileSection As New FlowLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .Padding = New Padding(0, 8, 0, 0),
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True
        }
        Dim fileLbl As New Label With {
            .Text = "Data File: ",
            .AutoSize = True,
            .ForeColor = Color.FromArgb(139, 92, 246),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        }
        workbookValueLabel.AutoSize = True
        workbookValueLabel.Font = New Font("Segoe UI", 7.5F)
        workbookValueLabel.ForeColor = Color.FromArgb(139, 92, 246)
        fileSection.Controls.Add(fileLbl)
        fileSection.Controls.Add(workbookValueLabel)

        detailsCard.Controls.Add(detailsTable)
        detailsCard.Controls.Add(detailsTitleLabel)
        detailsCard.Controls.Add(detailsCardLabel)

        ' Other Hearings section
        Dim otherSection As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 110,
            .Padding = New Padding(0, 8, 0, 8)
        }
        otherHearingsTitleLabel.Dock = DockStyle.Top
        otherHearingsTitleLabel.Height = 20
        otherHearingsTitleLabel.Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        otherHearingsTitleLabel.ForeColor = Color.FromArgb(107, 33, 168)
        otherHearingsTitleLabel.Text = "ALL DATES FOR THIS PERSON"

        otherHearingsListBox.Dock = DockStyle.Fill
        otherHearingsListBox.BorderStyle = BorderStyle.FixedSingle
        otherHearingsListBox.Font = New Font("Segoe UI", 9.0F)
        otherHearingsListBox.BackColor = Color.White
        otherHearingsListBox.ItemHeight = 20
        otherHearingsListBox.FormattingEnabled = True
        AddHandler otherHearingsListBox.Format, Sub(s, ev)
                                                    Dim h = TryCast(ev.ListItem, HearingRecord)
                                                    If h IsNot Nothing Then
                                                        Dim dtStr = If(h.NextHearing = Date.MinValue, "Pending", h.NextHearing.ToString("yyyy-MM-dd"))
                                                        ev.Value = $"{dtStr} - {h.BrCourt} ({h.Hearing1})"
                                                    End If
                                                End Sub
        AddHandler otherHearingsListBox.SelectedIndexChanged, Sub()
                                                                  If _isUpdatingDisplay Then Return
                                                                  Dim item = TryCast(otherHearingsListBox.SelectedItem, HearingRecord)
                                                                  If item IsNot Nothing Then
                                                                      DisplayHearing(item)
                                                                  End If
                                                              End Sub
        otherSection.Controls.Add(otherHearingsListBox)
        otherSection.Controls.Add(otherHearingsTitleLabel)

        innerPanel.Controls.Add(fileSection)
        innerPanel.Controls.Add(otherSection)
        innerPanel.Controls.Add(detailsCard)
        innerPanel.Controls.Add(div2)
        innerPanel.Controls.Add(listSection)
        innerPanel.Controls.Add(div1)
        innerPanel.Controls.Add(searchSection)

        ' Button strip at bottom
        Dim btnStrip As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 252,
            .BackColor = Color.FromArgb(245, 243, 255),
            .Padding = New Padding(12, 8, 12, 10)
        }

        Dim addButton = MakeSideButton("+  Add New Hearing", Color.FromArgb(76, 29, 149), Color.White)
        addButton.Dock = DockStyle.Top
        addButton.Height = 38
        AddHandler addButton.Click, Sub() ShowAddDialog(selectedDate)

        Dim spacer1 As New Panel With {.Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent}

        Dim importButton = MakeSideButton("^  Import Excel / XML", Color.FromArgb(109, 40, 217), Color.White)
        importButton.Dock = DockStyle.Top
        importButton.Height = 36
        AddHandler importButton.Click, Sub() ImportDataFile()

        Dim spacer2 As New Panel With {.Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent}

        Dim exportButton = MakeSideButton("v  Export Hearings", Color.FromArgb(251, 191, 36), Color.FromArgb(76, 29, 149))
        exportButton.Dock = DockStyle.Top
        exportButton.Height = 36
        AddHandler exportButton.Click, Async Sub()
                                           Try
                                               Await ExportAllHearings()
                                           Catch ex As Exception
                                               MessageBox.Show(Me, "Failed to export: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                           End Try
                                       End Sub

        Dim spacer3 As New Panel With {.Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent}

        Dim pdfButton = MakeSideButton("PDF Export", Color.FromArgb(234, 179, 8), Color.FromArgb(76, 29, 149))
        pdfButton.Dock = DockStyle.Top
        pdfButton.Height = 36
        AddHandler pdfButton.Click, Async Sub()
                                         Try
                                             Await ExportPdfReportOnly()
                                         Catch ex As Exception
                                             MessageBox.Show(Me, "Failed to export PDF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                         End Try
                                     End Sub

        Dim spacerPdf As New Panel With {.Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent}

        Dim settingsButton = MakeSideButton("Settings / Clear Data", Color.FromArgb(109, 40, 217), Color.White)
        settingsButton.Dock = DockStyle.Top
        settingsButton.Height = 34
        AddHandler settingsButton.Click, Sub() ShowSettingsPopup()

        Dim spacer4 As New Panel With {.Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent}

        Dim refreshButton = MakeSideButton("Refresh Calendar", Color.FromArgb(221, 214, 254), Color.FromArgb(30, 27, 75))
        refreshButton.Dock = DockStyle.Top
        refreshButton.Height = 34
        AddHandler refreshButton.Click, Sub() ReloadCalendar()

        btnStrip.Controls.Add(refreshButton)
        btnStrip.Controls.Add(spacer4)
        btnStrip.Controls.Add(settingsButton)
        btnStrip.Controls.Add(spacer3)
        btnStrip.Controls.Add(pdfButton)
        btnStrip.Controls.Add(spacerPdf)
        btnStrip.Controls.Add(exportButton)
        btnStrip.Controls.Add(spacer2)
        btnStrip.Controls.Add(importButton)
        btnStrip.Controls.Add(spacer1)
        btnStrip.Controls.Add(addButton)

        ' Progress / status bar
        Dim progressPanel As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 44,
            .BackColor = Color.FromArgb(76, 29, 149),
            .Visible = False,
            .Padding = New Padding(12, 0, 12, 0)
        }
        _statusLabel.Dock = DockStyle.Fill
        _statusLabel.ForeColor = Color.White
        _statusLabel.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft
        _progressBar.Dock = DockStyle.Bottom
        _progressBar.Height = 6
        _progressBar.Style = ProgressBarStyle.Marquee
        _progressBar.MarqueeAnimationSpeed = 30
        progressPanel.Controls.Add(_statusLabel)
        progressPanel.Controls.Add(_progressBar)
        _progressPanel = progressPanel

        detailsPanel.Controls.Add(innerPanel)
        detailsPanel.Controls.Add(btnStrip)
        detailsPanel.Controls.Add(progressPanel)
        detailsPanel.Controls.Add(headerStrip)

        mainLayout.Controls.Add(webView, 0, 0)
        mainLayout.Controls.Add(detailsPanel, 1, 0)
        Controls.Add(mainLayout)
    End Sub

    Private Shared Function MakeSideButton(text As String, bgColor As Color, fgColor As Color) As Button
        Dim btn As New Button With {
            .Text = text,
            .BackColor = bgColor,
            .ForeColor = fgColor,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI Semibold", 9.0F),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(6, 0, 0, 0)
        }
        btn.FlatAppearance.BorderSize = 0
        Return btn
    End Function

    Private Shared Sub AddDetailRow(table As TableLayoutPanel, labelText As String, valueLabel As Label)
        Dim captionLbl As New Label With {
            .Text = labelText,
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(109, 40, 217),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(0, 6, 0, 6)
        }
        valueLabel.Text = "-"
        valueLabel.AutoSize = True
        valueLabel.Dock = DockStyle.Fill
        valueLabel.ForeColor = Color.FromArgb(76, 29, 149)
        valueLabel.Font = New Font("Segoe UI", 9.5F)
        valueLabel.TextAlign = ContentAlignment.MiddleLeft
        valueLabel.Padding = New Padding(0, 6, 0, 6)
        table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        table.Controls.Add(captionLbl)
        table.Controls.Add(valueLabel)
    End Sub

    Private Shared Function BuildCompactHearingMeta(hearing As HearingRecord) As String
        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(hearing.No) Then parts.Add(hearing.No.Trim())
        If Not String.IsNullOrWhiteSpace(hearing.BrCourt) Then parts.Add(hearing.BrCourt.Trim())
        If Not String.IsNullOrWhiteSpace(hearing.Hearing1) Then parts.Add(hearing.Hearing1.Trim())
        If Not String.IsNullOrWhiteSpace(hearing.Hearing2) Then parts.Add(hearing.Hearing2.Trim())
        Return String.Join(" " & ChrW(&H2022) & " ", parts)
    End Function

    Private Shared Function CreateRoundedPath(rect As Rectangle, radius As Integer) As System.Drawing.Drawing2D.GraphicsPath
        Dim path As New System.Drawing.Drawing2D.GraphicsPath()
        Dim d = radius * 2
        If d <= 0 Then
            path.AddRectangle(rect)
            path.CloseFigure()
            Return path
        End If

        path.StartFigure()
        path.AddArc(rect.X, rect.Y, d, d, 180, 90)
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    Private Sub ShowError(title As String, message As String, ex As Exception, Optional icon As MessageBoxIcon = MessageBoxIcon.Error)
        Dim details = ex.Message
        If ex.InnerException IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ex.InnerException.Message) Then
            details &= Environment.NewLine & Environment.NewLine & ex.InnerException.Message
        End If

        MessageBox.Show(Me, $"{message}{Environment.NewLine}{Environment.NewLine}{details}", title, MessageBoxButtons.OK, icon)
    End Sub

    Private Async Sub CalendarMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Dim messageJson As String = ""
        Try
            messageJson = e.WebMessageAsJson
        Catch ex As Exception
            ShowError("Calendar Update", "The calendar sent an unreadable message.", ex, MessageBoxIcon.Warning)
            Return
        End Try

        Await Task.Delay(20)

        Try
            Using document = JsonDocument.Parse(messageJson)
                Dim root = document.RootElement
                Dim actionElement As JsonElement
                If Not root.TryGetProperty("action", actionElement) Then
                    Return
                End If

                Dim action = actionElement.GetString()
                If String.IsNullOrWhiteSpace(action) Then
                    Return
                End If

                Select Case action
                    Case "ready"
                        ReloadCalendar()
                    Case "dateview"
                        Dim dateElement As JsonElement
                        If Not root.TryGetProperty("date", dateElement) Then
                            Throw New FormatException("The calendar did not provide a date for the selected day.")
                        End If

                        Dim clickedDate As Date
                        If Not Date.TryParse(dateElement.GetString(), clickedDate) Then
                            Throw New FormatException("The selected date from the calendar was not valid.")
                        End If

                        SelectDate(clickedDate)
                        ShowDatePopup(clickedDate)
                    Case "select"
                        Dim idElement As JsonElement
                        If Not root.TryGetProperty("id", idElement) Then
                            Return
                        End If

                        Dim idText = idElement.GetString()
                        SelectHearing(idText)
                    Case "detail"
                        Dim idElement As JsonElement
                        If Not root.TryGetProperty("id", idElement) Then
                            Return
                        End If

                        Dim idText = idElement.GetString()
                        Dim rowId As Integer
                        If Not Integer.TryParse(idText, rowId) Then
                            Return
                        End If

                        Dim hearing = hearings.FirstOrDefault(Function(item) item.Id = rowId)
                        If hearing Is Nothing Then
                            Return
                        End If

                        SelectHearing(idText)
                        ShowHearingDetailPopup(hearing)
                    Case "move"
                        Dim oldDateStr = ""
                        Dim oldDateElement As JsonElement
                        If root.TryGetProperty("oldDate", oldDateElement) Then
                            oldDateStr = oldDateElement.GetString()
                        End If

                        Dim idElement As JsonElement
                        If Not root.TryGetProperty("id", idElement) Then
                            Throw New FormatException("The calendar did not provide a record id for the move request.")
                        End If

                        Dim dateElement As JsonElement
                        If Not root.TryGetProperty("date", dateElement) Then
                            Throw New FormatException("The calendar did not provide a destination date for the move request.")
                        End If

                        Dim targetDate As Date
                        If Not Date.TryParse(dateElement.GetString(), targetDate) Then
                            Throw New FormatException("The destination date from the calendar was not valid.")
                        End If

                        Dim targetId = idElement.GetString()
                        MoveHearing(targetId, targetDate, oldDateStr)
                End Select
            End Using
        Catch ex As Exception
            ShowError("Calendar Update", "The calendar could not process that action.", ex, MessageBoxIcon.Warning)
            ReloadCalendar()
        End Try
    End Sub

    Private Async Sub ReloadCalendar()
        Try
            hearings = repository.LoadHearings()
            For Each h In hearings
                If _historyCache.ContainsKey(h.Id) Then
                    h.HistoryLog = _historyCache(h.Id)
                End If
            Next
            RefreshSideList()
            Dim calendarEvents = hearings.
                Where(Function(hearing) hearing.NextHearing <> Date.MinValue).
                Select(Function(hearing) New With {
                    .id = hearing.Id.ToString(),
                    .title = hearing.No,
                    .start = hearing.NextHearing.ToString("yyyy-MM-dd"),
                    .color = "#4C1D95",
                    .textColor = "#ffffff",
                    .extendedProps = New With {
                        .caseNo = hearing.No,
                        .name = hearing.NameOfPdl,
                        .hearing1 = hearing.Hearing1,
                        .hearing2 = hearing.Hearing2,
                        .court = hearing.BrCourt
                    }
                }).ToList()

            If webView.CoreWebView2 IsNot Nothing Then
                Await webView.CoreWebView2.ExecuteScriptAsync($"window.renderHearings({JsonSerializer.Serialize(calendarEvents, JsonOptions)});")
            End If
        Catch ex As Exception
            ShowError("Calendar Error", "Failed to reload the calendar.", ex, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub ShowAddDialog(selectedDate As Date)
        Try
            Using dialog As New AddEditHearingForm(selectedDate, hearings)
                If dialog.ShowDialog(Me) = DialogResult.OK Then
                    ShowProgress("Adding hearing...")
                    Try
                        Dim saved = repository.AddHearing(dialog.Hearing)
                        ReloadCalendar()
                        DisplayHearing(saved)
                    Finally
                        HideProgress()
                    End Try
                End If
            End Using
        Catch ex As Exception
            HideProgress()
            ShowError("Add Error", "Failed to add hearing.", ex)
        End Try
    End Sub

    Private Sub ShowDatePopup(clickedDate As Date)
        Dim shouldShow As Boolean = True
        While shouldShow
            shouldShow = False
            Dim allDateHearings = hearings.
                Where(Function(h) h.NextHearing.Date = clickedDate.Date).
                OrderBy(Function(h) h.No).
                ToList()

            Using popup As New Form()
                popup.Text = $"Hearings - {clickedDate:MMMM d, yyyy}"
                popup.StartPosition = FormStartPosition.CenterParent
                popup.FormBorderStyle = FormBorderStyle.FixedDialog
                popup.MinimizeBox = False
                popup.MaximizeBox = False
                popup.BackColor = Color.White
                popup.Font = New Font("Segoe UI", 10.0F)
                ' Extra height for the search bar (+54px)
                popup.ClientSize = New Size(520, Math.Min(134 + allDateHearings.Count * 44 + 60, 580))

                ' Title bar
                Dim titlePanel As New Panel With {
                    .Dock = DockStyle.Top,
                    .Height = 60,
                    .BackColor = Color.FromArgb(76, 29, 149),
                    .Padding = New Padding(18, 0, 18, 0)
                }
                Dim titleLbl As New Label With {
                    .Text = $"{clickedDate:dddd, MMMM d, yyyy}",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.White,
                    .Font = New Font("Segoe UI Semibold", 13.0F),
                    .TextAlign = ContentAlignment.MiddleLeft
                }
                Dim countLbl As New Label With {
                    .Text = $"{allDateHearings.Count} hearing(s)",
                    .Dock = DockStyle.Right,
                    .Width = 130,
                    .ForeColor = Color.FromArgb(251, 191, 36),
                    .Font = New Font("Segoe UI", 9.5F),
                    .TextAlign = ContentAlignment.MiddleRight
                }
                titlePanel.Controls.Add(titleLbl)
                titlePanel.Controls.Add(countLbl)

                ' Search bar
                Dim searchPanel As New Panel With {
                    .Dock = DockStyle.Top,
                    .Height = 46,
                    .BackColor = Color.FromArgb(245, 243, 255),
                    .Padding = New Padding(12, 8, 12, 6)
                }
                Dim popupSearchBox As New TextBox With {
                    .Dock = DockStyle.Fill,
                    .PlaceholderText = "Search by name or court...",
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Font = New Font("Segoe UI", 10.0F),
                    .BackColor = Color.White
                }
                Dim searchResultLbl As New Label With {
                    .Dock = DockStyle.Right,
                    .Width = 110,
                    .Text = $"{allDateHearings.Count} of {allDateHearings.Count}",
                    .ForeColor = Color.FromArgb(107, 33, 168),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleRight
                }
                searchPanel.Controls.Add(popupSearchBox)
                searchPanel.Controls.Add(searchResultLbl)

                ' Scroll panel for hearing rows
                Dim scroll As New Panel With {
                    .Dock = DockStyle.Fill,
                    .AutoScroll = True,
                    .Padding = New Padding(14, 8, 14, 8)
                }

                ' Local helper: build/rebuild hearing rows
                Dim BuildRows As Action(Of String) = Nothing
                BuildRows = Sub(filterText As String)
                                scroll.Controls.Clear()

                                Dim filtered As List(Of HearingRecord)
                                If String.IsNullOrWhiteSpace(filterText) Then
                                    filtered = allDateHearings
                                Else
                                    filtered = allDateHearings.Where(Function(h)
                                                                         Return h.NameOfPdl.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                                                                h.BrCourt.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                                                                h.No.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
                                                                     End Function).ToList()
                                End If

                                ' Update result count label
                                searchResultLbl.Text = $"{filtered.Count} of {allDateHearings.Count}"
                                searchResultLbl.ForeColor = If(filtered.Count = 0, Color.FromArgb(198, 40, 40), Color.FromArgb(107, 33, 168))

                                If filtered.Count = 0 Then
                                    Dim emptyLbl As New Label With {
                                        .Text = If(String.IsNullOrWhiteSpace(filterText),
                                                   "No hearings scheduled for this date.",
                                                   $"No results for ""{filterText}""."),
                                        .Dock = DockStyle.Top,
                                        .Height = 50,
                                        .ForeColor = Color.FromArgb(139, 92, 246),
                                        .TextAlign = ContentAlignment.MiddleCenter,
                                        .Font = New Font("Segoe UI", 10.0F, FontStyle.Italic)
                                    }
                                    scroll.Controls.Add(emptyLbl)
                                    Return
                                End If

                                ' Build rows in reverse order so DockStyle.Top stacks correctly
                                For i As Integer = filtered.Count - 1 To 0 Step -1
                                    Dim currentH = filtered(i)
                                    Dim row As New Panel With {
                                        .Dock = DockStyle.Top,
                                        .Height = 40,
                                        .BackColor = If(i Mod 2 = 0, Color.FromArgb(245, 248, 252), Color.White),
                                        .Cursor = Cursors.Hand,
                                        .Padding = New Padding(0, 1, 0, 1)
                                    }

                                    Dim statusDot As New Label With {
                                        .Location = New Point(10, 14),
                                        .Size = New Size(12, 12),
                                        .BackColor = Color.FromArgb(76, 29, 149)
                                    }
                                    Dim noLbl As New Label With {
                                        .Text = currentH.No,
                                        .Location = New Point(32, 0),
                                        .Size = New Size(38, 40),
                                        .TextAlign = ContentAlignment.MiddleCenter,
                                        .ForeColor = Color.FromArgb(107, 33, 168),
                                        .Font = New Font("Segoe UI", 9.0F)
                                    }
                                    Dim nameLbl As New Label With {
                                        .Text = currentH.NameOfPdl,
                                        .Location = New Point(72, 0),
                                        .Size = New Size(240, 40),
                                        .TextAlign = ContentAlignment.MiddleLeft,
                                        .ForeColor = Color.FromArgb(76, 29, 149),
                                        .Font = New Font("Segoe UI Semibold", 10.0F)
                                    }

                                    If Not String.IsNullOrWhiteSpace(filterText) AndAlso
                                       currentH.NameOfPdl.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                        nameLbl.ForeColor = Color.FromArgb(245, 158, 11)
                                    End If

                                    Dim courtLbl As New Label With {
                                        .Text = currentH.BrCourt,
                                        .Location = New Point(312, 0),
                                        .Size = New Size(92, 40),
                                        .TextAlign = ContentAlignment.MiddleLeft,
                                        .ForeColor = Color.FromArgb(91, 33, 182),
                                        .Font = New Font("Segoe UI", 9.5F)
                                    }
                                    Dim viewBtn As New Button With {
                                        .Text = "View",
                                        .Location = New Point(410, 7),
                                        .Size = New Size(55, 26),
                                        .BackColor = Color.FromArgb(76, 29, 149),
                                        .ForeColor = Color.White,
                                        .FlatStyle = FlatStyle.Flat,
                                        .Font = New Font("Segoe UI", 8.5F)
                                    }
                                    viewBtn.FlatAppearance.BorderSize = 0
                                    AddHandler viewBtn.Click, Sub()
                                                                  popup.Tag = currentH
                                                                  popup.DialogResult = DialogResult.OK
                                                                  popup.Close()
                                                              End Sub
                                    row.Controls.AddRange(New Control() {statusDot, noLbl, nameLbl, courtLbl, viewBtn})
                                    scroll.Controls.Add(row)
                                Next
                            End Sub

                ' Initial render (no filter)
                BuildRows("")

                ' Wire search box to rebuild rows on each keystroke
                AddHandler popupSearchBox.TextChanged, Sub()
                                                           BuildRows(popupSearchBox.Text.Trim())
                                                       End Sub

                ' Bottom buttons
                Dim btnPanel As New Panel With {
                    .Dock = DockStyle.Bottom,
                    .Height = 60,
                    .BackColor = Color.FromArgb(242, 246, 250),
                    .Padding = New Padding(14, 10, 14, 10)
                }
                Dim addBtn As New Button With {
                    .Text = "+ Add Hearing",
                    .Dock = DockStyle.Right,
                    .Width = 155,
                    .BackColor = Color.FromArgb(76, 29, 149),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI Semibold", 10.0F)
                }
                addBtn.FlatAppearance.BorderSize = 0
                AddHandler addBtn.Click, Sub()
                                             popup.Tag = "ADD"
                                             popup.DialogResult = DialogResult.OK
                                             popup.Close()
                                         End Sub
                Dim closeBtn As New Button With {
                    .Text = "Close",
                    .Dock = DockStyle.Right,
                    .Width = 96,
                    .BackColor = Color.FromArgb(91, 33, 182),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 10.0F),
                    .Margin = New Padding(0, 0, 8, 0)
                }
                closeBtn.FlatAppearance.BorderSize = 0
                AddHandler closeBtn.Click, Sub() popup.Close()
                btnPanel.Controls.Add(addBtn)
                btnPanel.Controls.Add(closeBtn)

                ' Stack controls: title (top) -> search bar (top) -> scroll (fill) -> buttons (bottom)
                popup.Controls.Add(scroll)
                popup.Controls.Add(searchPanel)
                popup.Controls.Add(btnPanel)
                popup.Controls.Add(titlePanel)

                ' Auto-focus search box so user can type immediately
                AddHandler popup.Shown, Sub() popupSearchBox.Focus()

                Dim result = popup.ShowDialog(Me)

                If result = DialogResult.OK Then
                    If TypeOf popup.Tag Is HearingRecord Then
                        Dim selectedHearing = DirectCast(popup.Tag, HearingRecord)
                        ShowHearingDetailPopup(selectedHearing)
                        shouldShow = True
                    ElseIf TypeOf popup.Tag Is String AndAlso DirectCast(popup.Tag, String) = "ADD" Then
                        ShowAddDialog(clickedDate)
                        shouldShow = True
                    End If
                End If
            End Using
        End While
    End Sub


    Private Sub SelectDate(dateValue As Date)
        selectedDate = dateValue.Date
        dateWasClicked = True
        detailsTitleLabel.Text = selectedDate.ToString("yyyy-MM-dd")
        RefreshSideList()
        If hearingListBox.Items.Count > 0 Then
            hearingListBox.SelectedIndex = 0
        End If
    End Sub

    Private Sub SelectHearing(idText As String)
        Dim rowId As Integer
        If Not Integer.TryParse(idText, rowId) Then
            Return
        End If

        Dim hearing = hearings.FirstOrDefault(Function(item) item.Id = rowId)
        If hearing IsNot Nothing Then
            NavigateToHearingDate(hearing.NextHearing.Date)
            DisplayHearing(hearing)
            RefreshSideList()
        End If
    End Sub

    Private Sub NavigateToHearingDate(targetDate As Date)
        selectedDate = targetDate.Date
        dateWasClicked = True
        detailsTitleLabel.Text = selectedDate.ToString("yyyy-MM-dd")
        RefreshSideList()

        If webView.CoreWebView2 IsNot Nothing Then
            webView.CoreWebView2.ExecuteScriptAsync($"if (window.focusCalendarDate) window.focusCalendarDate('{selectedDate:yyyy-MM-dd}');")
        End If
    End Sub

    Private Sub MoveHearing(idText As String, nextHearing As Date, Optional oldDateStr As String = "")
        Dim rowId As Integer
        If Not Integer.TryParse(idText, rowId) Then Return

        Dim originalHearing = hearings.FirstOrDefault(Function(item) item.Id = rowId)
        If originalHearing Is Nothing Then Return

        Try
            ShowProgress("Moving hearing...")

            Dim oldDate = originalHearing.NextHearing.Date
            If Not String.IsNullOrWhiteSpace(oldDateStr) Then
                Dim parsedOld As Date
                If Date.TryParse(oldDateStr, parsedOld) Then oldDate = parsedOld.Date
            End If

            Dim duplicatedHearing As New HearingRecord With {
                .No = originalHearing.No,
                .NameOfPdl = originalHearing.NameOfPdl,
                .BrCourt = originalHearing.BrCourt,
                .Hearing1 = originalHearing.Hearing1,
                .Hearing2 = originalHearing.Hearing2,
                .NextHearing = nextHearing.Date
            }

            Dim logEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} - Duplicated to {nextHearing.Date:MMMM d, yyyy} from original date {oldDate:MMMM d, yyyy}"
            originalHearing.HistoryLog.Add(logEntry)
            duplicatedHearing.HistoryLog.Add(logEntry)

            Dim saved = repository.AddHearing(duplicatedHearing)
            ReloadCalendar()

            Dim reloadedOriginal = hearings.FirstOrDefault(Function(item) item.Id = originalHearing.Id)
            If reloadedOriginal IsNot Nothing Then reloadedOriginal.HistoryLog = originalHearing.HistoryLog

            Dim reloadedDuplicate = hearings.FirstOrDefault(Function(item) item.Id = saved.Id)
            If reloadedDuplicate IsNot Nothing Then
                reloadedDuplicate.HistoryLog = duplicatedHearing.HistoryLog
                DisplayHearing(reloadedDuplicate)
            End If
        Catch ex As Exception
            ReloadCalendar()
            ShowError("Move Error", "Failed to move hearing.", ex)
        Finally
            HideProgress()
        End Try
    End Sub

    Private Sub DisplayHearing(hearing As HearingRecord)
        If _isUpdatingDisplay Then Return
        _isUpdatingDisplay = True
        Try
            detailsTitleLabel.Text = hearing.NameOfPdl
            noValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.No), "-", hearing.No)
            nameValueLabel.Text = hearing.NameOfPdl
            courtValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.BrCourt), "-", hearing.BrCourt)
            hearing1ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing1), "-", hearing.Hearing1)
            hearing2ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing2), "-", hearing.Hearing2)
            dateValueLabel.Text = If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("MMMM d, yyyy"))

            Dim samePersonHearings = hearings.Where(Function(h) String.Equals(h.NameOfPdl, hearing.NameOfPdl, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(h) h.NextHearing).ToList()

            otherHearingsListBox.BeginUpdate()
            otherHearingsListBox.Items.Clear()
            Dim selectIdx As Integer = -1
            For i As Integer = 0 To samePersonHearings.Count - 1
                Dim h = samePersonHearings(i)
                otherHearingsListBox.Items.Add(h)
                If h.Id = hearing.Id Then
                    selectIdx = i
                End If
            Next
            otherHearingsListBox.EndUpdate()

            If selectIdx >= 0 Then
                otherHearingsListBox.SelectedIndex = selectIdx
            End If

            Dim dates = samePersonHearings.Where(Function(h) h.NextHearing <> Date.MinValue).Select(Function(h) h.NextHearing.ToString("yyyy-MM-dd")).Distinct().ToList()
            Dim datesJson = JsonSerializer.Serialize(dates)
            If webView.CoreWebView2 IsNot Nothing Then
                webView.CoreWebView2.ExecuteScriptAsync($"if (window.highlightPdlHearings) window.highlightPdlHearings({datesJson});")
            End If
        Finally
            _isUpdatingDisplay = False
        End Try
    End Sub

    Private Sub ShowHearingDetailPopup(initialHearing As HearingRecord)
        Dim currentHearing As HearingRecord = initialHearing
        While currentHearing IsNot Nothing
            Dim hearing = currentHearing
            currentHearing = Nothing

            DisplayHearing(hearing)

            Dim historyList As New ListBox With {
                .Dock = DockStyle.Fill,
                .BorderStyle = BorderStyle.None,
                .BackColor = Color.FromArgb(250, 249, 255),
                .Font = New Font("Segoe UI", 8.5F),
                .ForeColor = Color.FromArgb(91, 33, 182),
                .ItemHeight = 20,
                .SelectionMode = SelectionMode.None
            }

            Dim otherHearingsList As New ListBox With {
                .Dock = DockStyle.Fill,
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = Color.White,
                .Font = New Font("Segoe UI", 9.0F),
                .ForeColor = Color.FromArgb(76, 29, 149),
                .ItemHeight = 18,
                .IntegralHeight = False,
                .FormattingEnabled = True
            }
            Dim isRefreshing As Boolean = False
            Dim refreshPopupView As Action(Of HearingRecord) = Nothing

            AddHandler otherHearingsList.Format, Sub(sender, ev)
                                                     Dim hearingItem = TryCast(ev.ListItem, HearingRecord)
                                                     If hearingItem Is Nothing Then Return

                                                     Dim dateText = If(hearingItem.NextHearing = Date.MinValue, "Pending", hearingItem.NextHearing.ToString("yyyy-MM-dd"))
                                                     Dim courtText = If(String.IsNullOrWhiteSpace(hearingItem.BrCourt), "No court", hearingItem.BrCourt.Trim())
                                                     Dim hearingText = If(String.IsNullOrWhiteSpace(hearingItem.Hearing1), "TRIAL", hearingItem.Hearing1.Trim())
                                                     ev.Value = $"{dateText} - {courtText} ({hearingText})"
                                                 End Sub

            Using dlg As New Form()
                dlg.Text = hearing.NameOfPdl
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MinimizeBox = False
                dlg.MaximizeBox = False
                dlg.BackColor = Color.White
                dlg.Font = New Font("Segoe UI", 10.0F)
                dlg.ClientSize = New Size(480, 680)

                Dim hdr As New Panel With {
                    .Height = 64,
                    .BackColor = Color.FromArgb(76, 29, 149),
                    .Margin = New Padding(0),
                    .Dock = DockStyle.Fill
                }
                Dim hdrName As New Label With {
                    .Text = hearing.NameOfPdl,
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.White,
                    .Font = New Font("Segoe UI Semibold", 13.0F),
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Padding = New Padding(0, 0, 0, 4)
                }
                Dim hdrGold As New Panel With {
                    .Dock = DockStyle.Bottom,
                    .Height = 4,
                    .BackColor = Color.FromArgb(76, 29, 149)
                }
                hdr.Controls.Add(hdrName)
                hdr.Controls.Add(hdrGold)

                Dim bodyPanel As New TableLayoutPanel With {
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 8,
                    .Padding = New Padding(24, 16, 24, 8),
                    .Margin = New Padding(0),
                    .AutoScroll = True,
                    .BackColor = Color.White
                }
                bodyPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 25))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 10))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 10))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

                Dim grid As New TableLayoutPanel With {
                    .Dock = DockStyle.Top,
                    .ColumnCount = 2,
                    .RowCount = 5,
                    .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                    .AutoSize = True,
                    .Margin = New Padding(0)
                }
                grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
                grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

                Dim cap0 As New Label With {.Text = "Case No.", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim val0 As New Label With {.Text = If(String.IsNullOrWhiteSpace(hearing.No), "-", hearing.No), .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(76, 29, 149), .Font = New Font("Segoe UI Semibold", 9.5F), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap0, 0, 0)
                grid.Controls.Add(val0, 1, 0)

                Dim cap1 As New Label With {.Text = "Name of PDL", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim nameText As New TextBox With {.Text = hearing.NameOfPdl, .Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F), .BorderStyle = BorderStyle.FixedSingle, .Margin = New Padding(0, 4, 0, 4)}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap1, 0, 1)
                grid.Controls.Add(nameText, 1, 1)

                Dim cap2 As New Label With {.Text = "BR / Court", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim courtText As New TextBox With {.Text = hearing.BrCourt, .Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F), .BorderStyle = BorderStyle.FixedSingle, .Margin = New Padding(0, 4, 0, 4)}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap2, 0, 2)
                grid.Controls.Add(courtText, 1, 2)

                Dim cap3 As New Label With {.Text = "Hearing", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim hearing1Text As New TextBox With {.Text = hearing.Hearing1, .Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F), .BorderStyle = BorderStyle.FixedSingle, .Margin = New Padding(0, 4, 0, 4)}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap3, 0, 3)
                grid.Controls.Add(hearing1Text, 1, 3)

                Dim cap4 As New Label With {.Text = "Hearing", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim hearing2Text As New TextBox With {.Text = hearing.Hearing2, .Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F), .BorderStyle = BorderStyle.FixedSingle, .Margin = New Padding(0, 4, 0, 4)}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap4, 0, 4)
                grid.Controls.Add(hearing2Text, 1, 4)

                Dim cap5 As New Label With {.Text = "Next Hearing", .Dock = DockStyle.Fill, .ForeColor = Color.FromArgb(109, 40, 217), .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(0, 6, 0, 6), .AutoSize = True}
                Dim datePicker As New DateTimePicker With {.Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F), .Format = DateTimePickerFormat.Custom, .CustomFormat = "MMMM d, yyyy", .Value = If(hearing.NextHearing = Date.MinValue, Date.Today, hearing.NextHearing.Date), .Margin = New Padding(0, 4, 0, 4)}
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap5, 0, 5)
                grid.Controls.Add(datePicker, 1, 5)

                Dim detailWarningLabel As New Label With {
                    .ForeColor = Color.FromArgb(198, 40, 40),
                    .Font = New Font("Segoe UI Semibold", 8.5F),
                    .Text = "",
                    .Visible = False,
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(0, 4, 0, 4),
                    .AutoSize = True
                }

                refreshPopupView = Sub(targetHearing As HearingRecord)
                                       If targetHearing Is Nothing Then Return

                                       isRefreshing = True
                                       Try
                                           hearing = targetHearing
                                           DisplayHearing(targetHearing)

                                           dlg.Text = targetHearing.NameOfPdl
                                           hdrName.Text = targetHearing.NameOfPdl
                                           val0.Text = If(String.IsNullOrWhiteSpace(targetHearing.No), "-", targetHearing.No)
                                           nameText.Text = targetHearing.NameOfPdl
                                           courtText.Text = targetHearing.BrCourt
                                           hearing1Text.Text = targetHearing.Hearing1
                                           hearing2Text.Text = targetHearing.Hearing2
                                           If targetHearing.NextHearing <> Date.MinValue Then
                                               datePicker.Value = targetHearing.NextHearing.Date
                                           Else
                                               datePicker.Value = Date.Today
                                           End If

                                           otherHearingsList.BeginUpdate()
                                           otherHearingsList.Items.Clear()
                                           Dim samePersonHearingsList = hearings.Where(Function(h) String.Equals(h.NameOfPdl, targetHearing.NameOfPdl, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(h) h.NextHearing).ToList()
                                           Dim currentIdx As Integer = -1
                                           For i = 0 To samePersonHearingsList.Count - 1
                                               Dim h = samePersonHearingsList(i)
                                               otherHearingsList.Items.Add(h)
                                               If h.Id = targetHearing.Id Then currentIdx = i
                                           Next
                                           otherHearingsList.EndUpdate()

                                           If currentIdx >= 0 Then
                                               otherHearingsList.SelectedIndex = currentIdx
                                           ElseIf otherHearingsList.Items.Count > 0 Then
                                               otherHearingsList.SelectedIndex = 0
                                           End If

                                           historyList.Items.Clear()
                                           If targetHearing.HistoryLog IsNot Nothing Then
                                               For i = targetHearing.HistoryLog.Count - 1 To 0 Step -1
                                                   historyList.Items.Add(targetHearing.HistoryLog(i))
                                               Next
                                           End If

                                           detailWarningLabel.Visible = False
                                       Finally
                                           isRefreshing = False
                                       End Try
                                   End Sub

                Dim checkDetailDuplicate = Sub()
                                               If isRefreshing Then Return
                                               Dim targetDate = datePicker.Value.Date
                                               If hearings Is Nothing OrElse hearings.Count = 0 Then
                                                   detailWarningLabel.Visible = False
                                                   Return
                                               End If
                                               Dim hasDup = hearings.Any(Function(h)
                                                                             Return h.Id <> hearing.Id AndAlso
                                                                                    String.Equals(h.NameOfPdl.Trim(), nameText.Text.Trim(), StringComparison.OrdinalIgnoreCase) AndAlso
                                                                                    h.NextHearing.Date = targetDate
                                                                         End Function)
                                               If hasDup Then
                                                   detailWarningLabel.Text = "Warning: This person already has a hearing scheduled on this date!"
                                                   detailWarningLabel.Visible = True
                                               Else
                                                   detailWarningLabel.Visible = False
                                           End If
                                       End Sub

                AddHandler datePicker.ValueChanged, Sub() checkDetailDuplicate()
                AddHandler nameText.TextChanged, Sub() checkDetailDuplicate()
                checkDetailDuplicate()
                AddHandler otherHearingsList.DoubleClick, Sub()
                                                              Dim selected = TryCast(otherHearingsList.SelectedItem, HearingRecord)
                                                              If selected Is Nothing OrElse selected.Id = hearing.Id Then Return
                                                              refreshPopupView(selected)
                                                          End Sub

                Dim divLine As New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(221, 214, 254), .Margin = New Padding(0, 12, 0, 12)}

                Dim btnPanel As New Panel With {.Dock = DockStyle.Top, .Height = 36, .Margin = New Padding(0)}

                Dim deleteBtn As New Button With {.Text = "Delete Hearing", .Location = New Point(0, 0), .Size = New Size(115, 34), .BackColor = Color.FromArgb(198, 40, 40), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI Semibold", 9.0F)}
                deleteBtn.FlatAppearance.BorderSize = 0
                AddHandler deleteBtn.Click, Sub()
                                                Dim confirm = MessageBox.Show(dlg, $"Are you sure you want to delete the hearing schedule for {hearing.NameOfPdl}?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                                If confirm = DialogResult.Yes Then
                                                    Try
                                                        repository.DeleteHearing(hearing.Id)
                                                        ReloadCalendar()
                                                        detailsTitleLabel.Text = "-  No hearing selected"
                                                        noValueLabel.Text = "-"
                                                        nameValueLabel.Text = "-"
                                                        courtValueLabel.Text = "-"
                                                        hearing1ValueLabel.Text = "-"
                                                        hearing2ValueLabel.Text = "-"
                                                        dateValueLabel.Text = "-"
                                                        MessageBox.Show(dlg, "Hearing deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                        dlg.Close()
                                                    Catch ex As Exception
                                                        MessageBox.Show(dlg, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                    End Try
                                                End If
                                            End Sub

                Dim saveBtn As New Button With {.Text = "Save Changes", .Location = New Point(200, 0), .Size = New Size(130, 34), .BackColor = Color.FromArgb(76, 29, 149), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI Semibold", 9.0F)}
                saveBtn.FlatAppearance.BorderSize = 0
                AddHandler saveBtn.Click, Sub()
                                              saveBtn.Enabled = False
                                              Try
                                                  Dim originalName = hearing.NameOfPdl
                                                  hearing.No = val0.Text.Trim()
                                                  hearing.NameOfPdl = nameText.Text.Trim()
                                                  hearing.BrCourt = courtText.Text.Trim()
                                                  hearing.Hearing1 = hearing1Text.Text.Trim()
                                                  hearing.Hearing2 = hearing2Text.Text.Trim()
                                                  dlg.Text = hearing.NameOfPdl
                                                  hdrName.Text = hearing.NameOfPdl
                                                  Dim nameChanged = Not String.Equals(originalName.Trim(), hearing.NameOfPdl.Trim(), StringComparison.OrdinalIgnoreCase)
                                                  Dim originalDate = hearing.NextHearing.Date
                                                  Dim selectedDateVal = datePicker.Value.Date
                                                  If originalDate <> selectedDateVal Then
                                                      repository.UpdateHearing(hearing)
                                                      Dim logEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} - Duplicated to {selectedDateVal:MMMM d, yyyy} from original date {originalDate:MMMM d, yyyy}"
                                                      If Not _historyCache.ContainsKey(hearing.Id) Then _historyCache(hearing.Id) = New List(Of String)()
                                                      _historyCache(hearing.Id).Add(logEntry)
                                                      hearing.HistoryLog = _historyCache(hearing.Id)
                                                      Dim duplicatedHearing As New HearingRecord With {.No = hearing.No, .NameOfPdl = hearing.NameOfPdl, .BrCourt = hearing.BrCourt, .Hearing1 = hearing.Hearing1, .Hearing2 = hearing.Hearing2, .NextHearing = selectedDateVal}
                                                      repository.AddHearing(duplicatedHearing)
                                                      Dim renamedCount = If(nameChanged, repository.RenameHearings(originalName, hearing.NameOfPdl), 0)
                                                      Dim dupLogEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} - Created from {originalDate:MMMM d, yyyy}"
                                                      _historyCache(duplicatedHearing.Id) = New List(Of String)() From {dupLogEntry}
                                                      duplicatedHearing.HistoryLog = _historyCache(duplicatedHearing.Id)
                                                      ReloadCalendar()
                                                      DisplayHearing(hearing)
                                                      If historyList IsNot Nothing Then
                                                          historyList.Items.Clear()
                                                          For i = hearing.HistoryLog.Count - 1 To 0 Step -1
                                                              historyList.Items.Add(hearing.HistoryLog(i))
                                                          Next
                                                      End If
                                                      saveBtn.Enabled = True
                                                      Dim totalAffected = If(nameChanged, renamedCount + 2, 0)
                                                      Dim renameNote = If(totalAffected > 1, $" The name was updated across {totalAffected} hearings for this person.", If(totalAffected = 1, " The name was updated across this hearing.", ""))
                                                      MessageBox.Show(dlg, $"Hearing duplicated to {selectedDateVal:yyyy-MM-dd}. The original hearing has been updated and remains on {originalDate:yyyy-MM-dd} to preserve history.{renameNote}", "Saved & Duplicated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                  Else
                                                      hearing.NextHearing = selectedDateVal
                                                      repository.UpdateHearing(hearing)
                                                      Dim renamedCount = If(nameChanged, repository.RenameHearings(originalName, hearing.NameOfPdl), 0)
                                                      ReloadCalendar()
                                                      DisplayHearing(hearing)
                                                      saveBtn.Enabled = True
                                                      Dim totalAffected = If(nameChanged, renamedCount + 1, 0)
                                                      Dim renameNote = If(totalAffected > 1, $" The name was updated across {totalAffected} hearings for this person.", If(totalAffected = 1, " The name was updated across this hearing.", ""))
                                                      MessageBox.Show(dlg, $"Hearing changes saved successfully.{renameNote}", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                  End If
                                              Catch ex As Exception
                                                  saveBtn.Enabled = True
                                                  MessageBox.Show(dlg, ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                              End Try
                                          End Sub

                Dim closeBtn As New Button With {.Text = "Close", .Location = New Point(340, 0), .Size = New Size(76, 34), .BackColor = Color.FromArgb(237, 233, 254), .ForeColor = Color.FromArgb(76, 29, 149), .FlatStyle = FlatStyle.Flat, .Font = New Font("Segoe UI Semibold", 9.0F)}
                closeBtn.FlatAppearance.BorderSize = 0
                closeBtn.FlatAppearance.BorderColor = Color.FromArgb(196, 181, 253)
                AddHandler closeBtn.Click, Sub() dlg.Close()

                btnPanel.Controls.Add(deleteBtn)
                btnPanel.Controls.Add(saveBtn)
                btnPanel.Controls.Add(closeBtn)

                bodyPanel.Controls.Add(grid, 0, 0)
                bodyPanel.Controls.Add(detailWarningLabel, 0, 1)
                bodyPanel.Controls.Add(divLine, 0, 2)
                bodyPanel.Controls.Add(btnPanel, 0, 3)

                Dim otherSpacer As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.Transparent}
                bodyPanel.Controls.Add(otherSpacer, 0, 4)

                Dim otherContainer As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(250, 249, 255), .Padding = New Padding(10, 8, 10, 8), .Margin = New Padding(0)}
                AddHandler otherContainer.Paint, Sub(s, ev)
                                                     Dim rect = New Rectangle(0, 0, otherContainer.Width - 1, otherContainer.Height - 1)
                                                     ev.Graphics.DrawRectangle(New Pen(Color.FromArgb(221, 214, 254)), rect)
                                                 End Sub

                Dim otherTitle As New Label With {.Text = "ALL SCHEDULED HEARINGS FOR THIS PERSON", .Dock = DockStyle.Top, .Height = 22, .ForeColor = Color.FromArgb(107, 33, 168), .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold), .Padding = New Padding(0, 0, 0, 4)}

                AddHandler otherHearingsList.Format, Sub(s, ev)
                                                         Dim h = TryCast(ev.ListItem, HearingRecord)
                                                         If h IsNot Nothing Then
                                                             Dim dtStr = If(h.NextHearing = Date.MinValue, "Pending", h.NextHearing.ToString("yyyy-MM-dd"))
                                                             ev.Value = $"{dtStr} - {h.BrCourt} ({h.Hearing1})"
                                                         End If
                                                     End Sub

                otherContainer.Controls.Add(otherHearingsList)
                otherContainer.Controls.Add(otherTitle)
                bodyPanel.Controls.Add(otherContainer, 0, 5)

                If hearing.HistoryLog IsNot Nothing Then
                    Dim historySpacer As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.Transparent}
                    bodyPanel.Controls.Add(historySpacer, 0, 6)

                    Dim historyContainer As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(250, 249, 255), .Padding = New Padding(10, 8, 10, 8), .Margin = New Padding(0)}
                    AddHandler historyContainer.Paint, Sub(s, ev)
                                                           Dim rect = New Rectangle(0, 0, historyContainer.Width - 1, historyContainer.Height - 1)
                                                           ev.Graphics.DrawRectangle(New Pen(Color.FromArgb(221, 214, 254)), rect)
                                                       End Sub

                    Dim historyTitle As New Label With {.Text = "HISTORY LOG", .Dock = DockStyle.Top, .Height = 22, .ForeColor = Color.FromArgb(107, 33, 168), .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold), .Padding = New Padding(0, 0, 0, 4)}

                    historyList.Items.Clear()

                    historyContainer.Controls.Add(historyList)
                    historyContainer.Controls.Add(historyTitle)
                    bodyPanel.Controls.Add(historyContainer, 0, 7)
                End If

                refreshPopupView(hearing)

                Dim mainLayout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(0), .Margin = New Padding(0)}
                mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 64))
                mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
                mainLayout.Controls.Add(hdr, 0, 0)
                mainLayout.Controls.Add(bodyPanel, 0, 1)

                dlg.Controls.Add(mainLayout)
                dlg.ShowDialog(Me)
            End Using
        End While
    End Sub


    Private Sub RefreshSideList()
        Dim searchText = searchTextBox.Text.Trim()
        Dim filtered As List(Of HearingRecord)

        If Not String.IsNullOrWhiteSpace(searchText) Then
            Dim allMatches = hearings.Where(Function(h)
                                                Return h.NameOfPdl.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0
                                            End Function)
            filtered = allMatches.
                OrderBy(Function(h) If(h.NextHearing = Date.MinValue, Date.MaxValue, h.NextHearing.Date)).
                ThenBy(Function(h) h.NameOfPdl).
                ThenBy(Function(h) h.No).
                ToList()
            Dim uniqueCount = filtered.Count
            dateListTitleLabel.Text = $"Search results - {uniqueCount} hearing(s)"

        ElseIf dateWasClicked Then
            filtered = hearings.Where(Function(h)
                                          Return h.NextHearing.Date = selectedDate.Date
                                      End Function).
                OrderBy(Function(h) h.NameOfPdl).
                ToList()
            dateListTitleLabel.Text = $"Hearings on {selectedDate:yyyy-MM-dd} ({filtered.Count})"

        Else
            filtered = hearings.
                OrderBy(Function(h) If(h.NextHearing = Date.MinValue, 1, 0)).
                ThenByDescending(Function(h) h.NextHearing).
                ThenBy(Function(h) h.NameOfPdl).
                ToList()
            Dim pastCount = filtered.Where(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date < Date.Today).Count()
            Dim upcomingCount = filtered.Where(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date >= Date.Today).Count()
            dateListTitleLabel.Text = $"All Hearings - {upcomingCount} upcoming, {pastCount} past"
        End If

        hearingListBox.BeginUpdate()
        hearingListBox.Items.Clear()
        For Each h In filtered
            hearingListBox.Items.Add(h)
        Next
        hearingListBox.EndUpdate()

        If filtered.Count > 0 AndAlso Not String.IsNullOrWhiteSpace(searchText) Then
            hearingListBox.SelectedIndex = 0
            DisplayHearing(filtered(0))
        End If
    End Sub


    Private Sub FormatHearingListItem(sender As Object, e As ListControlConvertEventArgs)
        Dim hearing = TryCast(e.ListItem, HearingRecord)
        If hearing IsNot Nothing Then
            Dim isSearching = Not String.IsNullOrWhiteSpace(searchTextBox.Text.Trim())
            If isSearching Then
                e.Value = $"{hearing.No} - {hearing.NameOfPdl}"
            Else
                Dim dateStr = If(hearing.NextHearing = Date.MinValue, "Pending", hearing.NextHearing.ToString("yyyy-MM-dd"))
                e.Value = $"{hearing.No} - {hearing.NameOfPdl} [{dateStr}]"
            End If
        End If
    End Sub

    Private Sub ImportDataFile()
        Dim backupPath As String = ""
        Using saveDlg As New SaveFileDialog()
            saveDlg.Title = "Save backup of current data before importing"
            saveDlg.Filter = "XML files (*.xml)|*.xml"
            saveDlg.FileName = $"hearings_backup_{DateTime.Now:yyyy-MM-dd}"
            saveDlg.DefaultExt = "xml"
            saveDlg.OverwritePrompt = True
            Dim saveResult = saveDlg.ShowDialog(Me)
            If saveResult = DialogResult.Cancel Then Return
            If saveResult = DialogResult.OK Then backupPath = saveDlg.FileName
        End Using

        Using openDlg As New OpenFileDialog()
            openDlg.Title = "Import hearing data"
            openDlg.Filter = "Excel or XML files (*.xlsx;*.xlsm;*.xml)|*.xlsx;*.xlsm;*.xml"
            If openDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Try
                ShowProgress("Backing up current data...")
                If Not String.IsNullOrEmpty(backupPath) Then repository.BackupCurrentData(backupPath)
                ShowProgress("Importing hearing data...")
                repository.ImportDataFile(openDlg.FileName)
                workbookValueLabel.Text = repository.WorkbookPath
                ShowProgress("Refreshing calendar...")
                ReloadCalendar()
                Dim scheduledCount = repository.CountSchedulableHearings()
                Dim backupNote As String = If(Not String.IsNullOrEmpty(backupPath), $"{Environment.NewLine}{Environment.NewLine}Previous data backed up to:{Environment.NewLine}  {backupPath}", "")
                HideProgress()
                MessageBox.Show(Me, $"Import complete! {scheduledCount} hearing(s) now in the master file." & backupNote, "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                HideProgress()
                ShowError("Import Error", "Import failed.", ex)
            End Try
        End Using
    End Sub

    Private Async Function ExportAllHearings() As Task
        Using dialog As New SaveFileDialog()
            dialog.Title = "Export Master Hearings"
            dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV UTF-8 (*.csv)|*.csv|PDF Document (*.pdf)|*.pdf"
            dialog.FileName = $"BJMP-Hearings-All-{DateTime.Now:yyyy-MM-dd}"
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Dim ext = Path.GetExtension(dialog.FileName).ToLowerInvariant()
            Try
                ShowProgress("Loading hearing data...")
                Dim currentHearings = repository.LoadHearings()

                Select Case ext
                    Case ".xlsx"
                        ShowProgress("Exporting to Excel...")
                        Dim exported = currentHearings.Where(Function(h) h.NextHearing <> Date.MinValue).ToList()
                        _lastExportedIds.Clear()
                        _lastExportedIds.AddRange(exported.Select(Function(h) h.Id))
                        repository.ExportToExcel(dialog.FileName)
                        HideProgress()
                        MessageBox.Show(Me, "Excel exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case ".csv"
                        ShowProgress("Exporting to CSV...")
                        _lastExportedIds.Clear()
                        _lastExportedIds.AddRange(currentHearings.Select(Function(h) h.Id))
                        repository.ExportToCsv(dialog.FileName)
                        HideProgress()
                        MessageBox.Show(Me, "CSV exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case ".pdf"
                        ShowProgress("Generating PDF report from Word template...")
                        Dim pdfMode = PromptPdfGroupingMode()
                        If pdfMode Is Nothing Then
                            HideProgress()
                            Return
                        End If

                        Dim weekMode As Nullable(Of PdfWeekRangeMode) = Nothing
                        Dim weekRange As PdfDateRange = Nothing
                        If pdfMode.Value = PdfGroupingMode.Week Then
                            weekMode = PromptPdfWeekRangeMode()
                            If weekMode Is Nothing Then
                                HideProgress()
                                Return
                            End If

                            weekRange = PromptPdfDateRange()
                            If weekRange Is Nothing Then
                                HideProgress()
                                Return
                            End If
                        End If

                        If Not PromptPdfExportConfirm(pdfMode.Value, weekMode, weekRange) Then
                            HideProgress()
                            Return
                        End If

                        Dim allHearings = repository.LoadHearings().OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                        _lastExportedIds.Clear()
                        _lastExportedIds.AddRange(allHearings.Select(Function(h) h.Id))
                        Await ExportHearingsToPdf(dialog.FileName, allHearings, pdfMode.Value, If(weekMode.HasValue, weekMode.Value, PdfWeekRangeMode.MondayToSunday), weekRange)
                        HideProgress()
                        MessageBox.Show(Me, "PDF exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Select
            Catch ex As Exception
                HideProgress()
                ShowError("Export Error", "Export failed.", ex)
            End Try
        End Using
    End Function

    Private Async Function ExportPdfReportOnly() As Task
        Using dialog As New SaveFileDialog()
            dialog.Title = "Export PDF Report"
            dialog.Filter = "PDF Document (*.pdf)|*.pdf"
            dialog.DefaultExt = "pdf"
            dialog.FileName = $"BJMP-Hearings-{DateTime.Now:yyyy-MM-dd}.pdf"
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Try
                ShowProgress("Loading hearing data...")
                Dim allHearings = repository.LoadHearings().OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                Dim pdfMode = PromptPdfGroupingMode()
                If pdfMode Is Nothing Then
                    HideProgress()
                    Return
                End If

                Dim weekMode As Nullable(Of PdfWeekRangeMode) = Nothing
                Dim weekRange As PdfDateRange = Nothing
                If pdfMode.Value = PdfGroupingMode.Week Then
                    weekMode = PromptPdfWeekRangeMode()
                    If weekMode Is Nothing Then
                        HideProgress()
                        Return
                    End If

                    weekRange = PromptPdfDateRange()
                    If weekRange Is Nothing Then
                        HideProgress()
                        Return
                    End If
                End If

                If Not PromptPdfExportConfirm(pdfMode.Value, weekMode, weekRange) Then
                    HideProgress()
                    Return
                End If

                _lastExportedIds.Clear()
                _lastExportedIds.AddRange(allHearings.Select(Function(h) h.Id))

                ShowProgress("Generating PDF report from Word template...")
                Await ExportHearingsToPdf(dialog.FileName, allHearings, pdfMode.Value, If(weekMode.HasValue, weekMode.Value, PdfWeekRangeMode.MondayToSunday), weekRange)
                HideProgress()
                MessageBox.Show(Me, "PDF exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                HideProgress()
                ShowError("Export Error", "PDF export failed.", ex)
            End Try
        End Using
    End Function

    Private Function ExportHearingsToPdf(outputPdfPath As String, records As List(Of HearingRecord), mode As PdfGroupingMode, weekMode As PdfWeekRangeMode, weekRange As PdfDateRange) As Task
        Dim templatePath = GetBundledTemplatePath()
        If Not File.Exists(templatePath) Then
            Throw New FileNotFoundException("The bundled Court Calendar template could not be found. PDF export cannot continue.", templatePath)
        End If

        Dim wordApp As Object = Nothing
        Dim document As Object = Nothing
        Dim hearingTable As Object = Nothing
        Dim workingCopyPath As String = Nothing

        Try
            Try
                wordApp = CreateObject("Word.Application")
            Catch ex As Exception
                Throw New InvalidOperationException("Microsoft Word is required to export the PDF template.", ex)
            End Try

            wordApp.Visible = False
            wordApp.DisplayAlerts = 0

            workingCopyPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}_CourtCalendarTemplate.docx")
            File.Copy(templatePath, workingCopyPath, True)
            File.SetAttributes(workingCopyPath, FileAttributes.Normal)
            document = wordApp.Documents.Open(workingCopyPath, False, False, False)
            Dim anchorParagraph As Object = FindParagraphContainingText(document, "COURT CALENDAR")
            Dim anchorParagraphIndex As Integer = FindParagraphIndexContainingText(document, "COURT CALENDAR")
            If anchorParagraph Is Nothing Then
                anchorParagraph = FindFirstBodyParagraph(document)
            End If
            If anchorParagraph Is Nothing Then
                Throw New InvalidOperationException("The Word template does not contain a usable title paragraph.")
            End If

            Dim insertRange As Object = Nothing
            If anchorParagraphIndex > 0 AndAlso anchorParagraphIndex < document.Paragraphs.Count Then
                Dim insertParagraph As Object = document.Paragraphs(anchorParagraphIndex + 1)
                insertRange = insertParagraph.Range.Duplicate
                insertRange.Collapse(0)
            Else
                ' Fallback to the title paragraph when the template structure is not what we expect.
                insertRange = anchorParagraph.Range.Duplicate
                insertRange.Collapse(0)
            End If

            Dim groups = BuildPdfGroups(records, mode, weekMode, weekRange)
            If mode = PdfGroupingMode.AllHearings Then
                Dim flattenedRecords = groups.SelectMany(Function(g) g.Records).ToList()
                Dim rowCount As Integer = Math.Max(flattenedRecords.Count + 1, 2)
                hearingTable = document.Tables.Add(insertRange, rowCount, 6)
                FormatHearingTable(hearingTable, flattenedRecords, document)
            Else
                Dim groupedRowCount As Integer = 1
                For Each group In groups
                    groupedRowCount += 2
                    groupedRowCount += Math.Max(group.Records.Count, 1)
                Next
                hearingTable = document.Tables.Add(insertRange, Math.Max(groupedRowCount, 2), 6)
                FormatGroupedHearingTable(hearingTable, groups, document, mode)
            End If

            If File.Exists(outputPdfPath) Then
                Try
                    File.SetAttributes(outputPdfPath, FileAttributes.Normal)
                    File.Delete(outputPdfPath)
                Catch
                End Try
            End If

            document.ExportAsFixedFormat(outputPdfPath, 17)
            Return Task.CompletedTask
        Finally
            If document IsNot Nothing Then
                Try
                    document.Close(0)
                Catch
                End Try
            End If

            If wordApp IsNot Nothing Then
                Try
                    wordApp.Quit(0)
                Catch
                End Try
            End If

            ReleaseComObjectSafe(hearingTable)
            ReleaseComObjectSafe(document)
            ReleaseComObjectSafe(wordApp)

            If Not String.IsNullOrWhiteSpace(workingCopyPath) AndAlso File.Exists(workingCopyPath) Then
                Try
                    File.Delete(workingCopyPath)
                Catch
                End Try
            End If
        End Try
    End Function

    Private Function PromptPdfGroupingMode() As Nullable(Of PdfGroupingMode)
        Using dialog As New Form()
            dialog.Text = "PDF Export Type"
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ClientSize = New Size(320, 210)
            dialog.BackColor = Color.White
            dialog.Font = New Font("Segoe UI", 9.5F)

            Dim titleLbl As New Label With {
                .Text = "Choose how the PDF should be grouped:",
                .Dock = DockStyle.Top,
                .Height = 32,
                .Padding = New Padding(14, 10, 14, 0),
                .ForeColor = Color.FromArgb(76, 29, 149),
                .Font = New Font("Segoe UI Semibold", 10.0F)
            }

            Dim choicePanel As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 128,
                .Padding = New Padding(14, 0, 14, 0)
            }

            Dim modeCombo As New ComboBox With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Font = New Font("Segoe UI", 9.5F)
            }
            modeCombo.Items.AddRange(New Object() {
                "All hearings",
                "Per day",
                "Per week",
                "Per month",
                "Per year"
            })
            modeCombo.SelectedIndex = 1

            Dim hintLbl As New Label With {
                .Text = "Tip: choose a grouped report if you want separate date sections.",
                .Dock = DockStyle.Top,
                .Height = 40,
                .Padding = New Padding(0, 10, 0, 0),
                .ForeColor = Color.FromArgb(107, 33, 168),
                .Font = New Font("Segoe UI", 8.5F)
            }

            choicePanel.Controls.Add(hintLbl)
            choicePanel.Controls.Add(modeCombo)

            Dim buttonPanel As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 58,
                .Padding = New Padding(14, 10, 14, 10),
                .BackColor = Color.FromArgb(245, 243, 255)
            }

            Dim okBtn As New Button With {
                .Text = "Continue",
                .Dock = DockStyle.Right,
                .Width = 100,
                .BackColor = Color.FromArgb(76, 29, 149),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }
            okBtn.FlatAppearance.BorderSize = 0
            AddHandler okBtn.Click, Sub()
                                        dialog.DialogResult = DialogResult.OK
                                        dialog.Close()
                                    End Sub

            Dim cancelBtn As New Button With {
                .Text = "Cancel",
                .Dock = DockStyle.Right,
                .Width = 90,
                .BackColor = Color.White,
                .ForeColor = Color.FromArgb(76, 29, 149),
                .FlatStyle = FlatStyle.Flat
            }
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(196, 181, 253)
            AddHandler cancelBtn.Click, Sub()
                                            dialog.DialogResult = DialogResult.Cancel
                                            dialog.Close()
                                        End Sub

            buttonPanel.Controls.Add(okBtn)
            buttonPanel.Controls.Add(cancelBtn)

            dialog.AcceptButton = okBtn
            dialog.CancelButton = cancelBtn

            dialog.Controls.Add(choicePanel)
            dialog.Controls.Add(titleLbl)
            dialog.Controls.Add(buttonPanel)

            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return Nothing
            End If

            Select Case modeCombo.SelectedIndex
                Case 0 : Return PdfGroupingMode.AllHearings
                Case 1 : Return PdfGroupingMode.Day
                Case 2 : Return PdfGroupingMode.Week
                Case 3 : Return PdfGroupingMode.Month
                Case 4 : Return PdfGroupingMode.Year
                Case Else : Return PdfGroupingMode.Day
            End Select
        End Using
    End Function

    Private Function PromptPdfWeekRangeMode() As Nullable(Of PdfWeekRangeMode)
        Using dialog As New Form()
            dialog.Text = "Weekly Range"
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ClientSize = New Size(330, 190)
            dialog.BackColor = Color.White
            dialog.Font = New Font("Segoe UI", 9.5F)

            Dim titleLbl As New Label With {
                .Text = "Choose the weekly range style:",
                .Dock = DockStyle.Top,
                .Height = 32,
                .Padding = New Padding(14, 10, 14, 0),
                .ForeColor = Color.FromArgb(76, 29, 149),
                .Font = New Font("Segoe UI Semibold", 10.0F)
            }

            Dim choicePanel As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 92,
                .Padding = New Padding(14, 0, 14, 0)
            }

            Dim modeCombo As New ComboBox With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Font = New Font("Segoe UI", 9.5F)
            }
            modeCombo.Items.AddRange(New Object() {
                "Automatic",
                "Monday to Sunday"
            })
            modeCombo.SelectedIndex = 0

            Dim hintLbl As New Label With {
                .Text = "Automatic groups by 7-day blocks from the first hearing date.",
                .Dock = DockStyle.Top,
                .Height = 40,
                .Padding = New Padding(0, 10, 0, 0),
                .ForeColor = Color.FromArgb(107, 33, 168),
                .Font = New Font("Segoe UI", 8.5F)
            }

            choicePanel.Controls.Add(hintLbl)
            choicePanel.Controls.Add(modeCombo)

            Dim buttonPanel As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 58,
                .Padding = New Padding(14, 10, 14, 10),
                .BackColor = Color.FromArgb(245, 243, 255)
            }

            Dim okBtn As New Button With {
                .Text = "Continue",
                .Dock = DockStyle.Right,
                .Width = 100,
                .BackColor = Color.FromArgb(76, 29, 149),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }
            okBtn.FlatAppearance.BorderSize = 0
            AddHandler okBtn.Click, Sub()
                                        dialog.DialogResult = DialogResult.OK
                                        dialog.Close()
                                    End Sub

            Dim cancelBtn As New Button With {
                .Text = "Cancel",
                .Dock = DockStyle.Right,
                .Width = 90,
                .BackColor = Color.White,
                .ForeColor = Color.FromArgb(76, 29, 149),
                .FlatStyle = FlatStyle.Flat
            }
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(196, 181, 253)
            AddHandler cancelBtn.Click, Sub()
                                            dialog.DialogResult = DialogResult.Cancel
                                            dialog.Close()
                                        End Sub

            buttonPanel.Controls.Add(okBtn)
            buttonPanel.Controls.Add(cancelBtn)

            dialog.AcceptButton = okBtn
            dialog.CancelButton = cancelBtn
            dialog.Controls.Add(choicePanel)
            dialog.Controls.Add(titleLbl)
            dialog.Controls.Add(buttonPanel)

            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return Nothing
            End If

            If modeCombo.SelectedIndex = 1 Then
                Return PdfWeekRangeMode.MondayToSunday
            End If
            Return PdfWeekRangeMode.Automatic
        End Using
    End Function

    Private Function PromptPdfDateRange() As PdfDateRange
        Using dialog As New Form()
            dialog.Text = "PDF Date Range"
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ClientSize = New Size(360, 230)
            dialog.BackColor = Color.White
            dialog.Font = New Font("Segoe UI", 9.5F)

            Dim titleLbl As New Label With {
                .Text = "Choose the date range for the weekly PDF:",
                .Dock = DockStyle.Top,
                .Height = 34,
                .Padding = New Padding(14, 10, 14, 0),
                .ForeColor = Color.FromArgb(76, 29, 149),
                .Font = New Font("Segoe UI Semibold", 10.0F)
            }

            Dim bodyPanel As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 122,
                .Padding = New Padding(14, 0, 14, 0)
            }

            Dim fromLbl As New Label With {
                .Text = "From",
                .Dock = DockStyle.Top,
                .Height = 18,
                .ForeColor = Color.FromArgb(107, 33, 168),
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            }
            Dim fromPicker As New DateTimePicker With {
                .Dock = DockStyle.Top,
                .Format = DateTimePickerFormat.Custom,
                .CustomFormat = "MMMM d, yyyy",
                .Value = Date.Today
            }

            Dim gap1 As New Panel With {.Dock = DockStyle.Top, .Height = 10, .BackColor = Color.Transparent}

            Dim toLbl As New Label With {
                .Text = "To",
                .Dock = DockStyle.Top,
                .Height = 18,
                .ForeColor = Color.FromArgb(107, 33, 168),
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
            }
            Dim toPicker As New DateTimePicker With {
                .Dock = DockStyle.Top,
                .Format = DateTimePickerFormat.Custom,
                .CustomFormat = "MMMM d, yyyy",
                .Value = Date.Today.AddDays(6)
            }

            Dim hintLbl As New Label With {
                .Text = "Only hearings inside this range will be included in the weekly PDF.",
                .Dock = DockStyle.Top,
                .Height = 34,
                .Padding = New Padding(0, 10, 0, 0),
                .ForeColor = Color.FromArgb(107, 33, 168),
                .Font = New Font("Segoe UI", 8.5F)
            }

            bodyPanel.Controls.Add(hintLbl)
            bodyPanel.Controls.Add(toPicker)
            bodyPanel.Controls.Add(toLbl)
            bodyPanel.Controls.Add(gap1)
            bodyPanel.Controls.Add(fromPicker)
            bodyPanel.Controls.Add(fromLbl)

            Dim buttonPanel As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 58,
                .Padding = New Padding(14, 10, 14, 10),
                .BackColor = Color.FromArgb(245, 243, 255)
            }

            Dim okBtn As New Button With {
                .Text = "Continue",
                .Dock = DockStyle.Right,
                .Width = 100,
                .BackColor = Color.FromArgb(76, 29, 149),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }
            okBtn.FlatAppearance.BorderSize = 0
            AddHandler okBtn.Click, Sub()
                                        If toPicker.Value.Date < fromPicker.Value.Date Then
                                            MessageBox.Show(dialog, "The end date must be on or after the start date.", "Invalid Date Range", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Return
                                        End If
                                        dialog.DialogResult = DialogResult.OK
                                        dialog.Close()
                                    End Sub

            Dim cancelBtn As New Button With {
                .Text = "Cancel",
                .Dock = DockStyle.Right,
                .Width = 90,
                .BackColor = Color.White,
                .ForeColor = Color.FromArgb(76, 29, 149),
                .FlatStyle = FlatStyle.Flat
            }
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(196, 181, 253)
            AddHandler cancelBtn.Click, Sub()
                                            dialog.DialogResult = DialogResult.Cancel
                                            dialog.Close()
                                        End Sub

            buttonPanel.Controls.Add(okBtn)
            buttonPanel.Controls.Add(cancelBtn)

            dialog.AcceptButton = okBtn
            dialog.CancelButton = cancelBtn
            dialog.Controls.Add(bodyPanel)
            dialog.Controls.Add(titleLbl)
            dialog.Controls.Add(buttonPanel)

            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return Nothing
            End If

            Return New PdfDateRange With {
                .StartDate = fromPicker.Value.Date,
                .EndDate = toPicker.Value.Date
            }
        End Using
    End Function

    Private Function PromptPdfExportConfirm(mode As PdfGroupingMode, weekMode As Nullable(Of PdfWeekRangeMode), weekRange As PdfDateRange) As Boolean
        Using dialog As New Form()
            dialog.Text = "Confirm PDF Export"
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ClientSize = New Size(400, 220)
            dialog.BackColor = Color.White
            dialog.Font = New Font("Segoe UI", 9.5F)

            Dim summary As String = $"Mode: {GetPdfModeLabel(mode)}"
            If mode = PdfGroupingMode.Week Then
                Dim weekStyle = If(weekMode.HasValue AndAlso weekMode.Value = PdfWeekRangeMode.MondayToSunday, "Monday to Sunday", "Automatic")
                summary &= Environment.NewLine & $"Week style: {weekStyle}"
                If weekRange IsNot Nothing Then
                    summary &= Environment.NewLine & $"Date range: {weekRange.StartDate:MMMM d, yyyy} to {weekRange.EndDate:MMMM d, yyyy}"
                End If
            End If

            Dim titleLbl As New Label With {
                .Text = "Review the export settings before continuing:",
                .Dock = DockStyle.Top,
                .Height = 32,
                .Padding = New Padding(14, 10, 14, 0),
                .ForeColor = Color.FromArgb(76, 29, 149),
                .Font = New Font("Segoe UI Semibold", 10.0F)
            }

            Dim summaryBox As New TextBox With {
                .Dock = DockStyle.Top,
                .Height = 110,
                .Multiline = True,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = Color.White,
                .Font = New Font("Segoe UI", 9.5F),
                .Text = summary
            }

            Dim buttonPanel As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 58,
                .Padding = New Padding(14, 10, 14, 10),
                .BackColor = Color.FromArgb(245, 243, 255)
            }

            Dim okBtn As New Button With {
                .Text = "Continue",
                .Dock = DockStyle.Right,
                .Width = 100,
                .BackColor = Color.FromArgb(76, 29, 149),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }
            okBtn.FlatAppearance.BorderSize = 0
            AddHandler okBtn.Click, Sub()
                                        dialog.DialogResult = DialogResult.OK
                                        dialog.Close()
                                    End Sub

            Dim cancelBtn As New Button With {
                .Text = "Cancel",
                .Dock = DockStyle.Right,
                .Width = 90,
                .BackColor = Color.White,
                .ForeColor = Color.FromArgb(76, 29, 149),
                .FlatStyle = FlatStyle.Flat
            }
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(196, 181, 253)
            AddHandler cancelBtn.Click, Sub()
                                            dialog.DialogResult = DialogResult.Cancel
                                            dialog.Close()
                                        End Sub

            buttonPanel.Controls.Add(okBtn)
            buttonPanel.Controls.Add(cancelBtn)

            dialog.AcceptButton = okBtn
            dialog.CancelButton = cancelBtn
            dialog.Controls.Add(summaryBox)
            dialog.Controls.Add(titleLbl)
            dialog.Controls.Add(buttonPanel)

            Return dialog.ShowDialog(Me) = DialogResult.OK
        End Using
    End Function

    Private Shared Function BuildPdfGroups(records As List(Of HearingRecord), mode As PdfGroupingMode, weekMode As PdfWeekRangeMode, weekRange As PdfDateRange) As List(Of PdfGroup)
        Dim groups As New List(Of PdfGroup)()
        Dim scheduled = records.Where(Function(h) h.NextHearing <> Date.MinValue).ToList()
        If weekRange IsNot Nothing Then
            scheduled = scheduled.Where(Function(h) h.NextHearing.Date >= weekRange.StartDate AndAlso h.NextHearing.Date <= weekRange.EndDate).ToList()
        End If
        Dim unscheduled = If(weekRange IsNot Nothing, New List(Of HearingRecord)(), records.Where(Function(h) h.NextHearing = Date.MinValue).ToList())

        Select Case mode
            Case PdfGroupingMode.Day
                For Each group In scheduled.GroupBy(Function(h) h.NextHearing.Date).OrderBy(Function(g) g.Key)
                    groups.Add(New PdfGroup With {
                        .Title = group.Key.ToString("dddd, MMMM d, yyyy"),
                        .Records = group.OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                    })
                Next
            Case PdfGroupingMode.Week
                For Each group In scheduled.GroupBy(Function(h) h.NextHearing.Date).OrderBy(Function(g) g.Key)
                    groups.Add(New PdfGroup With {
                        .Title = group.Key.ToString("dddd, MMMM d, yyyy"),
                        .Records = group.OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                    })
                Next
            Case PdfGroupingMode.Month
                For Each group In scheduled.GroupBy(Function(h) New With {Key .Year = h.NextHearing.Year, Key .Month = h.NextHearing.Month}).OrderBy(Function(g) g.Key.Year).ThenBy(Function(g) g.Key.Month)
                    Dim monthDate = New DateTime(group.Key.Year, group.Key.Month, 1)
                    groups.Add(New PdfGroup With {
                        .Title = monthDate.ToString("MMMM yyyy"),
                        .Records = group.OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                    })
                Next
            Case PdfGroupingMode.Year
                For Each group In scheduled.GroupBy(Function(h) h.NextHearing.Year).OrderBy(Function(g) g.Key)
                    groups.Add(New PdfGroup With {
                        .Title = group.Key.ToString(),
                        .Records = group.OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                    })
                Next
        End Select

        If unscheduled.Count > 0 Then
            groups.Add(New PdfGroup With {
                .Title = "Pending / Unscheduled",
                .Records = unscheduled.OrderBy(Function(h) h.NameOfPdl).ThenBy(Function(h) h.No).ToList()
            })
        End If

        Return groups
    End Function

    Private Shared Function GetMondayWeekStart(value As Date) As Date
        Dim offset As Integer = CInt(value.DayOfWeek)
        If offset = 0 Then
            offset = 6
        Else
            offset -= 1
        End If
        Return value.Date.AddDays(-offset)
    End Function

    Private Shared Function GetAutomaticWeekStart(value As Date, baseDate As Date) As Date
        Dim daysFromBase = CInt((value.Date - baseDate.Date).TotalDays)
        Dim bucketStart = baseDate.Date.AddDays((daysFromBase \ 7) * 7)
        Return bucketStart
    End Function

    Private Shared Sub FormatGroupedHearingTable(table As Object, groups As List(Of PdfGroup), document As Object, mode As PdfGroupingMode)
        Dim yellow As Integer = ColorTranslator.ToOle(Color.FromArgb(245, 202, 66))
        Dim lightBlue As Integer = ColorTranslator.ToOle(Color.FromArgb(176, 224, 245))
        Dim white As Integer = ColorTranslator.ToOle(Color.White)
        Dim black As Integer = ColorTranslator.ToOle(Color.Black)
        Dim darkText As Integer = ColorTranslator.ToOle(Color.FromArgb(31, 41, 55))
        Dim mutedText As Integer = ColorTranslator.ToOle(Color.FromArgb(107, 114, 128))

        table.Style = "Table Grid"
        table.AllowAutoFit = False
        table.Range.Font.Name = "Arial"
        table.Range.Font.Size = 8.5F
        table.Range.Cells.VerticalAlignment = 1

        Dim usableWidth As Double = CDbl(document.Sections(1).PageSetup.PageWidth) - CDbl(document.Sections(1).PageSetup.LeftMargin) - CDbl(document.Sections(1).PageSetup.RightMargin)
        Dim ratios() As Double = {0.08, 0.34, 0.14, 0.15, 0.15, 0.14}

        For i As Integer = 1 To 6
            table.Columns(i).Width = usableWidth * ratios(i - 1)
        Next

        Dim row As Integer = 1
        If groups.Count = 0 Then
            table.Cell(row, 1).Merge(table.Cell(row, 6))
            table.Cell(row, 1).Range.Text = "No hearings scheduled for this export."
            table.Cell(row, 1).Range.Font.Italic = True
            table.Cell(row, 1).Range.Font.Color = mutedText
            table.Cell(row, 1).Range.ParagraphFormat.Alignment = 1
            table.Cell(row, 1).Shading.BackgroundPatternColor = white
            Return
        End If

        For Each group In groups
            table.Cell(row, 1).Merge(table.Cell(row, 6))
            table.Cell(row, 1).Range.Text = group.Title.ToUpperInvariant()
            table.Cell(row, 1).Range.Font.Bold = True
            table.Cell(row, 1).Range.Font.Size = 9.5F
            table.Cell(row, 1).Range.Font.Color = black
            table.Cell(row, 1).Shading.BackgroundPatternColor = yellow
            table.Cell(row, 1).Range.ParagraphFormat.Alignment = 1
            row += 1

            table.Cell(row, 1).Range.Text = "NO"
            table.Cell(row, 2).Range.Text = "NAME OF PDL"
            table.Cell(row, 3).Range.Text = "BR/COURT"
            table.Cell(row, 4).Range.Text = "HEARING STATUS"
            table.Cell(row, 5).Range.Text = "HEARING RESULT"
            table.Cell(row, 6).Range.Text = "NEXT HEARING"
            table.Rows(row).HeadingFormat = True
            table.Rows(row).Range.Font.Bold = True
            table.Rows(row).Range.Font.Color = black
            table.Rows(row).Shading.BackgroundPatternColor = lightBlue
            table.Rows(row).Range.ParagraphFormat.Alignment = 1
            row += 1

            If group.Records.Count = 0 Then
                table.Cell(row, 1).Merge(table.Cell(row, 6))
                table.Cell(row, 1).Range.Text = "No hearings in this group."
                table.Cell(row, 1).Range.Font.Italic = True
                table.Cell(row, 1).Range.Font.Color = mutedText
                table.Cell(row, 1).Range.ParagraphFormat.Alignment = 1
                table.Cell(row, 1).Shading.BackgroundPatternColor = white
                row += 1
            Else
                For recordIndex As Integer = 0 To group.Records.Count - 1
                    Dim hearing = group.Records(recordIndex)

                    table.Cell(row, 1).Range.Text = hearing.No
                    table.Cell(row, 2).Range.Text = hearing.NameOfPdl
                    table.Cell(row, 3).Range.Text = hearing.BrCourt
                    table.Cell(row, 4).Range.Text = hearing.Hearing1
                    table.Cell(row, 5).Range.Text = hearing.Hearing2
                    table.Cell(row, 6).Range.Text = If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("MMMM d, yyyy"))

                    table.Rows(row).Shading.BackgroundPatternColor = white
                    table.Rows(row).Range.Font.Color = darkText
                    table.Rows(row).Range.ParagraphFormat.Alignment = 0
                    table.Cell(row, 1).Range.ParagraphFormat.Alignment = 1
                    table.Cell(row, 3).Range.ParagraphFormat.Alignment = 1
                    table.Cell(row, 4).Range.ParagraphFormat.Alignment = 1
                    table.Cell(row, 5).Range.ParagraphFormat.Alignment = 1
                    table.Cell(row, 6).Range.ParagraphFormat.Alignment = 1
                    row += 1
                Next
            End If
        Next
    End Sub

    Private Shared Function GetPdfModeLabel(mode As PdfGroupingMode) As String
        Select Case mode
            Case PdfGroupingMode.AllHearings : Return "Master Report"
            Case PdfGroupingMode.Day : Return "Daily Report"
            Case PdfGroupingMode.Week : Return "Weekly Report"
            Case PdfGroupingMode.Month : Return "Monthly Report"
            Case PdfGroupingMode.Year : Return "Yearly Report"
            Case Else : Return "Report"
        End Select
    End Function

    Private Shared Function GetBundledTemplatePath() As String
        Return Path.Combine(Application.StartupPath, "Templates", "Court Calendar Template.docx")
    End Function

    Private Shared Function FindFirstBodyParagraph(document As Object) As Object
        For Each paragraph As Object In document.Paragraphs
            Dim text = NormalizeWordText(CStr(paragraph.Range.Text))
            If text <> "" Then Return paragraph
        Next

        Return Nothing
    End Function

    Private Shared Function FindParagraphContainingText(document As Object, needle As String) As Object
        For Each paragraph As Object In document.Paragraphs
            Dim text = NormalizeWordText(CStr(paragraph.Range.Text))
            If text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return paragraph
            End If
        Next

        Return Nothing
    End Function

    Private Shared Function FindParagraphIndexContainingText(document As Object, needle As String) As Integer
        For i As Integer = 1 To document.Paragraphs.Count
            Dim paragraph As Object = document.Paragraphs(i)
            Dim text = NormalizeWordText(CStr(paragraph.Range.Text))
            If text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return i
            End If
        Next

        Return 0
    End Function

    Private Shared Function NormalizeWordText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Return value.Replace(ChrW(13), " ").Replace(ChrW(7), " ").Trim()
    End Function

    Private Shared Sub FormatHearingTable(table As Object, records As List(Of HearingRecord), document As Object)
        Dim lightBlue As Integer = ColorTranslator.ToOle(Color.FromArgb(176, 224, 245))
        Dim white As Integer = ColorTranslator.ToOle(Color.White)
        Dim black As Integer = ColorTranslator.ToOle(Color.Black)
        Dim darkText As Integer = ColorTranslator.ToOle(Color.FromArgb(31, 41, 55))
        Dim mutedText As Integer = ColorTranslator.ToOle(Color.FromArgb(107, 114, 128))

        table.Style = "Table Grid"
        table.AllowAutoFit = False
        table.Range.Font.Name = "Arial"
        table.Range.Font.Size = 8.5F
        table.Range.Cells.VerticalAlignment = 1

        Dim usableWidth As Double = CDbl(document.Sections(1).PageSetup.PageWidth) - CDbl(document.Sections(1).PageSetup.LeftMargin) - CDbl(document.Sections(1).PageSetup.RightMargin)
        Dim ratios() As Double = {0.08, 0.34, 0.14, 0.15, 0.15, 0.14}

        For i As Integer = 1 To 6
            table.Columns(i).Width = usableWidth * ratios(i - 1)
        Next

        table.Cell(1, 1).Range.Text = "NO"
        table.Cell(1, 2).Range.Text = "NAME OF PDL"
        table.Cell(1, 3).Range.Text = "BR/COURT"
        table.Cell(1, 4).Range.Text = "HEARING STATUS"
        table.Cell(1, 5).Range.Text = "HEARING RESULT"
        table.Cell(1, 6).Range.Text = "NEXT HEARING"
        table.Rows(1).HeadingFormat = True
        table.Rows(1).Range.Font.Bold = True
        table.Rows(1).Range.Font.Color = black
        table.Rows(1).Shading.BackgroundPatternColor = lightBlue
        table.Rows(1).Range.ParagraphFormat.Alignment = 1

        If records.Count = 0 Then
            table.Cell(2, 1).Merge(table.Cell(2, 6))
            table.Cell(2, 1).Range.Text = "No hearings scheduled for this export."
            table.Cell(2, 1).Range.Font.Italic = True
            table.Cell(2, 1).Range.Font.Color = mutedText
            table.Cell(2, 1).Range.ParagraphFormat.Alignment = 1
            table.Cell(2, 1).Shading.BackgroundPatternColor = white
            Return
        End If

        For rowIndex As Integer = 0 To records.Count - 1
            Dim hearing = records(rowIndex)
            Dim wordRow As Integer = rowIndex + 2

            table.Cell(wordRow, 1).Range.Text = hearing.No
            table.Cell(wordRow, 2).Range.Text = hearing.NameOfPdl
            table.Cell(wordRow, 3).Range.Text = hearing.BrCourt
            table.Cell(wordRow, 4).Range.Text = hearing.Hearing1
            table.Cell(wordRow, 5).Range.Text = hearing.Hearing2
            table.Cell(wordRow, 6).Range.Text = If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("MMMM d, yyyy"))

            table.Rows(wordRow).Shading.BackgroundPatternColor = white
            table.Rows(wordRow).Range.Font.Color = darkText
            table.Rows(wordRow).Range.ParagraphFormat.Alignment = 0
            table.Cell(wordRow, 1).Range.ParagraphFormat.Alignment = 1
            table.Cell(wordRow, 3).Range.ParagraphFormat.Alignment = 1
            table.Cell(wordRow, 4).Range.ParagraphFormat.Alignment = 1
            table.Cell(wordRow, 5).Range.ParagraphFormat.Alignment = 1
            table.Cell(wordRow, 6).Range.ParagraphFormat.Alignment = 1
        Next
    End Sub

    Private Shared Sub ReleaseComObjectSafe(ByRef comObject As Object)
        If comObject Is Nothing Then Return

        Try
            If Marshal.IsComObject(comObject) Then
                Marshal.FinalReleaseComObject(comObject)
            End If
        Catch
        Finally
            comObject = Nothing
        End Try
    End Sub

    Private Sub ShowSettingsPopup()
        Using dlg As New Form()
            dlg.Text = "Settings & Clear Data"
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MinimizeBox = False
            dlg.MaximizeBox = False
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 10.0F)
            dlg.ClientSize = New Size(400, 380)

            Dim titlePanel As New Panel With {.Height = 56, .BackColor = Color.FromArgb(76, 29, 149), .Dock = DockStyle.Top}
            Dim titleLbl As New Label With {.Text = "SETTINGS & CLEAR DATA", .Dock = DockStyle.Fill, .ForeColor = Color.White, .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleCenter}
            titlePanel.Controls.Add(titleLbl)

            Dim container As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(20), .FlowDirection = FlowDirection.TopDown, .WrapContents = False}

            Dim descLbl As New Label With {.Text = "Select an option below to clear/clean hearing data. A backup file is automatically generated before clearing.", .Width = 360, .Height = 45, .ForeColor = Color.FromArgb(139, 92, 246), .Font = New Font("Segoe UI", 9.0F)}
            container.Controls.Add(descLbl)

            Dim btnWidth = 350
            Dim btnHeight = 36

            Dim btnClearAll = MakeSideButton("Clear All Hearings", Color.FromArgb(198, 40, 40), Color.White)
            btnClearAll.Width = btnWidth
            btnClearAll.Height = btnHeight
            AddHandler btnClearAll.Click, Sub()
                                              Dim confirm = MessageBox.Show(dlg, "Are you sure you want to delete ALL hearings from the database? This cannot be undone.", "Confirm Clear All", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                              If confirm = DialogResult.Yes Then
                                                  Try
                                                      ShowProgress("Backing up and clearing all hearings...")
                                                      repository.BackupCurrentData()
                                                      repository.ClearHearings(Function(h) True)
                                                      ReloadCalendar()
                                                      HideProgress()
                                                      MessageBox.Show(dlg, "All hearings have been deleted. Backup created in Backups folder.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                  Catch ex As Exception
                                                      HideProgress()
                                                      ShowError("Error", "Failed to clear hearings.", ex)
                                                  End Try
                                              End If
                                          End Sub
            container.Controls.Add(btnClearAll)
            container.Controls.Add(New Panel With {.Height = 10})

            Dim btnClearMonth = MakeSideButton("Clear Hearings by Month", Color.FromArgb(109, 40, 217), Color.White)
            btnClearMonth.Width = btnWidth
            btnClearMonth.Height = btnHeight
            AddHandler btnClearMonth.Click, Sub()
                                                Dim months = hearings.Where(Function(h) h.NextHearing <> Date.MinValue).
                                                    Select(Function(h) New DateTime(h.NextHearing.Year, h.NextHearing.Month, 1)).
                                                    Distinct().OrderBy(Function(d) d).ToList()
                                                If months.Count = 0 Then
                                                    MessageBox.Show(dlg, "No scheduled hearings found to clear.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                    Return
                                                End If
                                                Using monthDlg As New Form()
                                                    monthDlg.Text = "Select Month to Clear"
                                                    monthDlg.ClientSize = New Size(300, 150)
                                                    monthDlg.StartPosition = FormStartPosition.CenterParent
                                                    monthDlg.FormBorderStyle = FormBorderStyle.FixedDialog
                                                    monthDlg.MinimizeBox = False
                                                    monthDlg.MaximizeBox = False
                                                    monthDlg.BackColor = Color.White
                                                    Dim lbl As New Label With {.Text = "Select Month:", .Location = New Point(20, 20), .AutoSize = True}
                                                    Dim cb As New ComboBox With {.Location = New Point(20, 45), .Width = 260, .DropDownStyle = ComboBoxStyle.DropDownList}
                                                    For Each m In months
                                                        cb.Items.Add(m.ToString("MMMM yyyy", CultureInfo.InvariantCulture))
                                                    Next
                                                    cb.SelectedIndex = 0
                                                    Dim okBtn As New Button With {.Text = "Clear Month", .Location = New Point(60, 90), .Width = 100, .DialogResult = DialogResult.OK}
                                                    Dim cancelBtn As New Button With {.Text = "Cancel", .Location = New Point(170, 90), .Width = 80, .DialogResult = DialogResult.Cancel}
                                                    monthDlg.Controls.Add(lbl)
                                                    monthDlg.Controls.Add(cb)
                                                    monthDlg.Controls.Add(okBtn)
                                                    monthDlg.Controls.Add(cancelBtn)
                                                    If monthDlg.ShowDialog(dlg) = DialogResult.OK Then
                                                        Dim selectedMonthText = cb.SelectedItem.ToString()
                                                        Dim confirm = MessageBox.Show(dlg, $"Are you sure you want to clear all hearings for {selectedMonthText}?", "Confirm Clear Month", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                                        If confirm = DialogResult.Yes Then
                                                            Try
                                                                ShowProgress($"Clearing hearings for {selectedMonthText}...")
                                                                repository.BackupCurrentData()
                                                                repository.ClearHearings(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.ToString("MMMM yyyy", CultureInfo.InvariantCulture) = selectedMonthText)
                                                                ReloadCalendar()
                                                                HideProgress()
                                                                MessageBox.Show(dlg, $"Cleared hearings for {selectedMonthText}. Backup created.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                            Catch ex As Exception
                                                                HideProgress()
                                                                ShowError("Error", $"Failed to clear hearings for {selectedMonthText}.", ex)
                                                            End Try
                                                        End If
                                                    End If
                                                End Using
                                            End Sub
            container.Controls.Add(btnClearMonth)
            container.Controls.Add(New Panel With {.Height = 10})

            Dim btnClearRange = MakeSideButton("Clear Hearings by Date Range", Color.FromArgb(109, 40, 217), Color.White)
            btnClearRange.Width = btnWidth
            btnClearRange.Height = btnHeight
            AddHandler btnClearRange.Click, Sub()
                                                Using rangeDlg As New Form()
                                                    rangeDlg.Text = "Select Date Range"
                                                    rangeDlg.ClientSize = New Size(340, 180)
                                                    rangeDlg.StartPosition = FormStartPosition.CenterParent
                                                    rangeDlg.FormBorderStyle = FormBorderStyle.FixedDialog
                                                    rangeDlg.MinimizeBox = False
                                                    rangeDlg.MaximizeBox = False
                                                    rangeDlg.BackColor = Color.White
                                                    Dim lblStart As New Label With {.Text = "Start Date:", .Location = New Point(20, 20), .AutoSize = True}
                                                    Dim dpStart As New DateTimePicker With {.Location = New Point(20, 45), .Width = 130, .Format = DateTimePickerFormat.Short}
                                                    Dim lblEnd As New Label With {.Text = "End Date:", .Location = New Point(170, 20), .AutoSize = True}
                                                    Dim dpEnd As New DateTimePicker With {.Location = New Point(170, 45), .Width = 130, .Format = DateTimePickerFormat.Short}
                                                    Dim okBtn As New Button With {.Text = "Clear Range", .Location = New Point(80, 110), .Width = 100, .DialogResult = DialogResult.OK}
                                                    Dim cancelBtn As New Button With {.Text = "Cancel", .Location = New Point(190, 110), .Width = 80, .DialogResult = DialogResult.Cancel}
                                                    rangeDlg.Controls.Add(lblStart)
                                                    rangeDlg.Controls.Add(dpStart)
                                                    rangeDlg.Controls.Add(lblEnd)
                                                    rangeDlg.Controls.Add(dpEnd)
                                                    rangeDlg.Controls.Add(okBtn)
                                                    rangeDlg.Controls.Add(cancelBtn)
                                                    If rangeDlg.ShowDialog(dlg) = DialogResult.OK Then
                                                        Dim startD = dpStart.Value.Date
                                                        Dim endD = dpEnd.Value.Date
                                                        If startD > endD Then
                                                            MessageBox.Show(dlg, "Start Date cannot be after End Date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                            Return
                                                        End If
                                                        Dim confirm = MessageBox.Show(dlg, $"Are you sure you want to clear hearings from {startD:yyyy-MM-dd} to {endD:yyyy-MM-dd}?", "Confirm Clear Range", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                                        If confirm = DialogResult.Yes Then
                                                            Try
                                                                ShowProgress($"Clearing hearings {startD:MMM d} - {endD:MMM d, yyyy}...")
                                                                repository.BackupCurrentData()
                                                                repository.ClearHearings(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date >= startD AndAlso h.NextHearing.Date <= endD)
                                                                ReloadCalendar()
                                                                HideProgress()
                                                                MessageBox.Show(dlg, $"Cleared hearings between {startD:yyyy-MM-dd} and {endD:yyyy-MM-dd}. Backup created.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                            Catch ex As Exception
                                                                HideProgress()
                                                                ShowError("Error", $"Failed to clear hearings from {startD:yyyy-MM-dd} to {endD:yyyy-MM-dd}.", ex)
                                                            End Try
                                                        End If
                                                    End If
                                                End Using
                                            End Sub
            container.Controls.Add(btnClearRange)
            container.Controls.Add(New Panel With {.Height = 10})

            Dim btnClearLast = MakeSideButton("Clear Last Exported Data", Color.FromArgb(109, 40, 217), Color.White)
            btnClearLast.Width = btnWidth
            btnClearLast.Height = btnHeight
            AddHandler btnClearLast.Click, Sub()
                                               If _lastExportedIds.Count = 0 Then
                                                   MessageBox.Show(dlg, "No records were exported in the current session yet.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                   Return
                                               End If
                                               Dim confirm = MessageBox.Show(dlg, $"Are you sure you want to clear the {_lastExportedIds.Count} hearings from the last export?", "Confirm Clear Last Export", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                               If confirm = DialogResult.Yes Then
                                                   Try
                                                       ShowProgress("Clearing last exported hearings...")
                                                       repository.BackupCurrentData()
                                                       repository.ClearHearings(Function(h) _lastExportedIds.Contains(h.Id))
                                                       _lastExportedIds.Clear()
                                                       ReloadCalendar()
                                                       HideProgress()
                                                       MessageBox.Show(dlg, "Last exported hearings have been cleared. Backup created.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                   Catch ex As Exception
                                                       HideProgress()
                                                       ShowError("Error", "Failed to clear the last exported hearings.", ex)
                                                   End Try
                                               End If
                                           End Sub
            container.Controls.Add(btnClearLast)

            dlg.Controls.Add(container)
            dlg.Controls.Add(titlePanel)
            dlg.ShowDialog(Me)
        End Using
    End Sub

    Private Sub ShowProgress(message As String)
        If InvokeRequired Then
            Invoke(Sub() ShowProgress(message))
            Return
        End If
        _statusLabel.Text = "Status: " & message
        If _progressPanel IsNot Nothing Then
            _progressPanel.Visible = True
            _progressPanel.Refresh()
        End If
        _statusLabel.Refresh()
    End Sub

    Private Sub HideProgress()
        If InvokeRequired Then
            Invoke(Sub() HideProgress())
            Return
        End If
        If _progressPanel IsNot Nothing Then _progressPanel.Visible = False
    End Sub

End Class
