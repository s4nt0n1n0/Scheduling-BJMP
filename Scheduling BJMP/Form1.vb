Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text.Encodings.Web
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Globalization
Imports System.Windows.Forms

Public Class Form1
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
        Await webView.EnsureCoreWebView2Async()
        AddHandler webView.CoreWebView2.WebMessageReceived, AddressOf CalendarMessageReceived

        Dim calendarPath = Path.Combine(AppContext.BaseDirectory, "Calendar", "calendar.html")
        webView.CoreWebView2.Navigate(New Uri(calendarPath).AbsoluteUri)
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

        ' ── Sidebar ──────────────────────────────────────────
        detailsPanel.Dock = DockStyle.Fill
        detailsPanel.BackColor = Color.FromArgb(250, 249, 255)
        detailsPanel.Padding = New Padding(0)

        ' Header strip
        Dim headerStrip As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .BackColor = Color.FromArgb(245, 158, 11)
        }
        Dim headerLbl As New Label With {
            .Text = "BJMP  HEARING PANEL",
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(30, 27, 75),
            .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Padding = New Padding(0, 0, 0, 2)
        }
        Dim goldBar As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 4,
            .BackColor = Color.FromArgb(245, 158, 11)
        }
        headerStrip.Controls.Add(headerLbl)
        headerStrip.Controls.Add(goldBar)

        ' Inner scroll panel
        Dim innerPanel As New Panel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .Padding = New Padding(16, 12, 16, 8)
        }

        ' ── Search section ─────────────────────
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
        Dim div1 As New Panel With { .Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(221, 214, 254) }

        ' ── Hearings list section ───────────────
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
        Dim div2 As New Panel With { .Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(221, 214, 254) }

        ' ── Details card section ─────────────────
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
        detailsTitleLabel.Text = "—  No hearing selected"
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

        ' ── Other Hearings section ───────────────
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
                ev.Value = $"📅 {dtStr} - {h.BrCourt} ({h.Hearing1})"
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

        ' ── Button strip at bottom ────────────────
        Dim btnStrip As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 212,
            .BackColor = Color.FromArgb(245, 243, 255),
            .Padding = New Padding(12, 8, 12, 10)
        }

        Dim addButton = MakeSideButton("＋  Add New Hearing", Color.FromArgb(76, 29, 149), Color.White)
        addButton.Dock = DockStyle.Top
        addButton.Height = 38
        AddHandler addButton.Click, Sub() ShowAddDialog(selectedDate)

        Dim spacer1 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim importButton = MakeSideButton("⬆  Import Excel / XML", Color.FromArgb(109, 40, 217), Color.White)
        importButton.Dock = DockStyle.Top
        importButton.Height = 36
        AddHandler importButton.Click, Sub() ImportDataFile()

        Dim spacer2 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim exportButton = MakeSideButton("⬇  Export Hearings", Color.FromArgb(245, 158, 11), Color.FromArgb(30, 27, 75))
        exportButton.Dock = DockStyle.Top
        exportButton.Height = 36
        AddHandler exportButton.Click, Async Sub()
            Try
                Await ExportAllHearings()
            Catch ex As Exception
                MessageBox.Show(Me, "Failed to export: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Dim spacer3 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim settingsButton = MakeSideButton("⚙  Settings / Clear Data", Color.FromArgb(109, 40, 217), Color.White)
        settingsButton.Dock = DockStyle.Top
        settingsButton.Height = 34
        AddHandler settingsButton.Click, Sub() ShowSettingsPopup()

        Dim spacer4 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim refreshButton = MakeSideButton("↺  Refresh Calendar", Color.FromArgb(221, 214, 254), Color.FromArgb(30, 27, 75))
        refreshButton.Dock = DockStyle.Top
        refreshButton.Height = 34
        AddHandler refreshButton.Click, Sub() ReloadCalendar()

        btnStrip.Controls.Add(refreshButton)
        btnStrip.Controls.Add(spacer4)
        btnStrip.Controls.Add(settingsButton)
        btnStrip.Controls.Add(spacer3)
        btnStrip.Controls.Add(exportButton)
        btnStrip.Controls.Add(spacer2)
        btnStrip.Controls.Add(importButton)
        btnStrip.Controls.Add(spacer1)
        btnStrip.Controls.Add(addButton)

        ' ── Progress / status bar ─────────────────────────────────
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
        valueLabel.Text = "—"
        valueLabel.AutoSize = True
        valueLabel.Dock = DockStyle.Fill
        valueLabel.ForeColor = Color.FromArgb(30, 27, 75)
        valueLabel.Font = New Font("Segoe UI", 9.5F)
        valueLabel.TextAlign = ContentAlignment.MiddleLeft
        valueLabel.Padding = New Padding(0, 6, 0, 6)
        table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        table.Controls.Add(captionLbl)
        table.Controls.Add(valueLabel)
    End Sub

    Private Async Sub CalendarMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Dim messageJson As String = ""
        Try
            messageJson = e.WebMessageAsJson
        Catch ex As Exception
            Return
        End Try

        ' Defer processing using Await Task.Delay to completely unwind and unblock 
        ' the WebView2 event dispatch stack before opening modal dialogs or invoking scripts.
        Await Task.Delay(20)

        Try
            Using document = JsonDocument.Parse(messageJson)
                Dim root = document.RootElement
                Dim action = root.GetProperty("action").GetString()

                Select Case action
                    Case "ready"
                        ReloadCalendar()
                    Case "dateview"
                        Dim clickedDate = Date.Parse(root.GetProperty("date").GetString())
                        SelectDate(clickedDate)
                        ShowDatePopup(clickedDate)
                    Case "select"
                        Dim idText = root.GetProperty("id").GetString()
                        SelectHearing(idText)
                    Case "move"
                        Dim oldDateStr = ""
                        Dim oldDateElement As JsonElement
                        If root.TryGetProperty("oldDate", oldDateElement) Then
                            oldDateStr = oldDateElement.GetString()
                        End If
                        Dim targetId = root.GetProperty("id").GetString()
                        Dim targetDate = Date.Parse(root.GetProperty("date").GetString())
                        MoveHearing(targetId, targetDate, oldDateStr)
                End Select
            End Using
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Calendar Update", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ReloadCalendar()
        End Try
    End Sub

    Private Async Sub ReloadCalendar()
        Try
            hearings = repository.LoadHearings()
            ' Restore in-memory history entries that survived the reload
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
            MessageBox.Show(Me, "Failed to reload calendar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            MessageBox.Show(Me, $"Failed to add hearing:{Environment.NewLine}{ex.Message}", "Add Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ShowDatePopup(clickedDate As Date)
        Dim shouldShow As Boolean = True
        While shouldShow
            shouldShow = False
            Dim dateHearings = hearings.
                Where(Function(h) h.NextHearing.Date = clickedDate.Date).
                OrderBy(Function(h) h.No).
                ToList()

            Using popup As New Form()
                popup.Text = $"Hearings — {clickedDate:MMMM d, yyyy}"
                popup.StartPosition = FormStartPosition.CenterParent
                popup.FormBorderStyle = FormBorderStyle.FixedDialog
                popup.MinimizeBox = False
                popup.MaximizeBox = False
                popup.BackColor = Color.White
                popup.Font = New Font("Segoe UI", 10.0F)
                popup.ClientSize = New Size(520, Math.Min(80 + dateHearings.Count * 44 + 60, 520))

                ' Title bar
                Dim titlePanel As New Panel With {
                    .Dock = DockStyle.Top,
                    .Height = 60,
                    .BackColor = Color.FromArgb(245, 158, 11),
                    .Padding = New Padding(18, 0, 18, 0)
                }
                Dim titleLbl As New Label With {
                    .Text = $"📅  {clickedDate:dddd, MMMM d, yyyy}",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .Font = New Font("Segoe UI Semibold", 13.0F),
                    .TextAlign = ContentAlignment.MiddleLeft
                }
                Dim countLbl As New Label With {
                    .Text = $"{dateHearings.Count} hearing(s)",
                    .Dock = DockStyle.Right,
                    .Width = 130,
                    .ForeColor = Color.FromArgb(245, 158, 11),
                    .Font = New Font("Segoe UI", 9.5F),
                    .TextAlign = ContentAlignment.MiddleRight
                }
                titlePanel.Controls.Add(titleLbl)
                titlePanel.Controls.Add(countLbl)

                ' Scroll panel for hearings
                Dim scroll As New Panel With {
                    .Dock = DockStyle.Fill,
                    .AutoScroll = True,
                    .Padding = New Padding(14, 8, 14, 8)
                }

                If dateHearings.Count = 0 Then
                    Dim emptyLbl As New Label With {
                        .Text = "No hearings scheduled for this date.",
                        .Dock = DockStyle.Top,
                        .Height = 50,
                        .ForeColor = Color.FromArgb(139, 92, 246),
                        .TextAlign = ContentAlignment.MiddleCenter,
                        .Font = New Font("Segoe UI", 10.0F, FontStyle.Italic)
                    }
                    scroll.Controls.Add(emptyLbl)
                Else
                    Dim yPos As Integer = 8
                    For Each h In dateHearings
                        Dim currentH = h
                        Dim row As New Panel With {
                            .Location = New Point(0, yPos),
                            .Size = New Size(470, 38),
                            .BackColor = Color.FromArgb(245, 248, 252),
                            .Cursor = Cursors.Hand
                        }
                        Dim statusDot As New Label With {
                            .Location = New Point(10, 13),
                            .Size = New Size(12, 12),
                            .BackColor = Color.FromArgb(245, 158, 11)
                        }
                        ' Make dot round-ish with a border
                        Dim noLbl As New Label With {
                            .Text = h.No,
                            .Location = New Point(32, 0),
                            .Size = New Size(38, 38),
                            .TextAlign = ContentAlignment.MiddleCenter,
                            .ForeColor = Color.FromArgb(107, 33, 168),
                            .Font = New Font("Segoe UI", 9.0F)
                        }
                        Dim nameLbl As New Label With {
                            .Text = h.NameOfPdl,
                            .Location = New Point(72, 0),
                            .Size = New Size(230, 38),
                            .TextAlign = ContentAlignment.MiddleLeft,
                            .ForeColor = Color.FromArgb(30, 27, 75),
                            .Font = New Font("Segoe UI Semibold", 10.0F)
                        }
                        Dim courtLbl As New Label With {
                            .Text = h.BrCourt,
                            .Location = New Point(304, 0),
                            .Size = New Size(100, 38),
                            .TextAlign = ContentAlignment.MiddleLeft,
                            .ForeColor = Color.FromArgb(91, 33, 182),
                            .Font = New Font("Segoe UI", 9.5F)
                        }
                        Dim viewBtn As New Button With {
                            .Text = "View",
                            .Location = New Point(410, 6),
                            .Size = New Size(55, 26),
                            .BackColor = Color.FromArgb(245, 158, 11),
                            .ForeColor = Color.FromArgb(30, 27, 75),
                            .FlatStyle = FlatStyle.Flat,
                            .Font = New Font("Segoe UI", 8.5F)
                        }
                        viewBtn.FlatAppearance.BorderSize = 0
                        AddHandler viewBtn.Click, Sub()
                            ' Close the popup cleanly by setting DialogResult, then show detail
                            popup.Tag = currentH  ' store the hearing to open after closing
                            popup.DialogResult = DialogResult.OK
                            popup.Close()
                        End Sub
                        row.Controls.AddRange(New Control() {statusDot, noLbl, nameLbl, courtLbl, viewBtn})
                        scroll.Controls.Add(row)
                        yPos += 44
                    Next
                End If

                ' Bottom buttons
                Dim btnPanel As New Panel With {
                    .Dock = DockStyle.Bottom,
                    .Height = 52,
                    .BackColor = Color.FromArgb(242, 246, 250),
                    .Padding = New Padding(14, 8, 14, 8)
                }
                Dim addBtn As New Button With {
                    .Text = "＋ Add Hearing",
                    .Dock = DockStyle.Right,
                    .Width = 140,
                    .BackColor = Color.FromArgb(245, 158, 11),
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI Semibold", 9.5F)
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
                    .Width = 80,
                    .BackColor = Color.FromArgb(237, 233, 254),
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.5F),
                    .Margin = New Padding(0, 0, 8, 0)
                }
                closeBtn.FlatAppearance.BorderSize = 0
                AddHandler closeBtn.Click, Sub() popup.Close()
                btnPanel.Controls.Add(addBtn)
                btnPanel.Controls.Add(closeBtn)

                popup.Controls.Add(scroll)
                popup.Controls.Add(btnPanel)
                popup.Controls.Add(titlePanel)

                Dim result = popup.ShowDialog(Me)

                ' After popup closes: if a hearing was tagged for viewing, show its detail popup
                If result = DialogResult.OK Then
                    If TypeOf popup.Tag Is HearingRecord Then
                        Dim selectedHearing = DirectCast(popup.Tag, HearingRecord)
                        ShowHearingDetailPopup(selectedHearing)
                        ' Re-open the date popup so the user can continue browsing
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
            selectedDate = hearing.NextHearing.Date
            DisplayHearing(hearing)
            RefreshSideList()
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

            Dim logEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} — Duplicated to {nextHearing.Date:MMMM d, yyyy} from original date {oldDate:MMMM d, yyyy}"
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
            MessageBox.Show(Me, $"Failed to move hearing:{Environment.NewLine}{ex.Message}", "Move Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            HideProgress()
        End Try
    End Sub

    Private Sub DisplayHearing(hearing As HearingRecord)
        If _isUpdatingDisplay Then Return
        _isUpdatingDisplay = True
        Try
            detailsTitleLabel.Text = hearing.NameOfPdl
            noValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.No), "—", hearing.No)
            nameValueLabel.Text = hearing.NameOfPdl
            courtValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.BrCourt), "—", hearing.BrCourt)
            hearing1ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing1), "—", hearing.Hearing1)
            hearing2ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing2), "—", hearing.Hearing2)
            dateValueLabel.Text = If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("MMMM d, yyyy"))

            ' Update other hearings list for the same person
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

            ' Highlight these dates on the calendar
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
            currentHearing = Nothing ' reset to exit loop if no other is double-clicked

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
                .BorderStyle = BorderStyle.None,
                .BackColor = Color.FromArgb(250, 249, 255),
                .Font = New Font("Segoe UI", 8.5F),
                .ForeColor = Color.FromArgb(91, 33, 182),
                .ItemHeight = 20,
                .FormattingEnabled = True
            }

            Using dlg As New Form()
                dlg.Text = hearing.NameOfPdl
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog
                dlg.MinimizeBox = False
                dlg.MaximizeBox = False
                dlg.BackColor = Color.White
                dlg.Font = New Font("Segoe UI", 10.0F)
                dlg.ClientSize = New Size(480, 680)

                ' Header bar
                Dim hdr As New Panel With {
                    .Height = 64,
                    .BackColor = Color.FromArgb(245, 158, 11),
                    .Margin = New Padding(0),
                    .Dock = DockStyle.Fill
                }
                Dim hdrName As New Label With {
                    .Text = hearing.NameOfPdl,
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .Font = New Font("Segoe UI Semibold", 13.0F),
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Padding = New Padding(0, 0, 0, 4)
                }
                Dim hdrGold As New Panel With {
                    .Dock = DockStyle.Bottom,
                    .Height = 4,
                    .BackColor = Color.FromArgb(245, 158, 11)
                }
                hdr.Controls.Add(hdrName)
                hdr.Controls.Add(hdrGold)

                ' Body container TableLayoutPanel (bulletproof layout!)
                Dim bodyPanel As New TableLayoutPanel With {
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 8,
                    .Padding = New Padding(24, 16, 24, 12),
                    .Margin = New Padding(0),
                    .AutoScroll = True,
                    .BackColor = Color.White
                }
                bodyPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Row 0: grid
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Row 1: warningLabel
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 25)) ' Row 2: divLine
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 36)) ' Row 3: btnPanel
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 15)) ' Row 4: other spacer
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F)) ' Row 5: other hearings list
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 15)) ' Row 6: history spacer
                bodyPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F)) ' Row 7: history log

                ' Info grid
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

                ' Row 0: Case No.
                Dim cap0 As New Label With {
                    .Text = "Case No.",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(109, 40, 217),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                Dim val0 As New Label With {
                    .Text = If(String.IsNullOrWhiteSpace(hearing.No), "—", hearing.No),
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .Font = New Font("Segoe UI Semibold", 9.5F),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap0, 0, 0)
                grid.Controls.Add(val0, 1, 0)

                ' Row 1: BR / Court
                Dim cap1 As New Label With {
                    .Text = "BR / Court",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(109, 40, 217),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                Dim val1 As New Label With {
                    .Text = If(String.IsNullOrWhiteSpace(hearing.BrCourt), "—", hearing.BrCourt),
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .Font = New Font("Segoe UI Semibold", 9.5F),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap1, 0, 1)
                grid.Controls.Add(val1, 1, 1)

                ' Row 2: Hearing 1 (text)
                Dim cap2 As New Label With {
                    .Text = "Hearing",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(109, 40, 217),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                Dim hearing1Text As New TextBox With {
                    .Text = hearing.Hearing1,
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.5F),
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Margin = New Padding(0, 4, 0, 4)
                }
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap2, 0, 2)
                grid.Controls.Add(hearing1Text, 1, 2)

                ' Row 3: Hearing 2 (text)
                Dim cap3 As New Label With {
                    .Text = "Hearing",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(109, 40, 217),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                Dim hearing2Text As New TextBox With {
                    .Text = hearing.Hearing2,
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.5F),
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Margin = New Padding(0, 4, 0, 4)
                }
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap3, 0, 3)
                grid.Controls.Add(hearing2Text, 1, 3)

                ' Row 4: Next Hearing date
                Dim cap4 As New Label With {
                    .Text = "Next Hearing",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.FromArgb(109, 40, 217),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleLeft,
                    .Padding = New Padding(0, 6, 0, 6),
                    .AutoSize = True
                }
                Dim datePicker As New DateTimePicker With {
                    .Dock = DockStyle.Fill,
                    .Font = New Font("Segoe UI", 9.5F),
                    .Format = DateTimePickerFormat.Custom,
                    .CustomFormat = "MMMM d, yyyy",
                    .Value = If(hearing.NextHearing = Date.MinValue, Date.Today, hearing.NextHearing.Date),
                    .Margin = New Padding(0, 4, 0, 4)
                }
                grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                grid.Controls.Add(cap4, 0, 4)
                grid.Controls.Add(datePicker, 1, 4)

                ' Warning Label for same-day duplicates in Detail View
                Dim detailWarningLabel As New Label With {
                    .ForeColor = Color.FromArgb(198, 40, 40),
                    .Font = New Font("Segoe UI Semibold", 8.5F),
                    .Text = "",
                    .Visible = False,
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(0, 4, 0, 4),
                    .AutoSize = True
                }

                Dim checkDetailDuplicate = Sub()
                    Dim targetDate = datePicker.Value.Date
                    If hearings Is Nothing OrElse hearings.Count = 0 Then
                        detailWarningLabel.Visible = False
                        Return
                    End If

                    Dim hasDup = hearings.Any(Function(h)
                        Return h.Id <> hearing.Id AndAlso
                               String.Equals(h.NameOfPdl.Trim(), hearing.NameOfPdl.Trim(), StringComparison.OrdinalIgnoreCase) AndAlso
                               h.NextHearing.Date = targetDate
                    End Function)

                    If hasDup Then
                        detailWarningLabel.Text = "⚠️ Note: This person already has a hearing scheduled on this date!"
                        detailWarningLabel.Visible = True
                    Else
                        detailWarningLabel.Visible = False
                    End If
                End Sub

                AddHandler datePicker.ValueChanged, Sub() checkDetailDuplicate()
                checkDetailDuplicate()

                ' Divider
                Dim divLine As New Panel With {
                    .Dock = DockStyle.Top,
                    .Height = 1,
                    .BackColor = Color.FromArgb(221, 214, 254),
                    .Margin = New Padding(0, 12, 0, 12)
                }

                ' Action buttons panel
                Dim btnPanel As New Panel With {
                    .Dock = DockStyle.Top,
                    .Height = 36,
                    .Margin = New Padding(0)
                }
                
                Dim deleteBtn As New Button With {
                    .Text = "Delete Hearing",
                    .Location = New Point(0, 0),
                    .Size = New Size(115, 34),
                    .BackColor = Color.FromArgb(198, 40, 40),
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI Semibold", 9.0F)
                }
                deleteBtn.FlatAppearance.BorderSize = 0
                AddHandler deleteBtn.Click, Sub()
                    Dim confirm = MessageBox.Show(dlg, $"Are you sure you want to delete the hearing schedule for {hearing.NameOfPdl}?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If confirm = DialogResult.Yes Then
                        Try
                            repository.DeleteHearing(hearing.Id)
                            ReloadCalendar()
                            
                            ' Clear the details sidebar
                            detailsTitleLabel.Text = "—  No hearing selected"
                            noValueLabel.Text = "—"
                            nameValueLabel.Text = "—"
                            courtValueLabel.Text = "—"
                            hearing1ValueLabel.Text = "—"
                            hearing2ValueLabel.Text = "—"
                            dateValueLabel.Text = "—"
                            
                            MessageBox.Show(dlg, "Hearing deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            dlg.Close()
                        Catch ex As Exception
                            MessageBox.Show(dlg, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End If
                End Sub

                Dim saveBtn As New Button With {
                    .Text = "Save Changes",
                    .Location = New Point(200, 0),
                    .Size = New Size(130, 34),
                    .BackColor = Color.FromArgb(245, 158, 11),
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI Semibold", 9.0F)
                }
                saveBtn.FlatAppearance.BorderSize = 0
                AddHandler saveBtn.Click, Sub()
                    ' Guard: disable button immediately to prevent double-clicks causing cascading errors
                    saveBtn.Enabled = False
                    Try
                        Dim originalDate = hearing.NextHearing.Date
                        Dim selectedDate = datePicker.Value.Date

                        If originalDate <> selectedDate Then
                            ' 1. Update the original hearing's description fields and save it (retains history)
                            hearing.Hearing1 = hearing1Text.Text.Trim()
                            hearing.Hearing2 = hearing2Text.Text.Trim()
                            repository.UpdateHearing(hearing)

                            ' 2. Build log entry before creating the duplicate
                            Dim logEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} — Duplicated to {selectedDate:MMMM d, yyyy} from original date {originalDate:MMMM d, yyyy}"

                            ' 3. Write log entry to original record's cache
                            If Not _historyCache.ContainsKey(hearing.Id) Then
                                _historyCache(hearing.Id) = New List(Of String)()
                            End If
                            _historyCache(hearing.Id).Add(logEntry)
                            hearing.HistoryLog = _historyCache(hearing.Id)

                            ' 4. Create a new hearing duplicate scheduled for the next target date.
                            '    Copy ALL fields from original (including Hearing1 / Hearing2)
                            Dim duplicatedHearing As New HearingRecord With {
                                .No = hearing.No,
                                .NameOfPdl = hearing.NameOfPdl,
                                .BrCourt = hearing.BrCourt,
                                .Hearing1 = hearing.Hearing1,
                                .Hearing2 = hearing.Hearing2,
                                .NextHearing = selectedDate
                            }
                            repository.AddHearing(duplicatedHearing)

                            ' 5. Write the same log entry to the new duplicate's cache (using its assigned Id)
                            Dim dupLogEntry = $"{DateTime.Now:MMM d, yyyy h:mm tt} — Created from {originalDate:MMMM d, yyyy}"
                            _historyCache(duplicatedHearing.Id) = New List(Of String)() From {dupLogEntry}
                            duplicatedHearing.HistoryLog = _historyCache(duplicatedHearing.Id)

                            ReloadCalendar()
                            DisplayHearing(hearing)

                            ' Update historyList directly inside the open dialog
                            If historyList IsNot Nothing Then
                                historyList.Items.Clear()
                                For i = hearing.HistoryLog.Count - 1 To 0 Step -1
                                    historyList.Items.Add(hearing.HistoryLog(i))
                                Next
                            End If

                            ' Re-enable the button after success so the user can make further changes
                            saveBtn.Enabled = True

                            MessageBox.Show(dlg, $"Hearing duplicated to {selectedDate:yyyy-MM-dd}. The original hearing has been updated and remains on {originalDate:yyyy-MM-dd} to preserve history.", "Saved & Duplicated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            ' Date unchanged — update the original record normally
                            hearing.Hearing1 = hearing1Text.Text.Trim()
                            hearing.Hearing2 = hearing2Text.Text.Trim()
                            hearing.NextHearing = selectedDate
                            repository.UpdateHearing(hearing)

                            ReloadCalendar()
                            DisplayHearing(hearing)

                            ' Re-enable the button after success
                            saveBtn.Enabled = True

                            MessageBox.Show(dlg, "Hearing changes saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                    Catch ex As Exception
                        ' Re-enable so the user can try again (e.g. after closing a locked file)
                        saveBtn.Enabled = True
                        MessageBox.Show(dlg, ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End Sub

                Dim closeBtn As New Button With {
                    .Text = "Close",
                    .Location = New Point(340, 0),
                    .Size = New Size(76, 34),
                    .BackColor = Color.FromArgb(237, 233, 254),
                    .ForeColor = Color.FromArgb(30, 27, 75),
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Segoe UI", 9.0F)
                }
                closeBtn.FlatAppearance.BorderSize = 0
                AddHandler closeBtn.Click, Sub() dlg.Close()

                btnPanel.Controls.Add(deleteBtn)
                btnPanel.Controls.Add(saveBtn)
                btnPanel.Controls.Add(closeBtn)

                bodyPanel.Controls.Add(grid, 0, 0)
                bodyPanel.Controls.Add(detailWarningLabel, 0, 1)
                bodyPanel.Controls.Add(divLine, 0, 2)
                bodyPanel.Controls.Add(btnPanel, 0, 3)

                ' ── Other Hearings Section ───────────────────────────────────
                Dim otherSpacer As New Panel With {
                    .Dock = DockStyle.Fill,
                    .BackColor = Color.Transparent
                }
                bodyPanel.Controls.Add(otherSpacer, 0, 4)

                Dim otherContainer As New Panel With {
                    .Dock = DockStyle.Fill,
                    .BackColor = Color.FromArgb(250, 249, 255),
                    .Padding = New Padding(10, 8, 10, 8),
                    .Margin = New Padding(0)
                }
                AddHandler otherContainer.Paint, Sub(s, ev)
                                                     Dim rect = New Rectangle(0, 0, otherContainer.Width - 1, otherContainer.Height - 1)
                                                     ev.Graphics.DrawRectangle(New Pen(Color.FromArgb(221, 214, 254)), rect)
                                                 End Sub

                Dim otherTitle As New Label With {
                    .Text = "📅  ALL SCHEDULED HEARINGS FOR THIS PERSON",
                    .Dock = DockStyle.Top,
                    .Height = 22,
                    .ForeColor = Color.FromArgb(107, 33, 168),
                    .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
                    .Padding = New Padding(0, 0, 0, 4)
                }

                AddHandler otherHearingsList.Format, Sub(s, ev)
                    Dim h = TryCast(ev.ListItem, HearingRecord)
                    If h IsNot Nothing Then
                        Dim dtStr = If(h.NextHearing = Date.MinValue, "Pending", h.NextHearing.ToString("yyyy-MM-dd"))
                        ev.Value = $"📅 {dtStr} - {h.BrCourt} ({h.Hearing1})"
                    End If
                End Sub

                AddHandler otherHearingsList.DoubleClick, Sub()
                    Dim selected = TryCast(otherHearingsList.SelectedItem, HearingRecord)
                    If selected IsNot Nothing AndAlso selected.Id <> hearing.Id Then
                        currentHearing = selected ' loop again for the new selection
                        dlg.DialogResult = DialogResult.OK
                        dlg.Close()
                    End If
                End Sub

                Dim samePersonHearingsList = hearings.Where(Function(h) String.Equals(h.NameOfPdl, hearing.NameOfPdl, StringComparison.OrdinalIgnoreCase)).OrderBy(Function(h) h.NextHearing).ToList()
                otherHearingsList.Items.Clear()
                Dim currentIdx As Integer = -1
                For i = 0 To samePersonHearingsList.Count - 1
                    Dim h = samePersonHearingsList(i)
                    otherHearingsList.Items.Add(h)
                    If h.Id = hearing.Id Then
                        currentIdx = i
                    End If
                Next
                If currentIdx >= 0 Then
                    otherHearingsList.SelectedIndex = currentIdx
                End If

                otherContainer.Controls.Add(otherHearingsList)
                otherContainer.Controls.Add(otherTitle)
                bodyPanel.Controls.Add(otherContainer, 0, 5)

                ' ── History Log Section ─────────────────────────────────────
                If hearing.HistoryLog IsNot Nothing Then
                    Dim historySpacer As New Panel With {
                        .Dock = DockStyle.Fill,
                        .BackColor = Color.Transparent
                    }
                    bodyPanel.Controls.Add(historySpacer, 0, 6)

                    Dim historyContainer As New Panel With {
                        .Dock = DockStyle.Fill,
                        .BackColor = Color.FromArgb(250, 249, 255),
                        .Padding = New Padding(10, 8, 10, 8),
                        .Margin = New Padding(0)
                    }
                    AddHandler historyContainer.Paint, Sub(s, ev)
                                                           Dim rect = New Rectangle(0, 0, historyContainer.Width - 1, historyContainer.Height - 1)
                                                           ev.Graphics.DrawRectangle(New Pen(Color.FromArgb(221, 214, 254)), rect)
                                                       End Sub

                    Dim historyTitle As New Label With {
                        .Text = "📋  HISTORY LOG",
                        .Dock = DockStyle.Top,
                        .Height = 22,
                        .ForeColor = Color.FromArgb(107, 33, 168),
                        .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
                        .Padding = New Padding(0, 0, 0, 4)
                    }

                    ' Add history entries in reverse chronological order (newest first)
                    historyList.Items.Clear()
                    For i = hearing.HistoryLog.Count - 1 To 0 Step -1
                        historyList.Items.Add(hearing.HistoryLog(i))
                    Next

                    historyContainer.Controls.Add(historyList)
                    historyContainer.Controls.Add(historyTitle)
                    bodyPanel.Controls.Add(historyContainer, 0, 7)
                End If

                ' Main layout to prevent overlap between header and scroll body
                Dim mainLayout As New TableLayoutPanel With {
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 2,
                    .Padding = New Padding(0),
                    .Margin = New Padding(0)
                }
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
            ' Search mode: group by person — show ONE entry per unique person (earliest upcoming hearing)
            Dim allMatches = hearings.Where(Function(h)
                Return h.NameOfPdl.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0
            End Function)
            ' Pick the representative record for each person: prefer earliest future hearing, else earliest any
            filtered = allMatches.
                GroupBy(Function(h) h.NameOfPdl.Trim().ToUpperInvariant()).
                Select(Function(g)
                    Dim upcoming = g.Where(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date >= Date.Today).
                                    OrderBy(Function(h) h.NextHearing).FirstOrDefault()
                    Return If(upcoming, g.OrderBy(Function(h) h.NextHearing).First())
                End Function).
                OrderBy(Function(h) h.NameOfPdl).
                ToList()
            Dim uniqueCount = filtered.Count
            dateListTitleLabel.Text = $"Search results — {uniqueCount} person(s)"

        ElseIf dateWasClicked Then
            ' Date selected mode: show only that date's hearings
            filtered = hearings.Where(Function(h)
                Return h.NextHearing.Date = selectedDate.Date
            End Function).
                OrderBy(Function(h) h.NameOfPdl).
                ToList()
            dateListTitleLabel.Text = $"Hearings on {selectedDate:yyyy-MM-dd} ({filtered.Count})"

        Else
            ' Default mode: show ALL hearings (past + future + unscheduled)
            filtered = hearings.
                OrderBy(Function(h) If(h.NextHearing = Date.MinValue, 1, 0)).
                ThenByDescending(Function(h) h.NextHearing).
                ThenBy(Function(h) h.NameOfPdl).
                ToList()
            Dim pastCount = filtered.Where(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date < Date.Today).Count()
            Dim upcomingCount = filtered.Where(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date >= Date.Today).Count()
            dateListTitleLabel.Text = $"All Hearings — {upcomingCount} upcoming, {pastCount} past"
        End If

        hearingListBox.BeginUpdate()
        hearingListBox.Items.Clear()
        For Each h In filtered
            hearingListBox.Items.Add(h)
        Next
        hearingListBox.EndUpdate()

        ' Auto-select and display first result when searching
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
                ' Search mode: one entry per person, no date shown to avoid confusion
                e.Value = $"{hearing.No} - {hearing.NameOfPdl}"
            Else
                ' Date / default mode: show date so multiple hearings on different dates are distinguishable
                Dim dateStr = If(hearing.NextHearing = Date.MinValue, "Pending", hearing.NextHearing.ToString("yyyy-MM-dd"))
                e.Value = $"{hearing.No} - {hearing.NameOfPdl} [{dateStr}]"
            End If
        End If
    End Sub

    Private Sub ImportDataFile()
        ' ── Step 1: Ask user where to save a backup of current data ──────────
        Dim backupPath As String = ""
        Using saveDlg As New SaveFileDialog()
            saveDlg.Title = "Save backup of current data before importing"
            saveDlg.Filter = "XML files (*.xml)|*.xml"
            saveDlg.FileName = $"hearings_backup_{DateTime.Now:yyyy-MM-dd}"
            saveDlg.DefaultExt = "xml"
            saveDlg.OverwritePrompt = True

            Dim saveResult = saveDlg.ShowDialog(Me)
            If saveResult = DialogResult.Cancel Then
                Return  ' User cancelled entirely
            End If
            If saveResult = DialogResult.OK Then
                backupPath = saveDlg.FileName
            End If
        End Using

        ' ── Step 2: Pick the new XML / Excel file to import ───────────────────
        Using openDlg As New OpenFileDialog()
            openDlg.Title = "Import hearing data"
            openDlg.Filter = "Excel or XML files (*.xlsx;*.xlsm;*.xml)|*.xlsx;*.xlsm;*.xml"
            If openDlg.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            Try
                ShowProgress("Backing up current data...")
                ' Save backup to user-chosen location (if they gave one)
                If Not String.IsNullOrEmpty(backupPath) Then
                    repository.BackupCurrentData(backupPath)
                End If

                ShowProgress("Importing hearing data...")
                repository.ImportDataFile(openDlg.FileName)
                workbookValueLabel.Text = repository.WorkbookPath

                ShowProgress("Refreshing calendar...")
                ReloadCalendar()

                Dim scheduledCount = repository.CountSchedulableHearings()
                Dim backupNote As String = If(Not String.IsNullOrEmpty(backupPath),
                    $"{Environment.NewLine}{Environment.NewLine}✔ Previous data backed up to:{Environment.NewLine}  {backupPath}",
                    "")

                HideProgress()
                MessageBox.Show(Me,
                    $"Import complete! {scheduledCount} hearing(s) now in the master file." & backupNote,
                    "Import Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Catch ex As Exception
                HideProgress()
                MessageBox.Show(Me,
                    $"Import failed:{Environment.NewLine}{ex.Message}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Async Function ExportAllHearings() As Task
        Using dialog As New SaveFileDialog()
            dialog.Title = "Export Master Hearings"
            dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|CSV UTF-8 (*.csv)|*.csv|PDF Document (*.pdf)|*.pdf"
            dialog.FileName = $"BJMP-Hearings-All-{DateTime.Now:yyyy-MM-dd}"
            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            Dim ext = Path.GetExtension(dialog.FileName).ToLowerInvariant()
            Try
                ShowProgress("Loading hearing data...")
                Dim currentHearings = repository.LoadHearings()
                Dim exported = currentHearings
                If ext = ".xlsx" Then
                    exported = currentHearings.Where(Function(h) h.NextHearing <> Date.MinValue).ToList()
                End If

                _lastExportedIds.Clear()
                _lastExportedIds.AddRange(exported.Select(Function(h) h.Id))

                Select Case ext
                    Case ".xlsx"
                        ShowProgress("Exporting to Excel...")
                        repository.ExportToExcel(dialog.FileName)
                        HideProgress()
                        MessageBox.Show(Me, "Excel exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case ".csv"
                        ShowProgress("Exporting to CSV...")
                        repository.ExportToCsv(dialog.FileName)
                        HideProgress()
                        MessageBox.Show(Me, "CSV exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Case ".pdf"
                        ShowProgress("Generating PDF report...")
                        Dim allHearings = repository.LoadHearings().OrderBy(Function(h) h.NextHearing).ThenBy(Function(h) h.NameOfPdl).ToList()
                        Dim titleText = "BJMP Hearing Schedule (All Records)"
                        Dim subtitleText = $"Exported on: {DateTime.Now:yyyy-MM-dd HH:mm}"
                        Dim reportHtml = BuildPdfReport(titleText, subtitleText, allHearings)
                        Dim reportView As New WebView2 With {
                            .Visible = False,
                            .Size = New Size(900, 1100)
                        }
                        Controls.Add(reportView)
                        Try
                            Await reportView.EnsureCoreWebView2Async()
                            Dim loaded As New TaskCompletionSource(Of Boolean)()
                            AddHandler reportView.CoreWebView2.NavigationCompleted,
                                Sub()
                                    loaded.TrySetResult(True)
                                End Sub
                            reportView.NavigateToString(reportHtml)
                            Await loaded.Task
                            ShowProgress("Saving PDF...")
                            Await reportView.CoreWebView2.PrintToPdfAsync(dialog.FileName)
                            HideProgress()
                            MessageBox.Show(Me, "PDF exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Finally
                            Controls.Remove(reportView)
                            reportView.Dispose()
                        End Try
                End Select
            Catch ex As Exception
                HideProgress()
                MessageBox.Show(Me, $"Export failed:{Environment.NewLine}{ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Function

    Private Shared Function BuildPdfReport(titleText As String, subtitleText As String, records As List(Of HearingRecord)) As String
        Dim rows = String.Join("", records.Select(Function(hearing)
            Return $"<tr><td>{WebUtility.HtmlEncode(hearing.No)}</td><td>{WebUtility.HtmlEncode(hearing.NameOfPdl)}</td><td>{WebUtility.HtmlEncode(hearing.BrCourt)}</td><td>{WebUtility.HtmlEncode(hearing.Hearing1)}</td><td>{WebUtility.HtmlEncode(hearing.Hearing2)}</td><td>{If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("yyyy-MM-dd"))}</td></tr>"
        End Function))

        If rows = "" Then
            rows = "<tr><td colspan=""6"" class=""empty"">No hearings scheduled for this date range.</td></tr>"
        End If

        Return String.Join(Environment.NewLine, {
            "<!doctype html><html><head><meta charset=""utf-8""><style>",
            "body { font-family: 'Segoe UI', Arial, sans-serif; color: #1f2937; margin: 36px; }",
            "h1 { color: #12365d; margin: 0 0 4px; font-size: 24px; }",
            ".subtitle { color: #526173; margin-bottom: 24px; }",
            "table { width: 100%; border-collapse: collapse; font-size: 12px; }",
            "th { background: #12365d; color: white; text-align: left; padding: 8px; }",
            "td { border: 1px solid #d9e1ea; padding: 7px; vertical-align: top; }",
            ".empty { text-align: center; color: #526173; padding: 20px; }",
            "</style></head><body>",
            $"<h1>{WebUtility.HtmlEncode(titleText)}</h1>",
            $"<div class=""subtitle"">{WebUtility.HtmlEncode(subtitleText)}</div>",
            $"<table><thead><tr><th>NO</th><th>NAME OF PDL</th><th>BR/COURT</th><th>HEARING</th><th>HEARING</th><th>NEXT HEARING</th></tr></thead><tbody>{rows}</tbody></table>",
            "</body></html>"
        })
    End Function


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

            ' Header strip
            Dim titlePanel As New Panel With {
                .Height = 56,
                .BackColor = Color.FromArgb(245, 158, 11),
                .Dock = DockStyle.Top
            }
            Dim titleLbl As New Label With {
                .Text = "⚙  SETTINGS & CLEAR DATA",
                .Dock = DockStyle.Fill,
                .ForeColor = Color.FromArgb(30, 27, 75),
                .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleCenter
            }
            titlePanel.Controls.Add(titleLbl)

            Dim container As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20),
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False
            }

            ' Description Label
            Dim descLbl As New Label With {
                .Text = "Select an option below to clear/clean hearing data. A backup file is automatically generated before clearing.",
                .Width = 360,
                .Height = 45,
                .ForeColor = Color.FromArgb(139, 92, 246),
                .Font = New Font("Segoe UI", 9.0F)
            }
            container.Controls.Add(descLbl)

            Dim btnWidth = 350
            Dim btnHeight = 36

            ' Button 1: Clear All Data
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
                        MessageBox.Show(dlg, $"Failed to clear hearings:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Sub
            container.Controls.Add(btnClearAll)

            Dim spacer1 As New Panel With { .Height = 10 }
            container.Controls.Add(spacer1)

            ' Button 2: Clear Data by Month
            Dim btnClearMonth = MakeSideButton("Clear Hearings by Month", Color.FromArgb(109, 40, 217), Color.White)
            btnClearMonth.Width = btnWidth
            btnClearMonth.Height = btnHeight
            AddHandler btnClearMonth.Click, Sub()
                Dim months = hearings.Where(Function(h) h.NextHearing <> Date.MinValue).
                    Select(Function(h) New DateTime(h.NextHearing.Year, h.NextHearing.Month, 1)).
                    Distinct().
                    OrderBy(Function(d) d).
                    ToList()

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

                    Dim lbl As New Label With { .Text = "Select Month:", .Location = New Point(20, 20), .AutoSize = True }
                    Dim cb As New ComboBox With { .Location = New Point(20, 45), .Width = 260, .DropDownStyle = ComboBoxStyle.DropDownList }
                    For Each m In months
                        cb.Items.Add(m.ToString("MMMM yyyy", CultureInfo.InvariantCulture))
                    Next
                    cb.SelectedIndex = 0

                    Dim okBtn As New Button With { .Text = "Clear Month", .Location = New Point(60, 90), .Width = 100, .DialogResult = DialogResult.OK }
                    Dim cancelBtn As New Button With { .Text = "Cancel", .Location = New Point(170, 90), .Width = 80, .DialogResult = DialogResult.Cancel }

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
                                MessageBox.Show(dlg, $"Failed to clear month:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                        End If
                    End If
                End Using
            End Sub
            container.Controls.Add(btnClearMonth)

            Dim spacer2 As New Panel With { .Height = 10 }
            container.Controls.Add(spacer2)

            ' Button 3: Clear Data by Date Range
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

                    Dim lblStart As New Label With { .Text = "Start Date:", .Location = New Point(20, 20), .AutoSize = True }
                    Dim dpStart As New DateTimePicker With { .Location = New Point(20, 45), .Width = 130, .Format = DateTimePickerFormat.Short }
                    Dim lblEnd As New Label With { .Text = "End Date:", .Location = New Point(170, 20), .AutoSize = True }
                    Dim dpEnd As New DateTimePicker With { .Location = New Point(170, 45), .Width = 130, .Format = DateTimePickerFormat.Short }

                    Dim okBtn As New Button With { .Text = "Clear Range", .Location = New Point(80, 110), .Width = 100, .DialogResult = DialogResult.OK }
                    Dim cancelBtn As New Button With { .Text = "Cancel", .Location = New Point(190, 110), .Width = 80, .DialogResult = DialogResult.Cancel }

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
                                ShowProgress($"Clearing hearings {startD:MMM d} – {endD:MMM d, yyyy}...")
                                repository.BackupCurrentData()
                                repository.ClearHearings(Function(h) h.NextHearing <> Date.MinValue AndAlso h.NextHearing.Date >= startD AndAlso h.NextHearing.Date <= endD)
                                ReloadCalendar()
                                HideProgress()
                                MessageBox.Show(dlg, $"Cleared hearings between {startD:yyyy-MM-dd} and {endD:yyyy-MM-dd}. Backup created.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Catch ex As Exception
                                HideProgress()
                                MessageBox.Show(dlg, $"Failed to clear date range:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                        End If
                    End If
                End Using
            End Sub
            container.Controls.Add(btnClearRange)

            Dim spacer3 As New Panel With { .Height = 10 }
            container.Controls.Add(spacer3)

            ' Button 4: Clear Last Exported Data
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
                        MessageBox.Show(dlg, $"Failed to clear last export:{Environment.NewLine}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        _statusLabel.Text = "⏳  " & message
        If _progressPanel IsNot Nothing Then
            _progressPanel.Visible = True
            _progressPanel.Refresh()
        End If
        _statusLabel.Refresh()
    End Sub

    ''' <summary>Hides the progress bar.</summary>
    Private Sub HideProgress()
        If InvokeRequired Then
            Invoke(Sub() HideProgress())
            Return
        End If
        If _progressPanel IsNot Nothing Then _progressPanel.Visible = False
    End Sub

End Class
