Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text.Encodings.Web
Imports System.Text.Json
Imports System.Threading.Tasks
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
        BackColor = Color.FromArgb(236, 241, 247)
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
        webView.DefaultBackgroundColor = Color.FromArgb(236, 241, 247)

        ' ── Sidebar ──────────────────────────────────────────
        detailsPanel.Dock = DockStyle.Fill
        detailsPanel.BackColor = Color.FromArgb(247, 250, 253)
        detailsPanel.Padding = New Padding(0)

        ' Header strip
        Dim headerStrip As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .BackColor = Color.FromArgb(18, 54, 93)
        }
        Dim headerLbl As New Label With {
            .Text = "BJMP  HEARING PANEL",
            .Dock = DockStyle.Fill,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Padding = New Padding(0, 0, 0, 2)
        }
        Dim goldBar As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 4,
            .BackColor = Color.FromArgb(242, 201, 76)
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
            .ForeColor = Color.FromArgb(96, 108, 123),
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
        Dim div1 As New Panel With { .Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(218, 226, 236) }

        ' ── Hearings list section ───────────────
        Dim listSection As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 200,
            .Padding = New Padding(0, 8, 0, 8)
        }
        dateListTitleLabel.Dock = DockStyle.Top
        dateListTitleLabel.Height = 22
        dateListTitleLabel.Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        dateListTitleLabel.ForeColor = Color.FromArgb(96, 108, 123)
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
        listSection.Controls.Add(hearingListBox)
        listSection.Controls.Add(dateListTitleLabel)

        ' Divider
        Dim div2 As New Panel With { .Dock = DockStyle.Top, .Height = 1, .BackColor = Color.FromArgb(218, 226, 236) }

        ' ── Details card section ─────────────────
        Dim detailsCardLabel As New Label With {
            .Text = "SELECTED HEARING INFO",
            .Dock = DockStyle.Top,
            .Height = 22,
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(96, 108, 123),
            .Padding = New Padding(0, 8, 0, 0)
        }

        detailsTitleLabel.Dock = DockStyle.Top
        detailsTitleLabel.AutoSize = True
        detailsTitleLabel.Font = New Font("Segoe UI Semibold", 12.0F)
        detailsTitleLabel.ForeColor = Color.FromArgb(18, 54, 93)
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
            ev.Graphics.DrawRectangle(New System.Drawing.Pen(Color.FromArgb(218, 226, 236)), rect)
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
            .ForeColor = Color.FromArgb(120, 130, 145),
            .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        }
        workbookValueLabel.AutoSize = True
        workbookValueLabel.Font = New Font("Segoe UI", 7.5F)
        workbookValueLabel.ForeColor = Color.FromArgb(120, 130, 145)
        fileSection.Controls.Add(fileLbl)
        fileSection.Controls.Add(workbookValueLabel)

        detailsCard.Controls.Add(detailsTable)
        detailsCard.Controls.Add(detailsTitleLabel)
        detailsCard.Controls.Add(detailsCardLabel)

        innerPanel.Controls.Add(fileSection)
        innerPanel.Controls.Add(detailsCard)
        innerPanel.Controls.Add(div2)
        innerPanel.Controls.Add(listSection)
        innerPanel.Controls.Add(div1)
        innerPanel.Controls.Add(searchSection)

        ' ── Button strip at bottom ────────────────
        Dim btnStrip As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 172,
            .BackColor = Color.FromArgb(236, 241, 247),
            .Padding = New Padding(12, 8, 12, 10)
        }

        Dim addButton = MakeSideButton("＋  Add New Hearing", Color.FromArgb(18, 54, 93), Color.White)
        addButton.Dock = DockStyle.Top
        addButton.Height = 38
        AddHandler addButton.Click, Sub() ShowAddDialog(selectedDate)

        Dim spacer1 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim importButton = MakeSideButton("⬆  Import Excel / XML", Color.FromArgb(52, 73, 102), Color.White)
        importButton.Dock = DockStyle.Top
        importButton.Height = 36
        AddHandler importButton.Click, Sub() ImportDataFile()

        Dim spacer2 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim exportButton = MakeSideButton("⬇  Export Date to PDF", Color.FromArgb(242, 201, 76), Color.FromArgb(31, 41, 55))
        exportButton.Dock = DockStyle.Top
        exportButton.Height = 36
        AddHandler exportButton.Click, Async Sub()
            Try
                Await ExportSelectedDatePdf()
            Catch ex As Exception
                MessageBox.Show(Me, "Failed to export PDF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Dim spacer3 As New Panel With { .Dock = DockStyle.Top, .Height = 6, .BackColor = Color.Transparent }

        Dim refreshButton = MakeSideButton("↺  Refresh Calendar", Color.FromArgb(218, 226, 236), Color.FromArgb(38, 50, 66))
        refreshButton.Dock = DockStyle.Top
        refreshButton.Height = 34
        AddHandler refreshButton.Click, Sub() ReloadCalendar()

        btnStrip.Controls.Add(refreshButton)
        btnStrip.Controls.Add(spacer3)
        btnStrip.Controls.Add(exportButton)
        btnStrip.Controls.Add(spacer2)
        btnStrip.Controls.Add(importButton)
        btnStrip.Controls.Add(spacer1)
        btnStrip.Controls.Add(addButton)

        detailsPanel.Controls.Add(innerPanel)
        detailsPanel.Controls.Add(btnStrip)
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
            .ForeColor = Color.FromArgb(100, 116, 139),
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(0, 6, 0, 6)
        }
        valueLabel.Text = "—"
        valueLabel.AutoSize = True
        valueLabel.Dock = DockStyle.Fill
        valueLabel.ForeColor = Color.FromArgb(31, 41, 55)
        valueLabel.Font = New Font("Segoe UI", 9.5F)
        valueLabel.TextAlign = ContentAlignment.MiddleLeft
        valueLabel.Padding = New Padding(0, 6, 0, 6)
        table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        table.Controls.Add(captionLbl)
        table.Controls.Add(valueLabel)
    End Sub

    Private Sub CalendarMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Using document = JsonDocument.Parse(e.WebMessageAsJson)
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
                        SelectHearing(root.GetProperty("id").GetString())
                    Case "move"
                        MoveHearing(root.GetProperty("id").GetString(), Date.Parse(root.GetProperty("date").GetString()))
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
            RefreshSideList()
            Dim calendarEvents = hearings.
                Where(Function(hearing) hearing.NextHearing <> Date.MinValue).
                Select(Function(hearing) New With {
                    .id = hearing.Id.ToString(),
                    .title = $"{hearing.No} - {hearing.NameOfPdl}",
                    .start = hearing.NextHearing.ToString("yyyy-MM-dd"),
                    .color = "#12365d",
                    .textColor = "#ffffff",
                    .extendedProps = New With {
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
        Using dialog As New AddEditHearingForm(selectedDate)
            If dialog.ShowDialog(Me) = DialogResult.OK Then
                Dim saved = repository.AddHearing(dialog.Hearing)
                ReloadCalendar()
                DisplayHearing(saved)
            End If
        End Using
    End Sub

    Private Sub ShowDatePopup(clickedDate As Date)
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
                .BackColor = Color.FromArgb(18, 54, 93),
                .Padding = New Padding(18, 0, 18, 0)
            }
            Dim titleLbl As New Label With {
                .Text = $"📅  {clickedDate:dddd, MMMM d, yyyy}",
                .Dock = DockStyle.Fill,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI Semibold", 13.0F),
                .TextAlign = ContentAlignment.MiddleLeft
            }
            Dim countLbl As New Label With {
                .Text = $"{dateHearings.Count} hearing(s)",
                .Dock = DockStyle.Right,
                .Width = 130,
                .ForeColor = Color.FromArgb(242, 201, 76),
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
                    .ForeColor = Color.FromArgb(120, 130, 145),
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
                        .BackColor = Color.FromArgb(18, 54, 93)
                    }
                    ' Make dot round-ish with a border
                    Dim noLbl As New Label With {
                        .Text = h.No,
                        .Location = New Point(32, 0),
                        .Size = New Size(38, 38),
                        .TextAlign = ContentAlignment.MiddleCenter,
                        .ForeColor = Color.FromArgb(96, 108, 123),
                        .Font = New Font("Segoe UI", 9.0F)
                    }
                    Dim nameLbl As New Label With {
                        .Text = h.NameOfPdl,
                        .Location = New Point(72, 0),
                        .Size = New Size(230, 38),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .ForeColor = Color.FromArgb(18, 54, 93),
                        .Font = New Font("Segoe UI Semibold", 10.0F)
                    }
                    Dim courtLbl As New Label With {
                        .Text = h.BrCourt,
                        .Location = New Point(304, 0),
                        .Size = New Size(100, 38),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .ForeColor = Color.FromArgb(70, 80, 94),
                        .Font = New Font("Segoe UI", 9.5F)
                    }
                    Dim viewBtn As New Button With {
                        .Text = "View",
                        .Location = New Point(410, 6),
                        .Size = New Size(55, 26),
                        .BackColor = Color.FromArgb(18, 54, 93),
                        .ForeColor = Color.White,
                        .FlatStyle = FlatStyle.Flat,
                        .Font = New Font("Segoe UI", 8.5F)
                    }
                    viewBtn.FlatAppearance.BorderSize = 0
                    AddHandler viewBtn.Click, Sub()
                        Dim originalDate = currentH.NextHearing.Date
                        popup.Close()
                        ShowHearingDetailPopup(currentH)
                        ShowDatePopup(originalDate)
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
                .BackColor = Color.FromArgb(18, 54, 93),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI Semibold", 9.5F)
            }
            addBtn.FlatAppearance.BorderSize = 0
            AddHandler addBtn.Click, Sub()
                popup.Close()
                ShowAddDialog(clickedDate)
            End Sub
            Dim closeBtn As New Button With {
                .Text = "Close",
                .Dock = DockStyle.Right,
                .Width = 80,
                .BackColor = Color.FromArgb(235, 239, 244),
                .ForeColor = Color.FromArgb(38, 50, 66),
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
            popup.ShowDialog(Me)
        End Using
    End Sub


    Private Sub SelectDate(dateValue As Date)
        selectedDate = dateValue.Date
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

    Private Sub MoveHearing(idText As String, nextHearing As Date)
        Dim rowId As Integer
        If Not Integer.TryParse(idText, rowId) Then
            Return
        End If

        repository.MoveHearing(rowId, nextHearing.Date)
        ReloadCalendar()
        Dim moved = hearings.FirstOrDefault(Function(item) item.Id = rowId)
        If moved IsNot Nothing Then
            moved.NextHearing = nextHearing.Date
            DisplayHearing(moved)
        End If
    End Sub

    Private Sub DisplayHearing(hearing As HearingRecord)
        detailsTitleLabel.Text = hearing.NameOfPdl
        noValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.No), "—", hearing.No)
        nameValueLabel.Text = hearing.NameOfPdl
        courtValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.BrCourt), "—", hearing.BrCourt)
        hearing1ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing1), "—", hearing.Hearing1)
        hearing2ValueLabel.Text = If(String.IsNullOrWhiteSpace(hearing.Hearing2), "—", hearing.Hearing2)
        dateValueLabel.Text = If(hearing.NextHearing = Date.MinValue, "Pending / Unscheduled", hearing.NextHearing.ToString("MMMM d, yyyy"))
    End Sub

    Private Sub ShowHearingDetailPopup(hearing As HearingRecord)
        DisplayHearing(hearing)

        Using dlg As New Form()
            dlg.Text = hearing.NameOfPdl
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MinimizeBox = False
            dlg.MaximizeBox = False
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 10.0F)
            dlg.ClientSize = New Size(480, 400)

            ' Header bar
            Dim hdr As New Panel With {
                .Height = 64,
                .BackColor = Color.FromArgb(18, 54, 93),
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
                .BackColor = Color.FromArgb(242, 201, 76)
            }
            hdr.Controls.Add(hdrName)
            hdr.Controls.Add(hdrGold)

            ' Body container TableLayoutPanel (bulletproof layout!)
            Dim bodyPanel As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(24, 16, 24, 12),
                .Margin = New Padding(0),
                .AutoScroll = True,
                .BackColor = Color.White
            }
            bodyPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            bodyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Row 0: grid
            bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 25)) ' Row 1: divLine
            bodyPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 36)) ' Row 2: btnPanel

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
                .ForeColor = Color.FromArgb(100, 116, 139),
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(0, 6, 0, 6),
                .AutoSize = True
            }
            Dim val0 As New Label With {
                .Text = If(String.IsNullOrWhiteSpace(hearing.No), "—", hearing.No),
                .Dock = DockStyle.Fill,
                .ForeColor = Color.FromArgb(31, 41, 55),
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
                .ForeColor = Color.FromArgb(100, 116, 139),
                .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(0, 6, 0, 6),
                .AutoSize = True
            }
            Dim val1 As New Label With {
                .Text = If(String.IsNullOrWhiteSpace(hearing.BrCourt), "—", hearing.BrCourt),
                .Dock = DockStyle.Fill,
                .ForeColor = Color.FromArgb(31, 41, 55),
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
                .ForeColor = Color.FromArgb(100, 116, 139),
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
                .ForeColor = Color.FromArgb(100, 116, 139),
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
                .ForeColor = Color.FromArgb(100, 116, 139),
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

            ' Divider
            Dim divLine As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 1,
                .BackColor = Color.FromArgb(218, 226, 236),
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
                .ForeColor = Color.White,
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
                .BackColor = Color.FromArgb(18, 54, 93),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI Semibold", 9.0F)
            }
            saveBtn.FlatAppearance.BorderSize = 0
            AddHandler saveBtn.Click, Sub()
                Try
                    ' Update hearing record with new hearing texts and date
                    hearing.Hearing1 = hearing1Text.Text.Trim()
                    hearing.Hearing2 = hearing2Text.Text.Trim()
                    hearing.NextHearing = datePicker.Value.Date
                    
                    repository.UpdateHearing(hearing)
                    
                    ReloadCalendar()
                    DisplayHearing(hearing)
                    
                    MessageBox.Show(dlg, "Hearing changes saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    dlg.Close()
                Catch ex As Exception
                    MessageBox.Show(dlg, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Sub

            Dim closeBtn As New Button With {
                .Text = "Close",
                .Location = New Point(340, 0),
                .Size = New Size(76, 34),
                .BackColor = Color.FromArgb(235, 239, 244),
                .ForeColor = Color.FromArgb(38, 50, 66),
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9.0F)
            }
            closeBtn.FlatAppearance.BorderSize = 0
            AddHandler closeBtn.Click, Sub() dlg.Close()

            btnPanel.Controls.Add(deleteBtn)
            btnPanel.Controls.Add(saveBtn)
            btnPanel.Controls.Add(closeBtn)

            bodyPanel.Controls.Add(grid, 0, 0)
            bodyPanel.Controls.Add(divLine, 0, 1)
            bodyPanel.Controls.Add(btnPanel, 0, 2)

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
    End Sub



    Private Sub RefreshSideList()
        Dim searchText = searchTextBox.Text.Trim()
        Dim filtered = hearings.Where(Function(hearing)
            Dim matchesSearch = Not String.IsNullOrWhiteSpace(searchText) AndAlso
                hearing.NameOfPdl.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0
            Dim matchesDate = String.IsNullOrWhiteSpace(searchText) AndAlso hearing.NextHearing.Date = selectedDate.Date
            Return matchesDate OrElse matchesSearch
        End Function).
            OrderBy(Function(hearing) hearing.NextHearing).
            ThenBy(Function(hearing) hearing.NameOfPdl).
            ToList()

        If String.IsNullOrWhiteSpace(searchText) Then
            dateListTitleLabel.Text = $"Hearings on {selectedDate:yyyy-MM-dd} ({filtered.Count})"
        Else
            dateListTitleLabel.Text = $"Search results ({filtered.Count})"
        End If
        hearingListBox.BeginUpdate()
        hearingListBox.Items.Clear()
        For Each hearing In filtered
            hearingListBox.Items.Add(hearing)
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
            e.Value = $"{hearing.No} - {hearing.NameOfPdl}"
        End If
    End Sub

    Private Sub ImportDataFile()
        Using dialog As New OpenFileDialog()
            dialog.Title = "Import hearing data"
            dialog.Filter = "Excel or XML files (*.xlsx;*.xlsm;*.xml)|*.xlsx;*.xlsm;*.xml"
            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            repository.ImportDataFile(dialog.FileName)
            workbookValueLabel.Text = repository.WorkbookPath
            ReloadCalendar()
            Dim scheduledCount = repository.CountSchedulableHearings()
            MessageBox.Show(Me, $"Data imported and saved to the application Excel file.{Environment.NewLine}{scheduledCount} hearing(s) have a valid NEXT HEARING date and can appear on the calendar.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    Private Async Function ExportSelectedDatePdf() As Task
        Dim rangeType As String = "Daily"
        Dim startRange As Date = selectedDate.Date
        Dim endRange As Date = selectedDate.Date
        Dim titleText As String = "BJMP Hearing Schedule"
        Dim subtitleText As String = $"Selected date: {selectedDate:yyyy-MM-dd}"
        Dim defaultFilename As String = $"BJMP-Hearings-{selectedDate:yyyy-MM-dd}.pdf"

        Using rangeDlg As New Form()
            rangeDlg.Text = "Export PDF Range"
            rangeDlg.Size = New Size(380, 280)
            rangeDlg.StartPosition = FormStartPosition.CenterParent
            rangeDlg.FormBorderStyle = FormBorderStyle.FixedDialog
            rangeDlg.MinimizeBox = False
            rangeDlg.MaximizeBox = False
            rangeDlg.BackColor = Color.White
            rangeDlg.Font = New Font("Segoe UI", 10.0F)

            Dim topPanel As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 50,
                .BackColor = Color.FromArgb(18, 54, 93)
            }
            Dim headerLbl As New Label With {
                .Text = "Select Export Range",
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI Semibold", 12.0F),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Fill
            }
            topPanel.Controls.Add(headerLbl)

            Dim mainPanel As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(20, 15, 20, 15)
            }

            Dim lblIntro As New Label With {
                .Text = $"Export hearings based on the selected date: {selectedDate:yyyy-MM-dd}",
                .Dock = DockStyle.Top,
                .Height = 40,
                .ForeColor = Color.FromArgb(70, 80, 94),
                .Font = New Font("Segoe UI", 9.0F)
            }

            Dim optDaily As New RadioButton With {
                .Text = $"Daily (Only {selectedDate:yyyy-MM-dd})",
                .Dock = DockStyle.Top,
                .Height = 30,
                .Checked = True
            }

            ' Calculate weekly range (Monday to Sunday)
            Dim currentDay As DayOfWeek = selectedDate.DayOfWeek
            Dim daysToSubtract As Integer = If(currentDay = DayOfWeek.Sunday, 6, CInt(currentDay) - 1)
            Dim startOfWeek As Date = selectedDate.AddDays(-daysToSubtract).Date
            Dim endOfWeek As Date = startOfWeek.AddDays(6).Date

            Dim optWeekly As New RadioButton With {
                .Text = $"Weekly (Mon-Sun: {startOfWeek:MM-dd} to {endOfWeek:MM-dd})",
                .Dock = DockStyle.Top,
                .Height = 30
            }

            ' Calculate monthly range
            Dim startOfMonth As Date = New Date(selectedDate.Year, selectedDate.Month, 1)
            Dim endOfMonth As Date = startOfMonth.AddMonths(1).AddDays(-1)

            Dim optMonthly As New RadioButton With {
                .Text = $"Monthly (Entire {selectedDate:MMMM yyyy})",
                .Dock = DockStyle.Top,
                .Height = 30
            }

            Dim btnPanel As New Panel With {
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .Padding = New Padding(0, 5, 0, 0)
            }

            Dim btnOk As New Button With {
                .Text = "Export",
                .DialogResult = DialogResult.OK,
                .Dock = DockStyle.Right,
                .Width = 100,
                .BackColor = Color.FromArgb(18, 54, 93),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }
            btnOk.FlatAppearance.BorderSize = 0

            Dim btnCancel As New Button With {
                .Text = "Cancel",
                .DialogResult = DialogResult.Cancel,
                .Dock = DockStyle.Left,
                .Width = 100,
                .BackColor = Color.FromArgb(235, 239, 244),
                .ForeColor = Color.FromArgb(38, 50, 66),
                .FlatStyle = FlatStyle.Flat
            }
            btnCancel.FlatAppearance.BorderSize = 0

            btnPanel.Controls.Add(btnOk)
            btnPanel.Controls.Add(btnCancel)

            mainPanel.Controls.Add(optMonthly)
            mainPanel.Controls.Add(optWeekly)
            mainPanel.Controls.Add(optDaily)
            mainPanel.Controls.Add(lblIntro)

            rangeDlg.Controls.Add(mainPanel)
            rangeDlg.Controls.Add(btnPanel)
            rangeDlg.Controls.Add(topPanel)

            If rangeDlg.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            If optDaily.Checked Then
                rangeType = "Daily"
                startRange = selectedDate.Date
                endRange = selectedDate.Date
                titleText = "BJMP Hearing Schedule"
                subtitleText = $"Selected date: {selectedDate:yyyy-MM-dd}"
                defaultFilename = $"BJMP-Hearings-{selectedDate:yyyy-MM-dd}.pdf"
            ElseIf optWeekly.Checked Then
                rangeType = "Weekly"
                startRange = startOfWeek
                endRange = endOfWeek
                titleText = "BJMP Hearing Schedule (Weekly)"
                subtitleText = $"Week: {startOfWeek:MMMM d, yyyy} (Monday) to {endOfWeek:MMMM d, yyyy} (Sunday)"
                defaultFilename = $"BJMP-Hearings-Week-{startOfWeek:yyyy-MM-dd}-to-{endOfWeek:yyyy-MM-dd}.pdf"
            ElseIf optMonthly.Checked Then
                rangeType = "Monthly"
                startRange = startOfMonth
                endRange = endOfMonth
                titleText = "BJMP Hearing Schedule (Monthly)"
                subtitleText = $"Month: {selectedDate:MMMM yyyy}"
                defaultFilename = $"BJMP-Hearings-Month-{selectedDate:yyyy-MM}.pdf"
            End If
        End Using

        Dim dateHearings = hearings.Where(Function(hearing) hearing.NextHearing.Date >= startRange AndAlso hearing.NextHearing.Date <= endRange).
            OrderBy(Function(hearing) hearing.NextHearing.Date).
            ThenBy(Function(hearing) hearing.NameOfPdl).ToList()

        Using dialog As New SaveFileDialog()
            dialog.Title = "Export Hearings to PDF"
            dialog.Filter = "PDF files (*.pdf)|*.pdf"
            dialog.FileName = defaultFilename
            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            Dim reportHtml = BuildPdfReport(titleText, subtitleText, dateHearings)
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
                Await reportView.CoreWebView2.PrintToPdfAsync(dialog.FileName)
                MessageBox.Show(Me, "PDF exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Finally
                Controls.Remove(reportView)
                reportView.Dispose()
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


End Class
