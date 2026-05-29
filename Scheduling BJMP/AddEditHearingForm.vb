Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class AddEditHearingForm
    Inherits Form

    Private ReadOnly noTextBox As New TextBox()
    Private ReadOnly nameTextBox As New TextBox()
    Private ReadOnly courtTextBox As New TextBox()
    Private ReadOnly hearing1TextBox As New TextBox()
    Private ReadOnly hearing2TextBox As New TextBox()
    Private ReadOnly nextHearingPicker As New DateTimePicker()
    
    Private ReadOnly _existingHearings As List(Of HearingRecord)
    Private _editingId As Integer = 0
    Private ReadOnly warningLabel As New Label()

    Public Sub New(selectedDate As Date, Optional existingHearings As List(Of HearingRecord) = Nothing)
        _existingHearings = If(existingHearings, New List(Of HearingRecord)())
        Text = "Add Hearing Information"
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ClientSize = New Size(590, 440) ' slightly increased height for warning label
        Font = New Font("Segoe UI", 10.0F)
        BackColor = Color.White

        Dim titleLabel As New Label With {
            .Text = "Hearing Information",
            .Font = New Font("Segoe UI Semibold", 15.0F),
            .ForeColor = Color.FromArgb(30, 27, 75),
            .Location = New Point(24, 20),
            .AutoSize = True
        }

        Dim table As New TableLayoutPanel With {
            .Location = New Point(24, 70),
            .Size = New Size(542, 250),
            .ColumnCount = 2,
            .RowCount = 6
        }
        table.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 165))
        table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

        ConfigureInput(noTextBox)
        ConfigureInput(nameTextBox)
        ConfigureInput(courtTextBox)
        ConfigureInput(hearing1TextBox)
        ConfigureInput(hearing2TextBox)

        nextHearingPicker.Format = DateTimePickerFormat.Custom
        nextHearingPicker.CustomFormat = "yyyy-MM-dd"
        nextHearingPicker.Value = selectedDate.Date

        AddRow(table, 0, "NO", noTextBox)
        AddRow(table, 1, "NAME OF PDL", nameTextBox)
        AddRow(table, 2, "BR/COURT", courtTextBox)
        AddRow(table, 3, "HEARING", hearing1TextBox)
        AddRow(table, 4, "HEARING", hearing2TextBox)
        AddRow(table, 5, "NEXT HEARING", nextHearingPicker)

        ' Warning Label for same-day duplicates
        warningLabel.ForeColor = Color.FromArgb(198, 40, 40)
        warningLabel.Font = New Font("Segoe UI Semibold", 9.0F)
        warningLabel.Location = New Point(24, 330)
        warningLabel.Size = New Size(542, 22)
        warningLabel.Text = ""
        warningLabel.Visible = False

        Dim checkDuplicate = Sub()
            Dim nameText = nameTextBox.Text.Trim()
            Dim targetDate = nextHearingPicker.Value.Date
            If String.IsNullOrWhiteSpace(nameText) OrElse _existingHearings.Count = 0 Then
                warningLabel.Visible = False
                Return
            End If

            Dim hasDup = _existingHearings.Any(Function(h)
                Return h.Id <> _editingId AndAlso
                       String.Equals(h.NameOfPdl.Trim(), nameText, StringComparison.OrdinalIgnoreCase) AndAlso
                       h.NextHearing.Date = targetDate
            End Function)

            If hasDup Then
                warningLabel.Text = "⚠️ Note: This person already has a hearing scheduled on this date!"
                warningLabel.Visible = True
            Else
                warningLabel.Visible = False
            End If
        End Sub

        AddHandler nameTextBox.TextChanged, Sub() checkDuplicate()
        AddHandler nextHearingPicker.ValueChanged, Sub() checkDuplicate()

        Dim saveButton As New Button With {
            .Text = "Save",
            .BackColor = Color.FromArgb(245, 158, 11),
            .ForeColor = Color.FromArgb(30, 27, 75),
            .FlatStyle = FlatStyle.Flat,
            .Location = New Point(376, 370),
            .Size = New Size(90, 36),
            .DialogResult = DialogResult.OK
        }
        saveButton.FlatAppearance.BorderSize = 0

        Dim cancelButton As New Button With {
            .Text = "Cancel",
            .BackColor = Color.FromArgb(245, 243, 255),
            .ForeColor = Color.FromArgb(30, 27, 75),
            .FlatStyle = FlatStyle.Flat,
            .Location = New Point(476, 370),
            .Size = New Size(90, 36),
            .DialogResult = DialogResult.Cancel
        }
        cancelButton.FlatAppearance.BorderSize = 0

        AcceptButton = saveButton
        CancelButton = cancelButton
        Controls.AddRange(New Control() {titleLabel, table, warningLabel, saveButton, cancelButton})
    End Sub

    Public ReadOnly Property Hearing As HearingRecord
        Get
            Return New HearingRecord With {
                .No = noTextBox.Text.Trim(),
                .NameOfPdl = nameTextBox.Text.Trim(),
                .BrCourt = courtTextBox.Text.Trim(),
                .Hearing1 = hearing1TextBox.Text.Trim(),
                .Hearing2 = hearing2TextBox.Text.Trim(),
                .NextHearing = nextHearingPicker.Value.Date
            }
        End Get
    End Property

    Public Sub LoadHearing(h As HearingRecord)
        _editingId = h.Id
        noTextBox.Text = h.No
        nameTextBox.Text = h.NameOfPdl
        courtTextBox.Text = h.BrCourt
        hearing1TextBox.Text = h.Hearing1
        hearing2TextBox.Text = h.Hearing2
        If h.NextHearing <> Date.MinValue Then
            nextHearingPicker.Value = h.NextHearing.Date
        End If
        Text = "Edit Hearing Information"
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If DialogResult = DialogResult.OK AndAlso String.IsNullOrWhiteSpace(nameTextBox.Text) Then
            MessageBox.Show(Me, "Please enter the NAME OF PDL.", "Required Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Cancel = True
            Return
        End If
        MyBase.OnFormClosing(e)
    End Sub

    Private Shared Sub ConfigureInput(textBox As TextBox)
        textBox.BorderStyle = BorderStyle.FixedSingle
        textBox.Width = 300
    End Sub

    Private Shared Sub AddRow(table As TableLayoutPanel, row As Integer, labelText As String, input As Control)
        table.RowStyles.Add(New RowStyle(SizeType.Absolute, 38))
        Dim label As New Label With {
            .Text = labelText,
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft,
            .ForeColor = Color.FromArgb(91, 33, 182)
        }
        input.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        table.Controls.Add(label, 0, row)
        table.Controls.Add(input, 1, row)
    End Sub
End Class
