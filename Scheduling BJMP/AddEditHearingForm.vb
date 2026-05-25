Imports System.Drawing
Imports System.Windows.Forms

Public Class AddEditHearingForm
    Inherits Form

    Private ReadOnly noTextBox As New TextBox()
    Private ReadOnly nameTextBox As New TextBox()
    Private ReadOnly courtTextBox As New TextBox()
    Private ReadOnly hearing1TextBox As New TextBox()
    Private ReadOnly hearing2TextBox As New TextBox()
    Private ReadOnly nextHearingPicker As New DateTimePicker()

    Public Sub New(selectedDate As Date)
        Text = "Add Hearing Information"
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ClientSize = New Size(590, 420)
        Font = New Font("Segoe UI", 10.0F)
        BackColor = Color.White

        Dim titleLabel As New Label With {
            .Text = "Hearing Information",
            .Font = New Font("Segoe UI Semibold", 15.0F),
            .ForeColor = Color.FromArgb(18, 54, 93),
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

        Dim saveButton As New Button With {
            .Text = "Save",
            .BackColor = Color.FromArgb(18, 54, 93),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Location = New Point(376, 350),
            .Size = New Size(90, 36),
            .DialogResult = DialogResult.OK
        }
        saveButton.FlatAppearance.BorderSize = 0

        Dim cancelButton As New Button With {
            .Text = "Cancel",
            .BackColor = Color.FromArgb(235, 239, 244),
            .ForeColor = Color.FromArgb(38, 50, 66),
            .FlatStyle = FlatStyle.Flat,
            .Location = New Point(476, 350),
            .Size = New Size(90, 36),
            .DialogResult = DialogResult.Cancel
        }
        cancelButton.FlatAppearance.BorderSize = 0

        AcceptButton = saveButton
        CancelButton = cancelButton
        Controls.AddRange(New Control() {titleLabel, table, saveButton, cancelButton})
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
            .ForeColor = Color.FromArgb(70, 80, 94)
        }
        input.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        table.Controls.Add(label, 0, row)
        table.Controls.Add(input, 1, row)
    End Sub
End Class
