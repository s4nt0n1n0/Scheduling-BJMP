Imports ClosedXML.Excel
Imports System.Globalization
Imports System.IO
Imports System.Xml.Linq

Public Class ExcelHearingRepository
    Private Const SheetName As String = "Hearings"
    Private ReadOnly _workbookPath As String
    Private ReadOnly _backupDirectory As String

    Private Class HeaderLocation
        Public Property Row As Integer
        Public Property Column As Integer
    End Class

    Private Class ColumnMapping
        Public Property HeaderRow As Integer = 1
        Public Property NoCol As Integer = 1
        Public Property NameCol As Integer = 2
        Public Property CourtCol As Integer = 3
        Public Property Hearing1Col As Integer = 4
        Public Property Hearing2Col As Integer = 5
        Public Property NextHearingCol As Integer = 6
    End Class

    Public Sub New()
        Dim dataDirectory = Path.Combine(AppContext.BaseDirectory, "Data")
        Directory.CreateDirectory(dataDirectory)
        _workbookPath = Path.Combine(dataDirectory, "hearings.xlsx")
        _backupDirectory = Path.Combine(dataDirectory, "Backups")
        Directory.CreateDirectory(_backupDirectory)
        EnsureWorkbook()
    End Sub

    Public ReadOnly Property WorkbookPath As String
        Get
            Return _workbookPath
        End Get
    End Property

    Public ReadOnly Property BackupDirectory As String
        Get
            Return _backupDirectory
        End Get
    End Property

    ''' <summary>
    ''' Returns all backup files sorted newest-first.
    ''' </summary>
    Public Function ListBackups() As String()
        If Not Directory.Exists(_backupDirectory) Then
            Return Array.Empty(Of String)()
        End If
        Return Directory.GetFiles(_backupDirectory, "hearings_backup_*.xlsx").
            OrderByDescending(Function(f) f).
            ToArray()
    End Function

    Public Function LoadHearings() As List(Of HearingRecord)
        EnsureWorkbook()

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            Dim mapping = FindColumnMapping(worksheet)
            Dim lastRow = LastUsedRowNumber(worksheet)
            Dim hearings As New List(Of HearingRecord)()

            For row = mapping.HeaderRow + 1 To lastRow
                If worksheet.Row(row).IsEmpty() Then
                    Continue For
                End If

                Dim hearingDate As Date
                If Not TryReadDate(worksheet.Cell(row, mapping.NextHearingCol), hearingDate) Then
                    hearingDate = Date.MinValue
                End If

                hearings.Add(New HearingRecord With {
                    .Id = row,
                    .No = worksheet.Cell(row, mapping.NoCol).GetString(),
                    .NameOfPdl = worksheet.Cell(row, mapping.NameCol).GetString(),
                    .BrCourt = worksheet.Cell(row, mapping.CourtCol).GetString(),
                    .Hearing1 = worksheet.Cell(row, mapping.Hearing1Col).GetString(),
                    .Hearing2 = worksheet.Cell(row, mapping.Hearing2Col).GetString(),
                    .NextHearing = hearingDate.Date
                })
            Next

            Return hearings
        End Using
    End Function

    Public Function AddHearing(hearing As HearingRecord) As HearingRecord
        EnsureWorkbook()

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            Dim mapping = FindColumnMapping(worksheet)
            Dim nextRow = LastUsedRowNumber(worksheet) + 1
            hearing.Id = nextRow
            WriteHearing(worksheet, nextRow, hearing, mapping)
            workbook.Save()
        End Using

        Return hearing
    End Function

    Public Sub UpdateHearing(hearing As HearingRecord)
        EnsureWorkbook()

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            If hearing.Id < 2 OrElse hearing.Id > LastUsedRowNumber(worksheet) Then
                Throw New InvalidOperationException("The selected hearing no longer exists in the Excel file.")
            End If

            Dim mapping = FindColumnMapping(worksheet)
            WriteHearing(worksheet, hearing.Id, hearing, mapping)
            workbook.Save()
        End Using
    End Sub

    Public Sub MoveHearing(rowId As Integer, nextHearing As Date)
        EnsureWorkbook()

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            If rowId < 2 OrElse rowId > LastUsedRowNumber(worksheet) Then
                Throw New InvalidOperationException("The moved hearing no longer exists in the Excel file.")
            End If

            Dim mapping = FindColumnMapping(worksheet)
            worksheet.Cell(rowId, mapping.NextHearingCol).Value = nextHearing.Date
            worksheet.Cell(rowId, mapping.NextHearingCol).Style.DateFormat.Format = "yyyy-mm-dd"
            workbook.Save()
        End Using
    End Sub

    Public Sub DeleteHearing(rowId As Integer)
        EnsureWorkbook()

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            If rowId < 2 OrElse rowId > LastUsedRowNumber(worksheet) Then
                Throw New InvalidOperationException("The selected hearing no longer exists in the Excel file.")
            End If

            worksheet.Row(rowId).Delete()
            workbook.Save()
        End Using
    End Sub


    ''' <summary>
    ''' Automatically backs up the current workbook to the internal Backups folder with a timestamped filename.
    ''' </summary>
    Public Sub BackupCurrentData()
        Dim filename = "hearings_backup_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
        Dim destinationPath = Path.Combine(_backupDirectory, filename)
        BackupCurrentData(destinationPath)
    End Sub

    ''' <summary>
    ''' Copies the current hearings.xlsx to the destination path chosen by the user.
    ''' Safe to call even if the workbook does not yet exist.
    ''' </summary>
    ''' <param name="destinationPath">Full file path where the backup should be saved.</param>
    Public Sub BackupCurrentData(destinationPath As String)
        If Not File.Exists(_workbookPath) Then
            Return  ' Nothing to back up
        End If

        ' Make sure the destination folder exists
        Dim destDir = Path.GetDirectoryName(destinationPath)
        If Not String.IsNullOrEmpty(destDir) Then
            Directory.CreateDirectory(destDir)
        End If

        File.Copy(_workbookPath, destinationPath, overwrite:=True)
    End Sub

    Public Function CountSchedulableHearings() As Integer
        Return LoadHearings().Count
    End Function

    Private Sub EnsureWorkbook()
        If File.Exists(_workbookPath) Then
            Return
        End If

        Using workbook As New XLWorkbook()
            Dim worksheet = workbook.Worksheets.Add(SheetName)
            worksheet.Cell(1, 1).Value = "NO"
            worksheet.Cell(1, 2).Value = "NAME OF PDL"
            worksheet.Cell(1, 3).Value = "BR/COURT"
            worksheet.Cell(1, 4).Value = "HEARING"
            worksheet.Cell(1, 5).Value = "HEARING"
            worksheet.Cell(1, 6).Value = "NEXT HEARING"
            worksheet.Range(1, 1, 1, 6).Style.Font.Bold = True
            worksheet.Range(1, 1, 1, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#12365d")
            worksheet.Range(1, 1, 1, 6).Style.Font.FontColor = XLColor.White
            worksheet.Columns().AdjustToContents()
            workbook.SaveAs(_workbookPath)
        End Using
    End Sub

    Public Sub ImportDataFile(sourcePath As String)
        ' ── Auto-backup current data before modifying it ──────────────────────
        BackupCurrentData()

        Dim extension = Path.GetExtension(sourcePath).ToLowerInvariant()
        Dim importedList As List(Of HearingRecord) = Nothing

        Select Case extension
            Case ".xlsx", ".xlsm"
                importedList = ParseWorkbook(sourcePath)
            Case ".xml"
                importedList = ParseSpreadsheetXml(sourcePath)
            Case Else
                Throw New InvalidOperationException("Please import an Excel .xlsx/.xlsm file or an Excel XML .xml file.")
        End Select

        If importedList Is Nothing OrElse importedList.Count = 0 Then
            Return
        End If

        ' Load existing active hearings to merge against
        Dim existingList = LoadHearings()
        Dim mergedCount = 0

        Using workbook As New XLWorkbook(_workbookPath)
            Dim worksheet = workbook.Worksheet(SheetName)
            Dim mapping = FindColumnMapping(worksheet)
            Dim currentLastRow = LastUsedRowNumber(worksheet)

            For Each newH In importedList
                ' Check if duplicate: same Case Number AND Name of PDL
                Dim duplicate = existingList.FirstOrDefault(Function(x) 
                    Return x.No.Trim().Equals(newH.No.Trim(), StringComparison.OrdinalIgnoreCase) AndAlso
                           x.NameOfPdl.Trim().Equals(newH.NameOfPdl.Trim(), StringComparison.OrdinalIgnoreCase)
                End Function)

                If duplicate IsNot Nothing Then
                    ' DUPLICATE: Update existing hearing status/results in-place if they were blank, otherwise skip
                    Dim rowId = duplicate.Id
                    If String.IsNullOrWhiteSpace(worksheet.Cell(rowId, mapping.Hearing1Col).GetString()) Then
                        worksheet.Cell(rowId, mapping.Hearing1Col).Value = newH.Hearing1
                    End If
                    If String.IsNullOrWhiteSpace(worksheet.Cell(rowId, mapping.Hearing2Col).GetString()) Then
                        worksheet.Cell(rowId, mapping.Hearing2Col).Value = newH.Hearing2
                    End If
                    If newH.NextHearing <> Date.MinValue AndAlso duplicate.NextHearing = Date.MinValue Then
                        worksheet.Cell(rowId, mapping.NextHearingCol).Value = newH.NextHearing.Date
                        worksheet.Cell(rowId, mapping.NextHearingCol).Style.DateFormat.Format = "yyyy-mm-dd"
                    End If
                Else
                    ' NOT DUPLICATE: Append new row to existing Excel sheet
                    currentLastRow += 1
                    newH.Id = currentLastRow
                    WriteHearing(worksheet, currentLastRow, newH, mapping)
                    existingList.Add(newH)
                    mergedCount += 1
                End If
            Next

            workbook.Save()
        End Using
    End Sub

    Private Function ParseWorkbook(path As String) As List(Of HearingRecord)
        Dim parsedList As New List(Of HearingRecord)()

        Using workbook As New XLWorkbook(path)
            For Each worksheet In workbook.Worksheets
                Dim mapping = FindColumnMapping(worksheet)
                If mapping.HeaderRow > 0 Then
                    Dim lastRow = LastUsedRowNumber(worksheet)
                    Dim currentDate As Date = Date.MinValue

                    For rowNumber = mapping.HeaderRow + 1 To lastRow
                        If worksheet.Row(rowNumber).IsEmpty() Then
                            Continue For
                        End If

                        Dim noVal = worksheet.Cell(rowNumber, mapping.NoCol).GetString().Trim()
                        Dim nameVal = worksheet.Cell(rowNumber, mapping.NameCol).GetString().Trim()
                        Dim courtVal = worksheet.Cell(rowNumber, mapping.CourtCol).GetString().Trim()
                        Dim hearing1Val = worksheet.Cell(rowNumber, mapping.Hearing1Col).GetString().Trim()
                        Dim hearing2Val = worksheet.Cell(rowNumber, mapping.Hearing2Col).GetString().Trim()
                        Dim nextVal = worksheet.Cell(rowNumber, mapping.NextHearingCol).GetString().Trim()

                        ' Check if it's a section date divider row
                        Dim col1To5Empty = String.IsNullOrWhiteSpace(nameVal) AndAlso
                                           String.IsNullOrWhiteSpace(courtVal) AndAlso
                                           String.IsNullOrWhiteSpace(hearing1Val) AndAlso
                                           String.IsNullOrWhiteSpace(hearing2Val) AndAlso
                                           String.IsNullOrWhiteSpace(nextVal)

                        Dim sectionDate As Date
                        If col1To5Empty AndAlso TryParseSectionDate(noVal, sectionDate) Then
                            currentDate = sectionDate
                            Continue For
                        End If

                        ' Skip helper/header-like lines that are not actual hearing rows
                        Dim hearingDate As Date = Date.MinValue
                        If Not TryReadDate(worksheet.Cell(rowNumber, mapping.NextHearingCol), hearingDate) Then
                            hearingDate = currentDate
                        End If

                        parsedList.Add(New HearingRecord With {
                            .No = noVal,
                            .NameOfPdl = nameVal,
                            .BrCourt = courtVal,
                            .Hearing1 = hearing1Val,
                            .Hearing2 = hearing2Val,
                            .NextHearing = hearingDate.Date
                        })
                    Next
                End If
            Next
        End Using

        Return parsedList
    End Function

    Private Function ParseSpreadsheetXml(sourcePath As String) As List(Of HearingRecord)
        Dim parsedList As New List(Of HearingRecord)()
        Dim document = XDocument.Load(sourcePath)
        Dim rows = document.Descendants().Where(Function(element) element.Name.LocalName = "Row").ToList()
        If rows.Count = 0 Then
            Throw New InvalidOperationException("The XML file does not contain Excel rows.")
        End If

        Dim importedRows = rows.Select(Function(row) ReadXmlRow(row)).ToList()
        Dim headerRowIndex = -1
        Dim mapping As New ColumnMapping()

        ' Scan XML rows to find a header row
        For rowIndex = 0 To importedRows.Count - 1
            Dim row = importedRows(rowIndex)
            Dim tempNoCol As Integer = -1
            Dim tempNameCol As Integer = -1
            Dim tempCourtCol As Integer = -1
            Dim tempStatusCol As Integer = -1
            Dim tempResultCol As Integer = -1
            Dim tempNextCol As Integer = -1
            Dim matchCount As Integer = 0

            For colIndex = 0 To row.Count - 1
                Dim cellValue = NormalizeHeader(row(colIndex))
                If String.IsNullOrEmpty(cellValue) Then Continue For

                If cellValue = "NO" Then
                    tempNoCol = colIndex
                    matchCount += 1
                ElseIf cellValue = "NAME OF PDL" OrElse cellValue = "NAME" OrElse cellValue.Contains("PDL") Then
                    tempNameCol = colIndex
                    matchCount += 1
                ElseIf cellValue.StartsWith("BR") OrElse cellValue.Contains("COURT") OrElse cellValue = "BRANCH" Then
                    tempCourtCol = colIndex
                    matchCount += 1
                ElseIf cellValue = "NEXT HEARING" OrElse cellValue.Contains("NEXT") Then
                    tempNextCol = colIndex
                    matchCount += 1
                ElseIf cellValue = "HEARING" OrElse cellValue.Contains("HEARING") Then
                    If tempStatusCol < 0 Then
                        tempStatusCol = colIndex
                    Else
                        tempResultCol = colIndex
                    End If
                    matchCount += 1
                End If
            Next

            If matchCount >= 4 Then
                headerRowIndex = rowIndex
                mapping.NoCol = If(tempNoCol >= 0, tempNoCol, 0)
                mapping.NameCol = If(tempNameCol >= 0, tempNameCol, 1)
                mapping.CourtCol = If(tempCourtCol >= 0, tempCourtCol, 2)
                mapping.Hearing1Col = If(tempStatusCol >= 0, tempStatusCol, 3)
                mapping.Hearing2Col = If(tempResultCol >= 0, tempResultCol, 4)
                mapping.NextHearingCol = If(tempNextCol >= 0, tempNextCol, 5)
                Exit For
            End If
        Next

        If headerRowIndex < 0 Then
            ThrowHeaderException()
        End If

        Dim currentDate As Date = Date.MinValue
        For rowIndex = headerRowIndex + 1 To importedRows.Count - 1
            Dim row = importedRows(rowIndex)
            Dim noVal = If(mapping.NoCol < row.Count, row(mapping.NoCol).Trim(), "")
            Dim nameVal = If(mapping.NameCol < row.Count, row(mapping.NameCol).Trim(), "")
            Dim courtVal = If(mapping.CourtCol < row.Count, row(mapping.CourtCol).Trim(), "")
            Dim hearing1Val = If(mapping.Hearing1Col < row.Count, row(mapping.Hearing1Col).Trim(), "")
            Dim hearing2Val = If(mapping.Hearing2Col < row.Count, row(mapping.Hearing2Col).Trim(), "")
            Dim nextVal = If(mapping.NextHearingCol < row.Count, row(mapping.NextHearingCol).Trim(), "")

            ' Check if it's a section date divider row
            Dim col1To5Empty = String.IsNullOrWhiteSpace(nameVal) AndAlso
                               String.IsNullOrWhiteSpace(courtVal) AndAlso
                               String.IsNullOrWhiteSpace(hearing1Val) AndAlso
                               String.IsNullOrWhiteSpace(hearing2Val) AndAlso
                               String.IsNullOrWhiteSpace(nextVal)

            Dim sectionDate As Date
            If col1To5Empty AndAlso TryParseSectionDate(noVal, sectionDate) Then
                currentDate = sectionDate
                Continue For
            End If

            ' Skip helper/header-like lines that are not actual hearing rows
            If String.IsNullOrWhiteSpace(nameVal) Then
                Continue For
            End If

            Dim hearingDate As Date = Date.MinValue
            ' Parse exact date
            If Not Date.TryParse(nextVal, CultureInfo.CurrentCulture, DateTimeStyles.None, hearingDate) AndAlso
               Not Date.TryParse(nextVal, CultureInfo.InvariantCulture, DateTimeStyles.None, hearingDate) Then
                hearingDate = currentDate
            End If

            parsedList.Add(New HearingRecord With {
                .No = noVal,
                .NameOfPdl = nameVal,
                .BrCourt = courtVal,
                .Hearing1 = hearing1Val,
                .Hearing2 = hearing2Val,
                .NextHearing = hearingDate.Date
            })
        Next

        Return parsedList
    End Function

    Private Shared Function ReadXmlRow(row As XElement) As List(Of String)
        Dim values As New List(Of String)()
        Dim columnIndex = 1

        For Each cell In row.Elements().Where(Function(element) element.Name.LocalName = "Cell")
            Dim indexAttribute = cell.Attributes().FirstOrDefault(Function(attribute) attribute.Name.LocalName = "Index")
            If indexAttribute IsNot Nothing Then
                Dim explicitIndex As Integer
                If Integer.TryParse(indexAttribute.Value, explicitIndex) Then
                    columnIndex = explicitIndex
                End If
            End If

            While values.Count < columnIndex - 1
                values.Add("")
            End While

            values.Add(cell.Descendants().FirstOrDefault(Function(element) element.Name.LocalName = "Data")?.Value)
            columnIndex += 1
        Next

        While values.Count < 6
            values.Add("")
        End While

        Return values
    End Function

    Private Shared Function FindColumnMapping(worksheet As IXLWorksheet) As ColumnMapping
        Dim mapping As New ColumnMapping()
        Dim lastRow = LastUsedRowNumber(worksheet)

        ' Scan rows to find the best header row candidate
        For row = 1 To Math.Min(100, lastRow)
            Dim lastCell = worksheet.Row(row).LastCellUsed()
            Dim lastCol = If(lastCell Is Nothing, 20, Math.Max(20, lastCell.Address.ColumnNumber))

            Dim tempNoCol As Integer = 0
            Dim tempNameCol As Integer = 0
            Dim tempCourtCol As Integer = 0
            Dim tempHearing1Col As Integer = 0
            Dim tempHearing2Col As Integer = 0
            Dim tempNextCol As Integer = 0
            Dim matchCount As Integer = 0

            For col = 1 To lastCol
                Dim cellValue = NormalizeHeader(worksheet.Cell(row, col).GetString())
                If String.IsNullOrEmpty(cellValue) Then Continue For

                If cellValue = "NO" Then
                    tempNoCol = col
                    matchCount += 1
                ElseIf cellValue = "NAME OF PDL" OrElse cellValue = "NAME" OrElse cellValue.Contains("PDL") Then
                    tempNameCol = col
                    matchCount += 1
                ElseIf cellValue.StartsWith("BR") OrElse cellValue.Contains("COURT") OrElse cellValue = "BRANCH" Then
                    tempCourtCol = col
                    matchCount += 1
                ElseIf cellValue = "NEXT HEARING" OrElse cellValue.Contains("NEXT") Then
                    tempNextCol = col
                    matchCount += 1
                ElseIf cellValue = "HEARING" OrElse cellValue.Contains("HEARING") Then
                    If tempHearing1Col = 0 Then
                        tempHearing1Col = col
                    Else
                        tempHearing2Col = col
                    End If
                    matchCount += 1
                End If
            Next

            ' If we matched at least 4 of the key headers, we consider this the header row
            If matchCount >= 4 Then
                mapping.HeaderRow = row
                mapping.NoCol = If(tempNoCol > 0, tempNoCol, 1)
                mapping.NameCol = If(tempNameCol > 0, tempNameCol, 2)
                mapping.CourtCol = If(tempCourtCol > 0, tempCourtCol, 3)
                mapping.Hearing1Col = If(tempHearing1Col > 0, tempHearing1Col, 4)
                mapping.Hearing2Col = If(tempHearing2Col > 0, tempHearing2Col, 5)
                mapping.NextHearingCol = If(tempNextCol > 0, tempNextCol, 6)
                Return mapping
            End If
        Next

        Return mapping
    End Function

    Private Sub SaveNormalizedWorkbook(rows As List(Of List(Of String)))
        If rows.Count = 0 OrElse Not HeadersMatch(rows(0).Take(6)) Then
            ThrowHeaderException()
        End If

        Dim normalizedRows As New List(Of List(Of String)) From {StandardHeaders().ToList()}
        Dim currentDate As Date = Date.MinValue

        For Each row In rows.Skip(1)
            Dim normalizedRow = Enumerable.Range(0, 6).
                Select(Function(index) If(index < row.Count AndAlso row(index) IsNot Nothing, row(index).Trim(), "")).
                ToList()

            If normalizedRow.All(Function(value) String.IsNullOrWhiteSpace(value)) OrElse HeadersMatch(normalizedRow) Then
                Continue For
            End If

            ' Check if it's a section date divider row (column 0 has a date, other columns are empty)
            Dim col0 = normalizedRow(0)
            Dim col1To5Empty = normalizedRow.Skip(1).All(Function(val) String.IsNullOrWhiteSpace(val))
            Dim sectionDate As Date

            If col1To5Empty AndAlso TryParseSectionDate(col0, sectionDate) Then
                currentDate = sectionDate
                Continue For ' Skip adding the section divider row itself to the database
            End If

            ' Skip helper/header-like lines that are not actual hearing rows
            If String.IsNullOrWhiteSpace(normalizedRow(1)) Then
                Continue For
            End If

            ' If the row has no Next Hearing date, assign the current active section date
            If String.IsNullOrWhiteSpace(normalizedRow(5)) AndAlso currentDate <> Date.MinValue Then
                normalizedRow(5) = currentDate.ToString("yyyy-MM-dd")
            End If

            normalizedRows.Add(normalizedRow)
        Next

        Using workbook As New XLWorkbook()
            Dim worksheet = workbook.Worksheets.Add(SheetName)
            For rowIndex = 0 To normalizedRows.Count - 1
                For columnIndex = 0 To Math.Min(5, normalizedRows(rowIndex).Count - 1)
                    worksheet.Cell(rowIndex + 1, columnIndex + 1).Value = normalizedRows(rowIndex)(columnIndex)
                Next
            Next

            worksheet.Range(1, 1, 1, 6).Style.Font.Bold = True
            worksheet.Range(1, 1, 1, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#12365d")
            worksheet.Range(1, 1, 1, 6).Style.Font.FontColor = XLColor.White
            worksheet.Columns().AdjustToContents()
            workbook.SaveAs(_workbookPath)
        End Using
    End Sub

    Private Shared Function TryParseSectionDate(text As String, ByRef value As Date) As Boolean
        If String.IsNullOrWhiteSpace(text) Then
            Return False
        End If

        Dim cleanText = text.Trim()

        ' Remove day of the week in parentheses if present, e.g., "APRIL 7, 2026 (TUESDAY)" -> "APRIL 7, 2026"
        Dim parenIndex = cleanText.IndexOf("(")
        If parenIndex >= 0 Then
            cleanText = cleanText.Substring(0, parenIndex).Trim()
        End If

        ' Remove other common day of week suffixes
        Dim daysOfWeek As String() = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"}
        For Each day In daysOfWeek
            If cleanText.ToUpperInvariant().EndsWith(day) Then
                cleanText = cleanText.Substring(0, cleanText.Length - day.Length).Trim().Trim(","c).Trim()
                Exit For
            End If
        Next

        ' Try standard parsing
        If Date.TryParse(cleanText, CultureInfo.CurrentCulture, DateTimeStyles.None, value) OrElse
           Date.TryParse(cleanText, CultureInfo.InvariantCulture, DateTimeStyles.None, value) Then
            Return True
        End If

        ' Try with custom format list
        Dim formats As String() = {
            "MMMM d, yyyy", "MMMM dd, yyyy", "MMM d, yyyy", "MMM dd, yyyy",
            "d MMMM yyyy", "dd MMMM yyyy", "d MMM yyyy", "dd MMM yyyy"
        }
        Return Date.TryParseExact(cleanText, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, value) OrElse
               Date.TryParseExact(cleanText, formats, CultureInfo.CurrentCulture, DateTimeStyles.None, value)
    End Function

    Private Shared Function HeadersMatch(headers As IEnumerable(Of String)) As Boolean
        Dim actualHeaders = headers.Select(Function(header) NormalizeHeader(header)).ToArray()

        Return actualHeaders.Length >= 6 AndAlso
            actualHeaders(0) = "NO" AndAlso
            actualHeaders(1) = "NAME OF PDL" AndAlso
            actualHeaders(2).StartsWith("BR") AndAlso
            actualHeaders(3) = "HEARING" AndAlso
            actualHeaders(4) = "HEARING" AndAlso
            actualHeaders(5) = "NEXT HEARING"
    End Function

    Private Shared Function StandardHeaders() As String()
        Return {"NO", "NAME OF PDL", "BR/COURT", "HEARING", "HEARING", "NEXT HEARING"}
    End Function

    Private Shared Function NormalizeHeader(header As String) As String
        If header Is Nothing Then
            Return ""
        End If

        Return String.Join(" ", header.Split(CType(Nothing, Char()), StringSplitOptions.RemoveEmptyEntries)).
            Trim().
            ToUpperInvariant()
    End Function

    Private Shared Sub ThrowHeaderException()
        Throw New InvalidOperationException("The imported file must contain a six-column hearing table in this order: NO, NAME OF PDL, BR/COURT, HEARING STATUS, HEARING RESULT, NEXT HEARING. The table may start in any column, and title rows above it are allowed.")
    End Sub

    Private Shared Sub WriteHearing(worksheet As IXLWorksheet, row As Integer, hearing As HearingRecord, mapping As ColumnMapping)
        worksheet.Cell(row, mapping.NoCol).Value = hearing.No
        worksheet.Cell(row, mapping.NameCol).Value = hearing.NameOfPdl
        worksheet.Cell(row, mapping.CourtCol).Value = hearing.BrCourt
        worksheet.Cell(row, mapping.Hearing1Col).Value = hearing.Hearing1
        worksheet.Cell(row, mapping.Hearing2Col).Value = hearing.Hearing2
        If hearing.NextHearing <> Date.MinValue Then
            worksheet.Cell(row, mapping.NextHearingCol).Value = hearing.NextHearing.Date
            worksheet.Cell(row, mapping.NextHearingCol).Style.DateFormat.Format = "yyyy-mm-dd"
        Else
            worksheet.Cell(row, mapping.NextHearingCol).Value = ""
        End If
        worksheet.Columns().AdjustToContents()
    End Sub

    Private Shared Function LastUsedRowNumber(worksheet As IXLWorksheet) As Integer
        Dim lastRow = worksheet.LastRowUsed()
        If lastRow Is Nothing Then
            Return 1
        End If

        Return lastRow.RowNumber()
    End Function

    Private Shared Function TryReadDate(cell As IXLCell, ByRef value As Date) As Boolean
        If cell.DataType = XLDataType.DateTime Then
            Try
                value = cell.GetDateTime()
                Return True
            Catch
            End Try
        End If

        If cell.DataType = XLDataType.Number Then
            Try
                value = Date.FromOADate(cell.GetDouble())
                Return True
            Catch
            End Try
        End If

        Dim textValue = cell.GetString().Trim()
        If String.IsNullOrWhiteSpace(textValue) Then
            Return False
        End If

        Dim numericValue As Double
        If Double.TryParse(textValue, NumberStyles.Number, CultureInfo.InvariantCulture, numericValue) AndAlso numericValue > 20000 Then
            Try
                value = Date.FromOADate(numericValue)
                Return True
            Catch
            End Try
        End If

        Dim commonFormats As String() = {
            "yyyy-MM-dd", "MM/dd/yyyy", "M/d/yyyy", "dd/MM/yyyy", "d/M/yyyy",
            "MMMM dd, yyyy", "MMMM d, yyyy", "MMM dd, yyyy", "MMM d, yyyy",
            "d-MMM-yy", "dd-MMM-yy", "d-MMM-yyyy", "dd-MMM-yyyy",
            "yyyy/MM/dd", "yyyy.MM.dd", "dd-MM-yyyy"
        }

        Return Date.TryParse(textValue, CultureInfo.CurrentCulture, DateTimeStyles.None, value) OrElse
            Date.TryParse(textValue, CultureInfo.InvariantCulture, DateTimeStyles.None, value) OrElse
            Date.TryParseExact(textValue, commonFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, value) OrElse
            Date.TryParseExact(textValue, commonFormats, CultureInfo.CurrentCulture, DateTimeStyles.None, value)
    End Function

End Class
