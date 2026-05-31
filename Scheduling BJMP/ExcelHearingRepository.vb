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
        _workbookPath = Path.Combine(dataDirectory, "hearings_saved.xml")
        _backupDirectory = Path.Combine(dataDirectory, "Backups")
        Directory.CreateDirectory(_backupDirectory)
        EnsureMasterFile()
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
        Return Directory.GetFiles(_backupDirectory, "hearings_backup_*.xml").
            OrderByDescending(Function(f) f).
            ToArray()
    End Function

    Private Sub EnsureMasterFile()
        If File.Exists(_workbookPath) Then
            Return
        End If

        Try
            Dim doc As New XDocument(
                New XDeclaration("1.0", "utf-8", "yes"),
                New XElement("Hearings")
            )
            doc.Save(_workbookPath)
        Catch ex As Exception
            Throw New IOException($"Unable to create the hearing data file at '{_workbookPath}'.", ex)
        End Try
    End Sub

    Private Sub SaveHearings(hearings As List(Of HearingRecord))
        Dim doc As New XDocument(
            New XDeclaration("1.0", "utf-8", "yes"),
            New XElement("Hearings",
                hearings.Select(Function(h)
                    Dim historyEl = New XElement("HistoryLog")
                    If h.HistoryLog IsNot Nothing Then
                        For Each log In h.HistoryLog
                            historyEl.Add(New XElement("Log", log))
                        Next
                    End If

                    Return New XElement("Hearing",
                        New XElement("Id", h.Id),
                        New XElement("CaseNo", h.No),
                        New XElement("NameOfPdl", h.NameOfPdl),
                        New XElement("BrCourt", h.BrCourt),
                        New XElement("Hearing1", h.Hearing1),
                        New XElement("Hearing2", h.Hearing2),
                        New XElement("NextHearing", If(h.NextHearing = Date.MinValue, "", h.NextHearing.ToString("yyyy-MM-dd"))),
                        historyEl
                    )
                End Function)
            )
        )

        Dim retries As Integer = 5
        Dim lastError As IOException = Nothing
        While retries > 0
            Try
                doc.Save(_workbookPath)
                Exit Sub
            Catch ex As IOException
                lastError = ex
                retries -= 1
                If retries <= 0 Then Exit While
                System.Threading.Thread.Sleep(100)
            End Try
        End While

        If lastError IsNot Nothing Then
            Throw New IOException($"Unable to save hearing data to '{_workbookPath}'.", lastError)
        End If
    End Sub

    Public Function LoadHearings() As List(Of HearingRecord)
        EnsureMasterFile()
        Dim hearings As New List(Of HearingRecord)()

        Try
            Dim doc = XDocument.Load(_workbookPath)
            Dim elements = doc.Descendants("Hearing")
            For Each el In elements
                Dim record As New HearingRecord With {
                    .Id = If(el.Element("Id") IsNot Nothing, Integer.Parse(el.Element("Id").Value), 0),
                    .No = If(el.Element("CaseNo") IsNot Nothing, el.Element("CaseNo").Value, If(el.Element("No") IsNot Nothing, el.Element("No").Value, "")),
                    .NameOfPdl = If(el.Element("NameOfPdl") IsNot Nothing, el.Element("NameOfPdl").Value, ""),
                    .BrCourt = If(el.Element("BrCourt") IsNot Nothing, el.Element("BrCourt").Value, ""),
                    .Hearing1 = If(el.Element("Hearing1") IsNot Nothing, el.Element("Hearing1").Value, ""),
                    .Hearing2 = If(el.Element("Hearing2") IsNot Nothing, el.Element("Hearing2").Value, ""),
                    .NextHearing = Date.MinValue
                }

                Dim nextHearingStr = If(el.Element("NextHearing") IsNot Nothing, el.Element("NextHearing").Value, "")
                Dim nextHearingDate As Date
                If Date.TryParse(nextHearingStr, CultureInfo.InvariantCulture, DateTimeStyles.None, nextHearingDate) Then
                    record.NextHearing = nextHearingDate.Date
                End If

                Dim headerNames As String() = {"NAME OF PDL", "NAME", "NAMES OF PDL"}
                Dim headerNos As String() = {"NO", "NO."}
                If headerNames.Any(Function(h) record.NameOfPdl.Equals(h, StringComparison.OrdinalIgnoreCase)) OrElse
                   headerNos.Any(Function(h) record.No.Equals(h, StringComparison.OrdinalIgnoreCase)) Then
                    Continue For
                End If

                Dim historyEl = el.Element("HistoryLog")
                If historyEl IsNot Nothing Then
                    For Each entryEl In historyEl.Elements("Log")
                        record.HistoryLog.Add(entryEl.Value)
                    Next
                End If

                hearings.Add(record)
            Next
        Catch ex As Exception
            Throw New InvalidOperationException($"Failed to load hearing data from '{_workbookPath}'.", ex)
        End Try

        Return hearings
    End Function

    Public Function AddHearing(hearing As HearingRecord) As HearingRecord
        Dim hearings = LoadHearings()
        Dim nextId = If(hearings.Count > 0, hearings.Max(Function(h) h.Id) + 1, 1)
        hearing.Id = nextId
        hearings.Add(hearing)
        SaveHearings(hearings)
        Return hearing
    End Function

    Public Sub UpdateHearing(hearing As HearingRecord)
        Dim hearings = LoadHearings()
        Dim existing = hearings.FirstOrDefault(Function(h) h.Id = hearing.Id)
        If existing Is Nothing Then
            Throw New InvalidOperationException("The selected hearing no longer exists.")
        End If
        existing.No = hearing.No
        existing.NameOfPdl = hearing.NameOfPdl
        existing.BrCourt = hearing.BrCourt
        existing.Hearing1 = hearing.Hearing1
        existing.Hearing2 = hearing.Hearing2
        existing.NextHearing = hearing.NextHearing
        existing.HistoryLog = hearing.HistoryLog
        SaveHearings(hearings)
    End Sub

    Public Function RenameHearings(oldName As String, newName As String) As Integer
        Dim trimmedOld = If(oldName, "").Trim()
        Dim trimmedNew = If(newName, "").Trim()
        If String.IsNullOrWhiteSpace(trimmedOld) OrElse String.Equals(trimmedOld, trimmedNew, StringComparison.OrdinalIgnoreCase) Then
            Return 0
        End If

        Dim hearings = LoadHearings()
        Dim updatedCount = 0

        For Each hearing In hearings
            If String.Equals(If(hearing.NameOfPdl, "").Trim(), trimmedOld, StringComparison.OrdinalIgnoreCase) Then
                hearing.NameOfPdl = trimmedNew
                updatedCount += 1
            End If
        Next

        If updatedCount > 0 Then
            SaveHearings(hearings)
        End If

        Return updatedCount
    End Function

    Public Sub MoveHearing(rowId As Integer, nextHearing As Date)
        Dim hearings = LoadHearings()
        Dim existing = hearings.FirstOrDefault(Function(h) h.Id = rowId)
        If existing Is Nothing Then
            Throw New InvalidOperationException("The moved hearing no longer exists.")
        End If
        existing.NextHearing = nextHearing.Date
        SaveHearings(hearings)
    End Sub

    Public Sub DeleteHearing(rowId As Integer)
        Dim hearings = LoadHearings()
        Dim existing = hearings.FirstOrDefault(Function(h) h.Id = rowId)
        If existing Is Nothing Then
            Throw New InvalidOperationException("The selected hearing no longer exists.")
        End If
        hearings.Remove(existing)
        SaveHearings(hearings)
    End Sub

    ''' <summary>
    ''' Automatically backs up the current master file to the internal Backups folder with a timestamped filename.
    ''' </summary>
    Public Sub BackupCurrentData()
        Dim filename = "hearings_backup_" & DateTime.Now.ToString("yyyy-MM-dd_HHmmss") & ".xml"
        Dim destinationPath = Path.Combine(_backupDirectory, filename)
        BackupCurrentData(destinationPath)
    End Sub

    ''' <summary>
    ''' Copies the current hearings_saved.xml to the destination path.
    ''' </summary>
    Public Sub BackupCurrentData(destinationPath As String)
        If Not File.Exists(_workbookPath) Then
            Return
        End If

        Try
            Dim destDir = Path.GetDirectoryName(destinationPath)
            If Not String.IsNullOrEmpty(destDir) Then
                Directory.CreateDirectory(destDir)
            End If

            File.Copy(_workbookPath, destinationPath, overwrite:=True)
        Catch ex As Exception
            Throw New IOException($"Failed to create a backup at '{destinationPath}'.", ex)
        End Try
    End Sub

    Public Function CountSchedulableHearings() As Integer
        Return LoadHearings().Count
    End Function

    Public Sub ImportDataFile(sourcePath As String)
        Try
            BackupCurrentData()

            Dim extension = Path.GetExtension(sourcePath).ToLowerInvariant()
            Dim importedList As List(Of HearingRecord) = Nothing

            Select Case extension
                Case ".xlsx", ".xlsm"
                    importedList = ParseWorkbook(sourcePath)
                Case ".xml"
                    importedList = ParseXmlOrSpreadsheetXml(sourcePath)
                Case Else
                    Throw New InvalidOperationException("Please import an Excel .xlsx/.xlsm file or an XML .xml file.")
            End Select

            If importedList Is Nothing OrElse importedList.Count = 0 Then
                Return
            End If

            Dim existingList = LoadHearings()

            For Each newH In importedList
                ' A record is only a duplicate if BOTH CaseNo AND NextHearing date match.
                ' This allows the same person to have multiple hearings on different dates.
                Dim duplicate = existingList.FirstOrDefault(Function(x)
                    Dim sameNo = Not String.IsNullOrWhiteSpace(x.No) AndAlso
                                 Not String.IsNullOrWhiteSpace(newH.No) AndAlso
                                 x.No.Trim().Equals(newH.No.Trim(), StringComparison.OrdinalIgnoreCase)
                    Dim sameDate = x.NextHearing <> Date.MinValue AndAlso
                                   newH.NextHearing <> Date.MinValue AndAlso
                                   x.NextHearing.Date = newH.NextHearing.Date
                    Return sameNo AndAlso sameDate
                End Function)

                If duplicate Is Nothing Then
                    Dim nextId = If(existingList.Count > 0, existingList.Max(Function(h) h.Id) + 1, 1)
                    newH.Id = nextId
                    existingList.Add(newH)
                End If
            Next

            SaveHearings(existingList)
        Catch ex As Exception
            Throw New InvalidOperationException($"Failed to import hearing data from '{sourcePath}'.", ex)
        End Try
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

                        ' Skip any repeated header rows that appear inside the worksheet data.
                        Dim headerNames As String() = {"NAME OF PDL", "NAME", "NAMES OF PDL"}
                        Dim headerNos As String() = {"NO", "NO."}
                        If headerNames.Any(Function(h) nameVal.Equals(h, StringComparison.OrdinalIgnoreCase)) OrElse
                           headerNos.Any(Function(h) noVal.Equals(h, StringComparison.OrdinalIgnoreCase)) Then
                            Continue For
                        End If

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

    Private Function ParseXmlOrSpreadsheetXml(sourcePath As String) As List(Of HearingRecord)
        Dim document = XDocument.Load(sourcePath)
        
        Dim hasRows = document.Descendants().Any(Function(element) element.Name.LocalName = "Row")
        If hasRows Then
            Return ParseSpreadsheetXml(sourcePath)
        End If

        Dim parsedList As New List(Of HearingRecord)()
        Dim elements = document.Descendants("Hearing")
        For Each el In elements
            Dim record As New HearingRecord With {
                .No = If(el.Element("CaseNo") IsNot Nothing, el.Element("CaseNo").Value.Trim(), If(el.Element("No") IsNot Nothing, el.Element("No").Value.Trim(), "")),
                .NameOfPdl = If(el.Element("NameOfPdl") IsNot Nothing, el.Element("NameOfPdl").Value.Trim(), ""),
                .BrCourt = If(el.Element("BrCourt") IsNot Nothing, el.Element("BrCourt").Value.Trim(), ""),
                .Hearing1 = If(el.Element("Hearing1") IsNot Nothing, el.Element("Hearing1").Value.Trim(), ""),
                .Hearing2 = If(el.Element("Hearing2") IsNot Nothing, el.Element("Hearing2").Value.Trim(), ""),
                .NextHearing = Date.MinValue
            }

            Dim nextHearingStr = If(el.Element("NextHearing") IsNot Nothing, el.Element("NextHearing").Value.Trim(), "")
            Dim nextHearingDate As Date
            If Date.TryParse(nextHearingStr, CultureInfo.InvariantCulture, DateTimeStyles.None, nextHearingDate) OrElse
               Date.TryParse(nextHearingStr, CultureInfo.CurrentCulture, DateTimeStyles.None, nextHearingDate) Then
                record.NextHearing = nextHearingDate.Date
            End If

            Dim historyEl = el.Element("HistoryLog")
            If historyEl IsNot Nothing Then
                For Each entryEl In historyEl.Elements("Log")
                    record.HistoryLog.Add(entryEl.Value)
                Next
            End If

            parsedList.Add(record)
        Next

        If parsedList.Count = 0 Then
            Dim root = document.Root
            If root IsNot Nothing Then
                For Each el In root.Elements()
                    Dim hasNo = el.Element("CaseNo") IsNot Nothing OrElse el.Element("No") IsNot Nothing
                    Dim hasName = el.Element("NameOfPdl") IsNot Nothing OrElse el.Element("Name") IsNot Nothing
                    If hasNo OrElse hasName Then
                        Dim record As New HearingRecord With {
                            .No = If(el.Element("CaseNo") IsNot Nothing, el.Element("CaseNo").Value.Trim(), If(el.Element("No") IsNot Nothing, el.Element("No").Value.Trim(), "")),
                            .NameOfPdl = If(el.Element("NameOfPdl") IsNot Nothing, el.Element("NameOfPdl").Value.Trim(), If(el.Element("Name") IsNot Nothing, el.Element("Name").Value.Trim(), "")),
                            .BrCourt = If(el.Element("BrCourt") IsNot Nothing, el.Element("BrCourt").Value.Trim(), If(el.Element("Court") IsNot Nothing, el.Element("Court").Value.Trim(), "")),
                            .Hearing1 = If(el.Element("Hearing1") IsNot Nothing, el.Element("Hearing1").Value.Trim(), If(el.Element("Hearing") IsNot Nothing, el.Element("Hearing").Value.Trim(), "")),
                            .Hearing2 = If(el.Element("Hearing2") IsNot Nothing, el.Element("Hearing2").Value.Trim(), ""),
                            .NextHearing = Date.MinValue
                        }
                        Dim nextHearingStr = If(el.Element("NextHearing") IsNot Nothing, el.Element("NextHearing").Value.Trim(), "")
                        Dim nextHearingDate As Date
                        If Date.TryParse(nextHearingStr, CultureInfo.InvariantCulture, DateTimeStyles.None, nextHearingDate) OrElse
                           Date.TryParse(nextHearingStr, CultureInfo.CurrentCulture, DateTimeStyles.None, nextHearingDate) Then
                            record.NextHearing = nextHearingDate.Date
                        End If
                        parsedList.Add(record)
                    End If
                Next
            End If
        End If

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

            If String.IsNullOrWhiteSpace(nameVal) Then
                Continue For
            End If

            ' Skip column-header rows that slipped through (e.g. "NAME OF PDL", "NO")
            Dim headerNames As String() = {"NAME OF PDL", "NAME", "NAMES OF PDL"}
            Dim headerNos   As String() = {"NO", "NO.", "#"}
            If headerNames.Any(Function(h) nameVal.Equals(h, StringComparison.OrdinalIgnoreCase)) OrElse
               headerNos.Any(Function(h) noVal.Equals(h, StringComparison.OrdinalIgnoreCase)) Then
                Continue For
            End If

            Dim hearingDate As Date = Date.MinValue
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

    Private Shared Function TryParseSectionDate(text As String, ByRef value As Date) As Boolean
        If String.IsNullOrWhiteSpace(text) Then
            Return False
        End If

        Dim cleanText = text.Trim()

        Dim parenIndex = cleanText.IndexOf("(")
        If parenIndex >= 0 Then
            cleanText = cleanText.Substring(0, parenIndex).Trim()
        End If

        Dim daysOfWeek As String() = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"}
        For Each day In daysOfWeek
            If cleanText.ToUpperInvariant().EndsWith(day) Then
                cleanText = cleanText.Substring(0, cleanText.Length - day.Length).Trim().Trim(","c).Trim()
                Exit For
            End If
        Next

        If Date.TryParse(cleanText, CultureInfo.CurrentCulture, DateTimeStyles.None, value) OrElse
           Date.TryParse(cleanText, CultureInfo.InvariantCulture, DateTimeStyles.None, value) Then
            Return True
        End If

        Dim formats As String() = {
            "MMMM d, yyyy", "MMMM dd, yyyy", "MMM d, yyyy", "MMM dd, yyyy",
            "d MMMM yyyy", "dd MMMM yyyy", "d MMM yyyy", "dd MMM yyyy"
        }
        Return Date.TryParseExact(cleanText, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, value) OrElse
               Date.TryParseExact(cleanText, formats, CultureInfo.CurrentCulture, DateTimeStyles.None, value)
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

    Private Shared Function ChooseColumn(col As Integer, mapping As ColumnMapping) As Integer
        Select Case col
            Case 1 : Return mapping.NoCol
            Case 2 : Return mapping.NameCol
            Case 3 : Return mapping.CourtCol
            Case 4 : Return mapping.Hearing1Col
            Case 5 : Return mapping.Hearing2Col
            Case 6 : Return mapping.NextHearingCol
            Case Else : Return col
        End Select
    End Function

    Public Sub ClearHearings(predicate As Func(Of HearingRecord, Boolean))
        If predicate Is Nothing Then
            Throw New ArgumentNullException(NameOf(predicate))
        End If

        Dim hearings = LoadHearings()
        hearings.RemoveAll(Function(h) predicate(h))
        SaveHearings(hearings)
    End Sub

    Public Sub ExportToExcel(destPath As String)
        Try
            Dim hearings = LoadHearings()
            Dim templatePath = Path.Combine(AppContext.BaseDirectory, "Data", "hearings.xlsx")
        
        ' Filter and group hearings by Month Year (excluding MinValue)
        Dim groupedHearings = hearings.
            Where(Function(h) h.NextHearing <> Date.MinValue).
            GroupBy(Function(h) New With {Key .Year = h.NextHearing.Year, Key .Month = h.NextHearing.Month}).
            OrderBy(Function(g) g.Key.Year).
            ThenBy(Function(g) g.Key.Month).
            ToList()

            If File.Exists(templatePath) Then
            Using workbook As New XLWorkbook(templatePath)
                Dim templateSheet = workbook.Worksheet(SheetName)
                Dim mapping = FindColumnMapping(templateSheet)
                
                If groupedHearings.Count > 0 Then
                    For Each g In groupedHearings
                        Dim sheetName = New DateTime(g.Key.Year, g.Key.Month, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture)
                        Dim newSheet = templateSheet.CopyTo(sheetName)
                        
                        ' Set column widths from template
                        For col = 1 To 6
                            newSheet.Column(col).Width = templateSheet.Column(col).Width
                        Next
                        
                        ' Set headers on Row 1
                        newSheet.Cell(1, 1).Value = "NO"
                        newSheet.Cell(1, 2).Value = "NAME OF PDL"
                        newSheet.Cell(1, 3).Value = "BR/COURT"
                        newSheet.Cell(1, 4).Value = "HEARING"
                        newSheet.Cell(1, 5).Value = "HEARING"
                        newSheet.Cell(1, 6).Value = "NEXT HEARING"
                        
                        ' Copy headers styling
                        For col = 1 To 6
                            Dim templateCol = ChooseColumn(col, mapping)
                            newSheet.Cell(1, col).Style = templateSheet.Cell(mapping.HeaderRow, templateCol).Style
                        Next
                        
                        ' Clear the rest of the sheet below header
                        Dim lastRow = newSheet.LastRowUsed()
                        Dim maxRow = If(lastRow Is Nothing, 100, Math.Max(100, lastRow.RowNumber()))
                        For r = 2 To maxRow
                            newSheet.Row(r).Clear()
                        Next
                        
                        ' Populate data starting on Row 2
                        Dim currentRow = 2
                        Dim templateDataRow = mapping.HeaderRow + 1
                        For Each h In g
                            newSheet.Cell(currentRow, 1).Value = h.No
                            newSheet.Cell(currentRow, 2).Value = h.NameOfPdl
                            newSheet.Cell(currentRow, 3).Value = h.BrCourt
                            newSheet.Cell(currentRow, 4).Value = h.Hearing1
                            newSheet.Cell(currentRow, 5).Value = h.Hearing2
                            If h.NextHearing <> Date.MinValue Then
                                newSheet.Cell(currentRow, 6).Value = h.NextHearing.Date
                                newSheet.Cell(currentRow, 6).Style.DateFormat.Format = "yyyy-mm-dd"
                            Else
                                newSheet.Cell(currentRow, 6).Value = ""
                            End If
                            
                            ' Copy style from template data row (mapping.HeaderRow + 1)
                            For col = 1 To 6
                                Dim templateCol = ChooseColumn(col, mapping)
                                newSheet.Cell(currentRow, col).Style = templateSheet.Cell(templateDataRow, templateCol).Style
                            Next
                            currentRow += 1
                        Next
                    Next
                    
                    ' Delete the original "Hearings" template sheet
                    workbook.Worksheet(SheetName).Delete()
                Else
                    ' If no scheduled hearings, keep template but clear data rows below header
                    Dim lastRow = LastUsedRowNumber(templateSheet)
                    For r = mapping.HeaderRow + 1 To lastRow
                        templateSheet.Row(r).Clear()
                    Next
                End If
                
                workbook.SaveAs(destPath)
            End Using
            Else
                ' Fallback if template is not found (highly resilient)
                Using workbook As New XLWorkbook()
                If groupedHearings.Count > 0 Then
                    For Each g In groupedHearings
                        Dim sheetName = New DateTime(g.Key.Year, g.Key.Month, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture)
                        Dim worksheet = workbook.Worksheets.Add(sheetName)
                        
                        worksheet.Cell(1, 1).Value = "NO"
                        worksheet.Cell(1, 2).Value = "NAME OF PDL"
                        worksheet.Cell(1, 3).Value = "BR/COURT"
                        worksheet.Cell(1, 4).Value = "HEARING"
                        worksheet.Cell(1, 5).Value = "HEARING"
                        worksheet.Cell(1, 6).Value = "NEXT HEARING"
                        
                        worksheet.Range(1, 1, 1, 6).Style.Font.Bold = True
                        worksheet.Range(1, 1, 1, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#12365d")
                        worksheet.Range(1, 1, 1, 6).Style.Font.FontColor = XLColor.White
                        
                        Dim row = 2
                        For Each h In g
                            worksheet.Cell(row, 1).Value = h.No
                            worksheet.Cell(row, 2).Value = h.NameOfPdl
                            worksheet.Cell(row, 3).Value = h.BrCourt
                            worksheet.Cell(row, 4).Value = h.Hearing1
                            worksheet.Cell(row, 5).Value = h.Hearing2
                            If h.NextHearing <> Date.MinValue Then
                                worksheet.Cell(row, 6).Value = h.NextHearing.Date
                                worksheet.Cell(row, 6).Style.DateFormat.Format = "yyyy-mm-dd"
                            Else
                                worksheet.Cell(row, 6).Value = ""
                            End If
                            row += 1
                        Next
                        worksheet.Columns().AdjustToContents()
                    Next
                Else
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
                End If
                    workbook.SaveAs(destPath)
                End Using
            End If
        Catch ex As Exception
            Throw New IOException($"Failed to export hearing data to Excel at '{destPath}'.", ex)
        End Try
    End Sub

    Public Sub ExportToCsv(destPath As String)
        Try
            Dim hearings = LoadHearings()
            Using writer As New StreamWriter(destPath, False, System.Text.Encoding.UTF8)
                writer.WriteLine("""NO"",""NAME OF PDL"",""BR/COURT"",""HEARING"",""HEARING"",""NEXT HEARING""")
                For Each h In hearings
                    Dim noEscaped = If(h.No IsNot Nothing, h.No.Replace("""", """"""), "")
                    Dim nameEscaped = If(h.NameOfPdl IsNot Nothing, h.NameOfPdl.Replace("""", """"""), "")
                    Dim courtEscaped = If(h.BrCourt IsNot Nothing, h.BrCourt.Replace("""", """"""), "")
                    Dim h1Escaped = If(h.Hearing1 IsNot Nothing, h.Hearing1.Replace("""", """"""), "")
                    Dim h2Escaped = If(h.Hearing2 IsNot Nothing, h.Hearing2.Replace("""", """"""), "")
                    Dim nextHearingStr = If(h.NextHearing = Date.MinValue, "", h.NextHearing.ToString("yyyy-MM-dd"))
                    writer.WriteLine($"""{noEscaped}"",""{nameEscaped}"",""{courtEscaped}"",""{h1Escaped}"",""{h2Escaped}"",""{nextHearingStr}""")
                Next
            End Using
        Catch ex As Exception
            Throw New IOException($"Failed to export hearing data to CSV at '{destPath}'.", ex)
        End Try
    End Sub

End Class
