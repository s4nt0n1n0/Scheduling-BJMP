Public Class HearingRecord
    Public Property Id As Integer
    Public Property No As String = ""
    Public Property NameOfPdl As String = ""
    Public Property BrCourt As String = ""
    Public Property Hearing1 As String = ""
    Public Property Hearing2 As String = ""
    Public Property NextHearing As Date

    ''' <summary>
    ''' In-memory history log for date changes. Not saved to Excel.
    ''' Each entry is a formatted string like "May 27, 2026 12:30 PM — Moved from 2026-05-20 to 2026-05-27"
    ''' </summary>
    Public Property HistoryLog As New List(Of String)()
End Class
