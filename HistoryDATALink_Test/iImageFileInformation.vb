
Public Interface iImageFileInformation
    Function GetHistoryData(InPath As String, HistoryData As ImageFileInfo) As Boolean
    Property ImageFileInfo As ImageFileInfo
End Interface

Public Structure ImageFileInfo
    Dim MatchRank As Integer
    Dim MatchedFile As String
    Dim EventPeriod As String
    Dim Description As String
End Structure