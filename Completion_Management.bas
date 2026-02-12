Attribute VB_Name = "Completion_Management"

Option Explicit
Public Sub MoveToCompleted(srcTbl As ListObject, targetCell As Range)
    Dim wsCompleted As Worksheet
    Dim destTbl As ListObject
    Dim srcRow As ListRow
    Dim destRow As ListRow
    Dim originName As String
    Dim srcDescCol As Long, srcMitCol As Long, srcRespCol As Long, srcDeadCol As Long
    Dim isRoadblock As Boolean
    
    On Error GoTo CleanExit
    
    Set wsCompleted = GetSheetSafe("Completed")
    If wsCompleted Is Nothing Then Exit Sub
    
    isRoadblock = (InStr(1, srcTbl.Name, "Roadblocks", vbTextCompare) > 0)
    originName = srcTbl.Parent.Name
    
    If isRoadblock Then
        Set destTbl = wsCompleted.ListObjects("Roadblocks_Done")
    Else
        Set destTbl = wsCompleted.ListObjects("Risk_Done")
    End If
    
    srcDescCol = DescColIndex(srcTbl)
    srcMitCol = ColIndex(srcTbl, "Mitigating actions")
    srcDeadCol = ColIndex(srcTbl, "Deadline")
    
    If isRoadblock Then
        srcRespCol = ColIndex(srcTbl, "AIT PM Action Response")
    Else
        srcRespCol = ColIndex(srcTbl, "AIT PM Risk Response")
    End If

    Set srcRow = srcTbl.ListRows(targetCell.Row - srcTbl.DataBodyRange.Row + 1)
    
    Application.EnableEvents = False
    
    Set destRow = destTbl.ListRows.Add
    
    Call AddHyperlinkToCell(destRow.Range.Cells(1, 1), originName, originName)
    
    If srcDescCol > 0 Then destRow.Range.Cells(1, 2).Value = srcRow.Range.Cells(1, srcDescCol).Value
    If srcMitCol > 0 Then destRow.Range.Cells(1, 3).Value = srcRow.Range.Cells(1, srcMitCol).Value
    If srcRespCol > 0 Then destRow.Range.Cells(1, 4).Value = srcRow.Range.Cells(1, srcRespCol).Value
    If srcDeadCol > 0 Then destRow.Range.Cells(1, 5).Value = srcRow.Range.Cells(1, srcDeadCol).Value
    
    destRow.Range.Cells(1, 6).Value = Date
    
    srcRow.Range.ClearContents
    
CleanExit:
    Application.EnableEvents = True
End Sub
Public Sub MoveWinnerToCompleted(srcTbl As ListObject, targetCell As Range)
    Dim wsCompleted As Worksheet
    Dim destTbl As ListObject
    Dim srcRow As ListRow
    Dim destRow As ListRow
    Dim originName As String
    Dim successCol As Long, dateCol As Long
    
    On Error GoTo CleanExit
    
    Set wsCompleted = GetSheetSafe("Completed")
    If wsCompleted Is Nothing Then Exit Sub
    
    Set destTbl = wsCompleted.ListObjects("Wins_Done")
    If destTbl Is Nothing Then Exit Sub
    
    originName = srcTbl.Parent.Name
    
    successCol = ColIndex(srcTbl, "Success")
    If successCol = 0 Then successCol = 1
    
    dateCol = ColIndex(srcTbl, "Date")
    If dateCol = 0 Then dateCol = 2
    
    Set srcRow = srcTbl.ListRows(targetCell.Row - srcTbl.DataBodyRange.Row + 1)
    
    Application.EnableEvents = False
    
    Set destRow = destTbl.ListRows.Add
    
    Call AddHyperlinkToCell(destRow.Range.Cells(1, 1), originName, originName)
    
    If successCol > 0 Then destRow.Range.Cells(1, 2).Value = srcRow.Range.Cells(1, successCol).Value
    destRow.Range.Cells(1, 3).Value = ""
    If dateCol > 0 Then destRow.Range.Cells(1, 4).Value = srcRow.Range.Cells(1, dateCol).Value
    destRow.Range.Cells(1, 5).Value = Date
    
    srcRow.Range.ClearContents
    
CleanExit:
    Application.EnableEvents = True
End Sub

Public Sub MoveToCancelled(srcTbl As ListObject, targetCell As Range)
    Dim wsCancelled As Worksheet
    Dim destTbl As ListObject
    Dim srcRow As ListRow
    Dim destRow As ListRow
    Dim originName As String
    Dim srcDescCol As Long, srcMitigatingCol As Long, srcResponseCol As Long, srcDeadlineCol As Long
    Dim destDescCol As Long, destMitigatingCol As Long, destResponseCol As Long, destDeadlineCol As Long
    Dim destDateCol As Long
    Dim lc As ListColumn

    On Error GoTo CleanExit

    Set wsCancelled = GetSheetSafe("Cancelled")
    If wsCancelled Is Nothing Then GoTo CleanExit

    originName = srcTbl.Parent.Name

    If InStr(1, srcTbl.Name, "Roadblocks", vbTextCompare) > 0 Then
        Set destTbl = wsCancelled.ListObjects("Roadblocks_Cancelled")
        srcResponseCol = ColIndex(srcTbl, "AIT PM Action Response")
    Else
        Set destTbl = wsCancelled.ListObjects("Risks_Cancelled")
        srcResponseCol = ColIndex(srcTbl, "AIT PM Risk Response")
    End If
    If destTbl Is Nothing Then GoTo CleanExit

    Set srcRow = srcTbl.ListRows(targetCell.Row - srcTbl.DataBodyRange.Row + 1)
    If srcRow Is Nothing Then GoTo CleanExit

    Application.EnableEvents = False

    srcDescCol = DescColIndex(srcTbl)
    srcMitigatingCol = ColIndex(srcTbl, "Mitigating actions")

    srcDeadlineCol = 0
    For Each lc In srcTbl.ListColumns
        If InStr(1, lc.Name, "Deadline", vbTextCompare) > 0 Then
            srcDeadlineCol = lc.Index
            Exit For
        End If
    Next lc

    destDescCol = DescColIndex(destTbl)
    destMitigatingCol = ColIndex(destTbl, "Mitigating actions")

    If InStr(1, destTbl.Name, "Roadblocks", vbTextCompare) > 0 Then
        destResponseCol = ColIndex(destTbl, "AIT PM Action Response")
    Else
        destResponseCol = ColIndex(destTbl, "AIT PM Risk Response")
    End If

    destDeadlineCol = 0
    For Each lc In destTbl.ListColumns
        If InStr(1, lc.Name, "Deadline", vbTextCompare) > 0 Then
            destDeadlineCol = lc.Index
            Exit For
        End If
    Next lc

    destDateCol = ColIndex(destTbl, "Date Cancelled")
    If destDateCol = 0 Then destDateCol = ColIndex(destTbl, "Cancelled date")
    If destDateCol = 0 Then destDateCol = ColIndex(destTbl, "Date")

    Set destRow = destTbl.ListRows.Add

    Call AddHyperlinkToCell(destRow.Range.Cells(1, 1), originName, originName)

    If srcDescCol > 0 And destDescCol > 0 Then
        destRow.Range.Cells(1, destDescCol).Value = srcRow.Range.Cells(1, srcDescCol).Value
    End If

    If srcMitigatingCol > 0 And destMitigatingCol > 0 Then
        destRow.Range.Cells(1, destMitigatingCol).Value = srcRow.Range.Cells(1, srcMitigatingCol).Value
    End If

    If srcResponseCol > 0 And destResponseCol > 0 Then
        destRow.Range.Cells(1, destResponseCol).Value = srcRow.Range.Cells(1, srcResponseCol).Value
    End If

    If srcDeadlineCol > 0 And destDeadlineCol > 0 Then
        destRow.Range.Cells(1, destDeadlineCol).Value = srcRow.Range.Cells(1, srcDeadlineCol).Value
    End If

    If destDateCol > 0 Then
        destRow.Range.Cells(1, destDateCol).Value = Date
    End If

    srcRow.Range.ClearContents

CleanExit:
    Application.EnableEvents = True
End Sub

Public Sub MoveToCompletedFromOverview(ovTbl As ListObject, targetCell As Range)
    Dim wsCompleted As Worksheet
    Dim destTbl As ListObject
    Dim destRow As ListRow
    Dim originName As String, shortDesc As String
    Dim ovRowIdx As Long
    Dim srcWs As Worksheet, srcTbl As ListObject, i As Long
    
    On Error GoTo CleanExit
    Set wsCompleted = GetSheetSafe("Completed")
    
    ovRowIdx = targetCell.Row - ovTbl.DataBodyRange.Row + 1
    
    originName = Trim$(NzStr(ovTbl.DataBodyRange(ovRowIdx, 1).Value))
    If ovTbl.DataBodyRange(ovRowIdx, 1).Hyperlinks.count > 0 Then
        originName = ovTbl.DataBodyRange(ovRowIdx, 1).Hyperlinks(1).TextToDisplay
    End If
    
    If InStr(ovTbl.Name, "Roadblocks") > 0 Then
        Set destTbl = wsCompleted.ListObjects("Roadblocks_Done")
    Else
        Set destTbl = wsCompleted.ListObjects("Risk_Done")
    End If
    
    Application.EnableEvents = False
    
    Set destRow = destTbl.ListRows.Add
    Call AddHyperlinkToCell(destRow.Range.Cells(1, 1), originName, originName)
    
    destRow.Range.Cells(1, 2).Value = ovTbl.DataBodyRange(ovRowIdx, 3).Value
    destRow.Range.Cells(1, 3).Value = ovTbl.DataBodyRange(ovRowIdx, 4).Value
    destRow.Range.Cells(1, 4).Value = ovTbl.DataBodyRange(ovRowIdx, 5).Value
    destRow.Range.Cells(1, 5).Value = ovTbl.DataBodyRange(ovRowIdx, 6).Value
    destRow.Range.Cells(1, 6).Value = Date
    
    shortDesc = NormalizeText(NzStr(ovTbl.DataBodyRange(ovRowIdx, 3).Value))
    Set srcWs = GetSheetSafe(originName)
    
    If Not srcWs Is Nothing Then
        For Each srcTbl In srcWs.ListObjects
            If (InStr(ovTbl.Name, "Roadblock") > 0 And InStr(destTbl.Name, "Roadblock") > 0) Or _
               (InStr(ovTbl.Name, "Risk") > 0 And InStr(destTbl.Name, "Risk") > 0) Then
               For i = 1 To srcTbl.ListRows.count
                    If NormalizeText(NzStr(srcTbl.DataBodyRange(i, DescColIndex(srcTbl)).Value)) = shortDesc Then
                        srcTbl.ListRows(i).Range.ClearContents
                        Exit For
                    End If
                Next i
            End If
        Next srcTbl
    End If
    
    ovTbl.ListRows(ovRowIdx).Delete
    
CleanExit:
    Application.EnableEvents = True
End Sub

