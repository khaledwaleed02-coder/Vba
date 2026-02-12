Attribute VB_Name = "Overview_Update"
Option Explicit

Public Sub UpdateOverview()
    '---
    ' Main procedure to update all overview tables from PM sheets
    '---
    Dim oldCalc As XlCalculation
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== UpdateOverview Started ==="
    
    ' Optimize performance
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Update only the tables that exist
    Debug.Print "Step 1: Copying Roadblocks..."
    Call CopyRoadblocks
    
    Debug.Print "Step 2: Copying Risks..."
    Call CopyRisks
    
    Debug.Print "Step 3: Copying Winners..."
    Call CopyWinners
    
    Debug.Print "Step 4: Updating tab colors..."
    Call UpdateTabColors
    
    Debug.Print "=== UpdateOverview Completed Successfully ==="
    MsgBox "Overview updated successfully!", vbInformation

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = oldCalc
    
    If Err.Number <> 0 Then
        Debug.Print "ERROR: " & Err.Description
        MsgBox "Error: " & Err.Description & " (Error #" & Err.Number & ")", vbCritical
    End If
End Sub

Private Sub CopyRoadblocks()
    '---
    ' Copy escalated roadblocks to overview
    '---
    Dim ws As Worksheet
    Dim srcTbl As ListObject
    Dim destTbl As ListObject
    Dim destRow As ListRow
    Dim i As Long
    Dim esclCol As Long
    Dim shortDescCol As Long, mitigatingCol As Long, responseCol As Long, deadlineCol As Long, progressCol As Long
    
    Dim preservedResponses As Object
    Set preservedResponses = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set destTbl = ThisWorkbook.Worksheets("Overview").ListObjects("Roadblocks_Overview")
    On Error GoTo 0
    
    If destTbl Is Nothing Then Exit Sub
    
    ' Store existing responses
    responseCol = ColIndex(destTbl, "AIT PM Action Response")
    shortDescCol = ColIndex(destTbl, "Roadblock description")
    If responseCol > 0 And shortDescCol > 0 Then
        Dim j As Long
        For j = 1 To destTbl.ListRows.count
            Dim key As String
            key = NormalizeText(NzStr(destTbl.DataBodyRange(j, shortDescCol).Value))
            If key <> "" Then
                preservedResponses(key) = NzStr(destTbl.DataBodyRange(j, responseCol).Value)
            End If
        Next j
    End If
    
    Call ClearTableData(destTbl)
    
    For Each ws In ThisWorkbook.Worksheets
        If IsPMSheet(ws) Then
            For Each srcTbl In ws.ListObjects
                If Left$(srcTbl.Name, 10) = "Roadblocks" Then
                    esclCol = ColIndex(srcTbl, "Escl (initial)")
                    If esclCol = 0 Then esclCol = 7
                    
                    shortDescCol = DescColIndex(srcTbl)
                    If shortDescCol = 0 Then shortDescCol = 2
                    
                    mitigatingCol = ColIndex(srcTbl, "Mitigating actions")
                    If mitigatingCol = 0 Then mitigatingCol = 3
                    
                    responseCol = ColIndex(srcTbl, "AIT PM Action Response")
                    If responseCol = 0 Then responseCol = 4
                    
                    deadlineCol = ColIndex(srcTbl, "Deadline")
                    If deadlineCol = 0 Then deadlineCol = 6
                    
                    progressCol = ColIndex(srcTbl, "Progress Status")
                    If progressCol = 0 Then progressCol = 1
                    
                    For i = 1 To srcTbl.ListRows.count
                        If Len(Trim$(NzStr(srcTbl.DataBodyRange(i, esclCol).Value))) > 0 Then
                            Set destRow = destTbl.ListRows.Add
                            Call AddHyperlinkToCell(destRow.Range(1, 1), ws.Name, ws.Name)
                            
                            destRow.Range(1, 2).Value = srcTbl.DataBodyRange(i, progressCol).Value
                            destRow.Range(1, 3).Value = srcTbl.DataBodyRange(i, shortDescCol).Value
                            destRow.Range(1, 4).Value = srcTbl.DataBodyRange(i, mitigatingCol).Value
                            destRow.Range(1, 5).Value = srcTbl.DataBodyRange(i, responseCol).Value
                            destRow.Range(1, 6).Value = srcTbl.DataBodyRange(i, deadlineCol).Value
                            destRow.Range(1, 7).Value = srcTbl.DataBodyRange(i, esclCol).Value
                            
                            Dim descKey As String
                            descKey = NormalizeText(NzStr(srcTbl.DataBodyRange(i, shortDescCol).Value))
                            If preservedResponses.Exists(descKey) Then
                                Dim preservedResp As String
                                preservedResp = preservedResponses(descKey)
                                If preservedResp <> "" Then
                                    destRow.Range(1, 5).Value = preservedResp
                                    srcTbl.DataBodyRange(i, responseCol).Value = preservedResp
                                End If
                            End If
                        End If
                    Next i
                End If
            Next srcTbl
        End If
    Next ws
End Sub

Private Sub CopyRisks()
    '---
    ' Copy escalated risks to overview
    '---
    Dim ws As Worksheet
    Dim srcTbl As ListObject
    Dim destTbl As ListObject
    Dim destRow As ListRow
    Dim i As Long
    Dim esclCol As Long
    Dim shortDescCol As Long, mitigatingCol As Long, responseCol As Long, deadlineCol As Long, progressCol As Long
    
    Dim preservedResponses As Object
    Set preservedResponses = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set destTbl = ThisWorkbook.Worksheets("Overview").ListObjects("Risk_Overview")
    On Error GoTo 0
    
    If destTbl Is Nothing Then Exit Sub
    
    ' Store existing responses
    responseCol = ColIndex(destTbl, "AIT PM Risk Response")
    shortDescCol = ColIndex(destTbl, "Risk description")
    If responseCol > 0 And shortDescCol > 0 Then
        Dim j As Long
        For j = 1 To destTbl.ListRows.count
            Dim key As String
            key = NormalizeText(NzStr(destTbl.DataBodyRange(j, shortDescCol).Value))
            If key <> "" Then
                preservedResponses(key) = NzStr(destTbl.DataBodyRange(j, responseCol).Value)
            End If
        Next j
    End If
    
    Call ClearTableData(destTbl)
    
    For Each ws In ThisWorkbook.Worksheets
        If IsPMSheet(ws) Then
            For Each srcTbl In ws.ListObjects
                If Left$(srcTbl.Name, 4) = "Risk" Then
                    esclCol = ColIndex(srcTbl, "Escl (initial)")
                    If esclCol = 0 Then esclCol = 8
                    
                    shortDescCol = DescColIndex(srcTbl)
                    If shortDescCol = 0 Then shortDescCol = 2
                    
                    mitigatingCol = ColIndex(srcTbl, "Mitigating actions")
                    If mitigatingCol = 0 Then mitigatingCol = 3
                    
                    responseCol = ColIndex(srcTbl, "AIT PM Risk Response")
                    If responseCol = 0 Then responseCol = 4
                    
                    deadlineCol = ColIndex(srcTbl, "Deadline")
                    If deadlineCol = 0 Then deadlineCol = 7
                    
                    progressCol = ColIndex(srcTbl, "Progress Status")
                    If progressCol = 0 Then progressCol = 1
                    
                    For i = 1 To srcTbl.ListRows.count
                        If Len(Trim$(NzStr(srcTbl.DataBodyRange(i, esclCol).Value))) > 0 Then
                            Set destRow = destTbl.ListRows.Add
                            Call AddHyperlinkToCell(destRow.Range(1, 1), ws.Name, ws.Name)
                            
                            destRow.Range(1, 2).Value = srcTbl.DataBodyRange(i, progressCol).Value
                            destRow.Range(1, 3).Value = srcTbl.DataBodyRange(i, shortDescCol).Value
                            destRow.Range(1, 4).Value = srcTbl.DataBodyRange(i, mitigatingCol).Value
                            destRow.Range(1, 5).Value = srcTbl.DataBodyRange(i, responseCol).Value
                            destRow.Range(1, 6).Value = srcTbl.DataBodyRange(i, deadlineCol).Value
                            destRow.Range(1, 7).Value = srcTbl.DataBodyRange(i, esclCol).Value
                            
                            Dim descKey As String
                            descKey = NormalizeText(NzStr(srcTbl.DataBodyRange(i, shortDescCol).Value))
                            If preservedResponses.Exists(descKey) Then
                                Dim preservedResp As String
                                preservedResp = preservedResponses(descKey)
                                If preservedResp <> "" Then
                                    destRow.Range(1, 5).Value = preservedResp
                                    srcTbl.DataBodyRange(i, responseCol).Value = preservedResp
                                End If
                            End If
                        End If
                    Next i
                End If
            Next srcTbl
        End If
    Next ws
End Sub

Private Sub CopyWinners()
    '---
    ' Copy winners to overview
    '---
    Dim ws As Worksheet
    Dim srcTbl As ListObject
    Dim destTbl As ListObject
    Dim destRow As ListRow
    Dim i As Long
    
    On Error Resume Next
    Set destTbl = ThisWorkbook.Worksheets("Overview").ListObjects("Winners_Overview")
    On Error GoTo 0
    
    If destTbl Is Nothing Then Exit Sub
    
    Call ClearTableData(destTbl)
    
    For Each ws In ThisWorkbook.Worksheets
        If IsPMSheet(ws) Then
            For Each srcTbl In ws.ListObjects
                If Left$(srcTbl.Name, 7) = "Winners" Then
                    For i = 1 To srcTbl.ListRows.count
                        If Len(Trim$(NzStr(srcTbl.DataBodyRange(i, 1).Value))) > 0 Then
                            Set destRow = destTbl.ListRows.Add
                            Call AddHyperlinkToCell(destRow.Range(1, 1), ws.Name, ws.Name)
                            destRow.Range(1, 2).Value = srcTbl.DataBodyRange(i, 1).Value
                            destRow.Range(1, 3).Value = srcTbl.DataBodyRange(i, 2).Value
                        End If
                    Next i
                End If
            Next srcTbl
        End If
    Next ws
End Sub

Private Sub UpdateTabColors()
    '---
    ' Update sheet tab colors based on overall status
    '---
    Dim ws As Worksheet
    Dim excludeSheets As Variant
    
    excludeSheets = Array("Overview", "Template", "Create", "Completed")
    
    For Each ws In ThisWorkbook.Worksheets
        If IsError(Application.Match(ws.Name, excludeSheets, 0)) Then
            On Error Resume Next
            ws.Tab.Color = ws.Range("C4").Interior.Color
            On Error GoTo 0
        End If
    Next ws
End Sub
