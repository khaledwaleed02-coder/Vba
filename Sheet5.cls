VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'========================================
' OVERVIEW SHEET EVENT HANDLER
' Handles bi-directional sync when editing overview
'========================================

Private Sub Worksheet_Change(ByVal Target As Range)
    '---
    ' Sync changes from Overview back to source PM sheets
    ' AND Handle Completion/Cancellation from Overview
    '---
    Dim tbl As ListObject
    Dim colIdx As Long
    Dim cell As Range
    Dim ovRow As Long
    Dim statusValue As String
    
    On Error GoTo CleanExit
    If Target.CountLarge > 1 Then Exit Sub
    
    ' === CHECK FOR COMPLETION/CANCELLATION ===
    ' Iterate through both Overview tables to check for status changes
    For Each tbl In Me.ListObjects
        If tbl.Name = "Roadblocks_Overview" Or tbl.Name = "Risk_Overview" Then
            colIdx = ColIndex(tbl, "Progress Status")
            
            If colIdx > 0 Then
                If Not Intersect(Target, tbl.ListColumns(colIdx).DataBodyRange) Is Nothing Then
                    statusValue = LCase$(Trim$(NzStr(Target.Value)))
                    
                    Application.EnableEvents = False ' Prevent loops
                    
                    If statusValue = "completed" Or statusValue = "resolved" Then
                        Call Completion_Management.MoveToCompletedFromOverview(tbl, Target)
                        GoTo CleanExit ' Row is deleted, stop processing
                    ElseIf statusValue = "cancelled" Then
                        ' You might want to add a MoveToCancelledFromOverview if needed,
                        ' otherwise it just stays as cancelled.
                        ' Call Completion_Management.MoveToCancelledFromOverview(tbl, Target)
                    End If
                    
                    Application.EnableEvents = True
                End If
            End If
        End If
    Next tbl
    
    ' === EXISTING SYNC LOGIC (Keep your existing sync logic below) ===
    ' Handle Roadblocks Sync (Text)
    Set tbl = Me.ListObjects("Roadblocks_Overview")
    If Not tbl Is Nothing Then
        colIdx = ColIndex(tbl, "AIT PM Action Response")
        If colIdx > 0 Then
            If Not Intersect(Target, tbl.ListColumns(colIdx).DataBodyRange) Is Nothing Then
                Application.EnableEvents = False
                For Each cell In Intersect(Target, tbl.ListColumns(colIdx).DataBodyRange)
                    ovRow = cell.Row - tbl.DataBodyRange.Row + 1
                    Call SyncRoadblockToSource(tbl, ovRow)
                Next cell
                Application.StatusBar = "Roadblock response synced back."
                Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
            End If
        End If
    End If
    
    ' Handle Risks Sync (Text)
    Set tbl = Me.ListObjects("Risk_Overview")
    If Not tbl Is Nothing Then
        colIdx = ColIndex(tbl, "AIT PM Risk Response")
        If colIdx > 0 Then
            If Not Intersect(Target, tbl.ListColumns(colIdx).DataBodyRange) Is Nothing Then
                Application.EnableEvents = False
                For Each cell In Intersect(Target, tbl.ListColumns(colIdx).DataBodyRange)
                    ovRow = cell.Row - tbl.DataBodyRange.Row + 1
                    Call SyncRiskToSource(tbl, ovRow)
                Next cell
                Application.StatusBar = "Risk response synced back."
                Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
            End If
        End If
    End If
    
CleanExit:
    Application.EnableEvents = True
End Sub

Private Sub SyncRoadblockToSource(ovTbl As ListObject, ovRow As Long)
    '---
    ' Sync roadblock response from overview to source PM sheet
    '---
    Dim originName As String
    Dim shortDesc As String
    Dim response As String
    Dim srcWs As Worksheet
    Dim srcTbl As ListObject
    Dim i As Long
    Dim descCol As Long
    Dim respCol As Long
    Dim matched As Boolean
    
    ' Get overview data
    originName = NzStr(ovTbl.DataBodyRange(ovRow, 1).Value)
    shortDesc = NormalizeText(NzStr(ovTbl.DataBodyRange(ovRow, 3).Value))
    response = NzStr(ovTbl.DataBodyRange(ovRow, 5).Value)
    
    If originName = "" Or shortDesc = "" Then Exit Sub
    
    ' Clean hyperlink
    If ovTbl.DataBodyRange(ovRow, 1).Hyperlinks.count > 0 Then
        originName = ovTbl.DataBodyRange(ovRow, 1).Hyperlinks(1).TextToDisplay
    End If
    
    Set srcWs = GetSheetSafe(Trim$(originName))
    If srcWs Is Nothing Then Exit Sub
    
    ' Find source table
    For Each srcTbl In srcWs.ListObjects
        If Left$(srcTbl.Name, 10) = "Roadblocks" Then
            descCol = ColIndex(srcTbl, "Roadblock description")
            If descCol = 0 Then descCol = 2
            
            respCol = EnsureColumn(srcTbl, "AIT PM Action Response")
            
            ' Find matching row
            For i = 1 To srcTbl.ListRows.count
                If NormalizeText(NzStr(srcTbl.DataBodyRange(i, descCol).Value)) = shortDesc Then
                    srcTbl.DataBodyRange(i, respCol).Value = response
                    matched = True
                    Exit For
                End If
            Next i
            Exit For
        End If
    Next srcTbl
    
    ' Highlight if not matched
    If Not matched Then
        ovTbl.DataBodyRange(ovRow, 5).Interior.Color = RGB(255, 220, 220)
    Else
        ovTbl.DataBodyRange(ovRow, 5).Interior.ColorIndex = xlColorIndexNone
    End If
End Sub

Private Sub SyncRiskToSource(ovTbl As ListObject, ovRow As Long)
    '---
    ' Sync risk response from overview to source PM sheet
    '---
    Dim originName As String
    Dim shortDesc As String
    Dim response As String
    Dim srcWs As Worksheet
    Dim srcTbl As ListObject
    Dim i As Long
    Dim descCol As Long
    Dim respCol As Long
    Dim matched As Boolean
    
    ' Get overview data
    originName = NzStr(ovTbl.DataBodyRange(ovRow, 1).Value)
    shortDesc = NormalizeText(NzStr(ovTbl.DataBodyRange(ovRow, 3).Value))
    response = NzStr(ovTbl.DataBodyRange(ovRow, 5).Value)
    
    If originName = "" Or shortDesc = "" Then Exit Sub
    
    ' Clean hyperlink
    If ovTbl.DataBodyRange(ovRow, 1).Hyperlinks.count > 0 Then
        originName = ovTbl.DataBodyRange(ovRow, 1).Hyperlinks(1).TextToDisplay
    End If
    
    Set srcWs = GetSheetSafe(Trim$(originName))
    If srcWs Is Nothing Then Exit Sub
    
    ' Find source table
    For Each srcTbl In srcWs.ListObjects
        If Left$(srcTbl.Name, 4) = "Risk" Then
            descCol = ColIndex(srcTbl, "Risk description")
            If descCol = 0 Then descCol = 2
            
            respCol = EnsureColumn(srcTbl, "AIT PM Risk Response")
            
            ' Find matching row
            For i = 1 To srcTbl.ListRows.count
                If NormalizeText(NzStr(srcTbl.DataBodyRange(i, descCol).Value)) = shortDesc Then
                    srcTbl.DataBodyRange(i, respCol).Value = response
                    matched = True
                    Exit For
                End If
            Next i
            Exit For
        End If
    Next srcTbl
    
    ' Highlight if not matched
    If Not matched Then
        ovTbl.DataBodyRange(ovRow, 5).Interior.Color = RGB(255, 220, 220)
    Else
        ovTbl.DataBodyRange(ovRow, 5).Interior.ColorIndex = xlColorIndexNone
    End If
End Sub

Public Sub ClearStatusBar()
    '---
    ' Clear Excel status bar (called via OnTime)
    '---
    Application.StatusBar = False
End Sub
