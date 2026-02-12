Attribute VB_Name = "Core_Utilities"
Option Explicit

'========================================
' CORE UTILITY FUNCTIONS
' Shared helper functions used across all modules
'========================================

Public Function ColIndex(lo As ListObject, headerName As String) As Long
    '---
    ' Find column index by header name (case-insensitive)
    ' Returns: Column index, or 0 if not found
    '---
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, headerName, vbTextCompare) = 0 Then
            ColIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColIndex = 0
End Function

Public Function DescColIndex(lo As ListObject) As Long
    'Returns the description column index for Roadblocks/Risks tables (current headers only)
    'Returns 0 if not found (no error raised)
    Dim idx As Long

    If Left$(lo.Name, 10) = "Roadblocks" Then
        idx = ColIndex(lo, "Roadblock description")
    ElseIf Left$(lo.Name, 4) = "Risk" Then
        idx = ColIndex(lo, "Risk description")
    Else
        idx = 0
    End If

    DescColIndex = idx
End Function


Public Function EnsureColumn(lo As ListObject, headerName As String) As Long
    '---
    ' Ensure a column exists in the table, create if missing
    ' Returns: Column index
    '---
    Dim idx As Long
    idx = ColIndex(lo, headerName)
    If idx = 0 Then
        idx = lo.ListColumns.Add.Index
        lo.ListColumns(idx).Name = headerName
    End If
    EnsureColumn = idx
End Function

Public Function NzStr(v As Variant) As String
    '---
    ' Convert variant to string, handling errors and nulls
    '---
    If IsError(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function NormalizeText(s As String) As String
    '---
    ' Normalize text for comparison (lowercase, trim, collapse spaces)
    '---
    s = LCase$(Trim$(s))
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeText = s
End Function

Public Function GetSheetSafe(sheetName As String) As Worksheet
    '---
    ' Safely get worksheet by name, returns Nothing if not found
    '---
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set GetSheetSafe = ws
End Function

Public Function IsPMSheet(ws As Worksheet) As Boolean
    '---
    ' Check if sheet is a PM one-pager (PPV or MA prefix)
    '---
    IsPMSheet = (Left$(ws.Name, 3) = "PPV" Or Left$(ws.Name, 2) = "MA")
End Function

Public Sub ClearTableData(tbl As ListObject)
    '---
    ' Clear all rows from a table
    '---
    On Error Resume Next
    Do While tbl.ListRows.count > 0
        tbl.ListRows(1).Delete
    Loop
    On Error GoTo 0
End Sub

Public Sub AddHyperlinkToCell(cell As Range, sheetName As String, displayText As String)
    '---
    ' Add hyperlink to a cell pointing to a sheet
    '---
    On Error Resume Next
    cell.Hyperlinks.Add _
        Anchor:=cell, _
        Address:="", _
        SubAddress:="'" & sheetName & "'!A1", _
        TextToDisplay:=displayText
    On Error GoTo 0
End Sub

Public Function FindTableByPrefix(ws As Worksheet, prefix As String) As ListObject
    '---
    ' Find first table whose name starts with the given prefix
    ' Returns: ListObject or Nothing if not found
    '---
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Left$(tbl.Name, Len(prefix)) = prefix Then
            Set FindTableByPrefix = tbl
            Exit Function
        End If
    Next tbl
    Set FindTableByPrefix = Nothing
End Function
