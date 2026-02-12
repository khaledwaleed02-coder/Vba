Attribute VB_Name = "Sheet_Management"

Option Explicit

'========================================
' SHEET MANAGEMENT MODULE
' Handles creation and organization of PM sheets
'========================================

Public Sub CreateSheetsFromTable()
    '---
    ' Create new PM sheets from the Create table
    '---
    Dim wsTemplate As Worksheet
    Dim wsOverview As Worksheet
    Dim wsCreate As Worksheet
    Dim tblCreate As ListObject
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim i As Long
    Dim afterIndex As Long
    
    Set wsTemplate = GetSheetSafe("Template")
    Set wsOverview = GetSheetSafe("Overview")
    Set wsCreate = GetSheetSafe("Create")
    
    If wsTemplate Is Nothing Or wsOverview Is Nothing Or wsCreate Is Nothing Then
        MsgBox "Required sheets not found!", vbCritical
        Exit Sub
    End If
    
    Set tblCreate = wsCreate.ListObjects("Create")
    If tblCreate Is Nothing Then
        MsgBox "Create table not found!", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Create sheets from table
    For i = 1 To tblCreate.ListRows.count
        ' Stop if row is empty
        If Trim$(NzStr(tblCreate.DataBodyRange(i, 1).Value)) = "" And _
           Trim$(NzStr(tblCreate.DataBodyRange(i, 2).Value)) = "" And _
           Trim$(NzStr(tblCreate.DataBodyRange(i, 3).Value)) = "" Then
            Exit For
        End If
        
        ' Build sheet name: Project_Package_WPOPM
        sheetName = tblCreate.DataBodyRange(i, 1).Value & "_" & _
                    tblCreate.DataBodyRange(i, 2).Value & "_" & _
                    tblCreate.DataBodyRange(i, 3).Value
        
        ' Sanitize sheet name (remove invalid characters)
        sheetName = SanitizeSheetName(sheetName)
        
        ' Skip if sheet already exists
        If Not GetSheetSafe(sheetName) Is Nothing Then
            Debug.Print "Sheet already exists: " & sheetName
            GoTo NextSheet
        End If
        
        ' Copy template
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        newSheet.Name = sheetName
        
        ' Fill in header data
        newSheet.Range("B2").Value = tblCreate.DataBodyRange(i, 1).Value ' Project
        newSheet.Range("C2").Value = tblCreate.DataBodyRange(i, 2).Value ' Package
        newSheet.Range("D2").Value = tblCreate.DataBodyRange(i, 3).Value ' WPO/PM
        
        ' Reset tab color
        newSheet.Tab.ColorIndex = xlColorIndexNone
        
NextSheet:
        Set newSheet = Nothing
    Next i
    
    ' Move new sheets after Overview
    afterIndex = wsOverview.Index
    For i = tblCreate.ListRows.count To 1 Step -1
        If Trim$(NzStr(tblCreate.DataBodyRange(i, 1).Value)) = "" Then GoTo NextMove
        
        sheetName = tblCreate.DataBodyRange(i, 1).Value & "_" & _
                    tblCreate.DataBodyRange(i, 2).Value & "_" & _
                    tblCreate.DataBodyRange(i, 3).Value
        sheetName = SanitizeSheetName(sheetName)
        
        Set newSheet = GetSheetSafe(sheetName)
        If Not newSheet Is Nothing Then
            newSheet.Move After:=ThisWorkbook.Sheets(afterIndex)
        End If
        
NextMove:
        Set newSheet = Nothing
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Sheets created successfully!", vbInformation
End Sub

Private Function SanitizeSheetName(ByVal sheetName As String) As String
    '---
    ' Remove invalid characters from sheet name
    '---
    Dim invalidChars As Variant
    Dim char As Variant
    
    invalidChars = Array("/", "\", ":", "?", "*", "[", "]")
    
    For Each char In invalidChars
        sheetName = Replace(sheetName, char, "-")
    Next char
    
    ' Limit to 31 characters (Excel limit)
    SanitizeSheetName = Left$(sheetName, 31)
End Function

Public Sub GoToOverview()
    '---
    ' Navigate to Overview sheet
    '---
    Dim wsOverview As Worksheet
    
    Set wsOverview = GetSheetSafe("Overview")
    If Not wsOverview Is Nothing Then
        wsOverview.Activate
        wsOverview.Range("A1").Select
    End If
End Sub

Public Sub SetupProgressStatusDropdowns()
    '---
    ' Add Progress Status dropdown to all Roadblocks and Risk tables
    '---
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIdx As Long
    Dim updatedCount As Long
    Dim cell As Range
    
    updatedCount = 0
    
    For Each ws In ThisWorkbook.Worksheets
        If IsPMSheet(ws) Then
            For Each tbl In ws.ListObjects
                ' Match by prefix
                If Left$(tbl.Name, 10) = "Roadblocks" Or Left$(tbl.Name, 4) = "Risk" Then
                    ' Ensure Progress Status column exists
                    colIdx = EnsureColumn(tbl, "Progress Status")
                    
                    If Not tbl.DataBodyRange Is Nothing Then
                        ' Add dropdown validation
                        With tbl.ListColumns(colIdx).DataBodyRange.Validation
                            .Delete
                            .Add Type:=xlValidateList, _
                                 AlertStyle:=xlValidAlertStop, _
                                 Operator:=xlBetween, _
                                 Formula1:="In progress,Awaiting,Completed,Resolved"
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .ShowError = True
                        End With
                        
                        ' Fill blanks with "Awaiting"
                        For Each cell In tbl.ListColumns(colIdx).DataBodyRange
                            If Trim$(NzStr(cell.Value)) = "" Then
                                cell.Value = "Awaiting"
                            End If
                        Next cell
                        
                        updatedCount = updatedCount + 1
                    End If
                End If
            Next tbl
        End If
    Next ws
    
    MsgBox "Updated " & updatedCount & " table(s) with Progress Status dropdowns.", vbInformation
End Sub

