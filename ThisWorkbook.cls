VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'========================================
' WORKBOOK EVENT HANDLER
' Handles completion logic across all sheets
'========================================

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim lo As ListObject
    Dim colIdx As Long
    Dim statusValue As String
    
    ' 1. Guard Clauses: Single cell edits only
    If Target.CountLarge > 1 Then Exit Sub
    
    ' 2. Check if we are inside a Table (ListObject)
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    ' ========== WINNERS/SUCCESS HANDLING ==========
    ' Handle any Winners/Success tables (matches "Winners" prefix)
    If Left$(lo.Name, 7) = "Winners" Then
        colIdx = ColIndex(lo, "Send into Complete(x)")
        If colIdx > 0 Then
            If Not Intersect(Target, lo.ListColumns(colIdx).DataBodyRange) Is Nothing Then
                statusValue = LCase$(Trim$(NzStr(Target.Value)))
                If statusValue = "x" Then
                    Application.EnableEvents = False
                    Call MoveWinnerToCompleted(lo, Target)
                    Application.EnableEvents = True
                    Exit Sub
                End If
            End If
        End If
    End If
    
    ' ========== ROADBLOCKS/RISKS HANDLING ==========
    ' Only run on PM Sheets for Roadblocks/Risks
    If Not IsPMSheet(Sh) Then Exit Sub
    
    ' 3. Identify if it is a Roadblock or Risk table
    If Left$(lo.Name, 10) <> "Roadblocks" And Left$(lo.Name, 4) <> "Risk" Then Exit Sub
    
    ' 4. Check if the changed column is "Progress Status"
    colIdx = ColIndex(lo, "Progress Status")
    If colIdx = 0 Then Exit Sub
    
    ' Check intersection to ensure we are editing the Status column
    If Not Intersect(Target, lo.ListColumns(colIdx).DataBodyRange) Is Nothing Then
        statusValue = LCase$(Trim$(NzStr(Target.Value)))
        
        ' 5. Trigger the logic in Completion_Management module
        If statusValue = "completed" Or statusValue = "resolved" Then
            Call MoveToCompleted(lo, Target)
        ElseIf statusValue = "cancelled" Then
            Call MoveToCancelled(lo, Target)
        End If
    End If
End Sub
