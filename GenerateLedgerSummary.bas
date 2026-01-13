Attribute VB_Name = "Module4"
Sub GenSummary()
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim statusSheet As Worksheet
    Dim folderPath As String, openFolderPath As String, closedFolderPath As String
    Dim newFilePath As String, newWorkbook As Workbook, newSheet As Worksheet
    Dim uniqueMatters As Collection, statusCollection As Collection
    Dim matterCell As Range, matterName As Variant
    Dim lastRow As Long, openSummaryRow As Long, closedSummaryRow As Long
    Dim finalBalance As Double, lastBalance As Double
    Dim matterStatus As String, retryCount As Integer
    Dim summaryWorkbook As Workbook, openSummarySheet As Worksheet, closedSummarySheet As Worksheet
    Dim matterRows As Range, outputRow As Long, currentRow As Long, lastMatchRow As Long
    Dim foundMatch As Boolean

    ' This module creates a file summarizing the Trust account balance for each matter, sorted by Open & Closed matters
    ' The summary does will not matters with no transaction history because those matters have no data in the Clio report
    
    ' Paths for saving files
    folderPath = "/Users/katielannin/desktop/California Center for Nonprofit Law/Ledgers/generated ledgers"
    openFolderPath = folderPath & Application.PathSeparator & "OPEN"
    closedFolderPath = folderPath & Application.PathSeparator & "CLOSED"

    ' Set source workbook and sheets
    Set sourceWorkbook = ThisWorkbook
    Set sourceSheet = sourceWorkbook.Sheets("Trust Ledger Report")
    Set statusSheet = sourceWorkbook.Sheets("Matter Report")

    ' Initialize collections
    Set uniqueMatters = New Collection
    Set statusCollection = New Collection

    ' Populate the status collection (Matter Number as key, Status as value)
    With statusSheet
        For Each matterCell In .Range("C2:C" & .Cells(.Rows.Count, "C").End(xlUp).Row)
            On Error Resume Next
            statusCollection.Add matterCell.Offset(0, 2).Value, CStr(matterCell.Value) ' Status is in column E
            On Error GoTo 0
        Next matterCell
    End With

    ' Populate the unique matters collection
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "C").End(xlUp).Row
    On Error Resume Next
    For Each matterCell In sourceSheet.Range("C2:C" & lastRow)
        uniqueMatters.Add matterCell.Value, CStr(matterCell.Value)
    Next matterCell
    On Error GoTo 0

    ' Create the summary workbook
    Set summaryWorkbook = Workbooks.Add
    Set openSummarySheet = summaryWorkbook.Sheets(1)
    openSummarySheet.Name = "OPEN"
    openSummarySheet.Range("A1:B1").Value = Array("Matter Number", "Balance")
    openSummaryRow = 2

    Set closedSummarySheet = summaryWorkbook.Sheets.Add
    closedSummarySheet.Name = "CLOSED"
    closedSummarySheet.Range("A1:B1").Value = Array("Matter Number", "Balance")
    closedSummaryRow = 2

    ' Loop through unique matters
    For Each matterName In uniqueMatters
        lastBalance = 0
        foundMatch = False

        ' Get the status for the current matter
        On Error Resume Next
        matterStatus = statusCollection(matterName)
        On Error GoTo 0

        If matterStatus = "" Then
            MsgBox "Status for matter '" & matterName & "' not found. Skipping..."
            GoTo NextMatter
        End If

        ' Calculate the final balance for the current matter
        For currentRow = 2 To lastRow
            If sourceSheet.Cells(currentRow, 3).Value = matterName Then
                lastBalance = sourceSheet.Cells(currentRow, 14).Value
                foundMatch = True
            End If
        Next currentRow

        If Not foundMatch Then
            MsgBox "No transactions found for matter '" & matterName & "'. Skipping..."
            GoTo NextMatter
        End If

        ' Add data to the summary workbook
        If matterStatus Like "Open*" Then
            openSummarySheet.Cells(openSummaryRow, 1).Value = matterName
            openSummarySheet.Cells(openSummaryRow, 2).Value = lastBalance
            openSummaryRow = openSummaryRow + 1
        ElseIf matterStatus Like "Closed*" Then
            closedSummarySheet.Cells(closedSummaryRow, 1).Value = matterName
            closedSummarySheet.Cells(closedSummaryRow, 2).Value = lastBalance
            closedSummaryRow = closedSummaryRow + 1
        End If

NextMatter:
    Next matterName

    ' Format summary sheets
    openSummarySheet.Columns("B").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    closedSummarySheet.Columns("B").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    MsgBox "Summary workbook created!"
End Sub


