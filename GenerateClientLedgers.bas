Attribute VB_Name = "mattername"
Sub GenMatterLedgers5()
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim statusSheet As Worksheet
    Dim filterSheet As Worksheet
    Dim folderPath As String
    Dim openFolderPath As String
    Dim closedFolderPath As String
    Dim newFilePath As String
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim uniqueMatters As Collection
    Dim matterCell As Range
    Dim filterCell As Range
    Dim matterName As Variant
    Dim lastRow As Long
    Dim statusCollection As Collection
    Dim statusRow As Range
    Dim matterStatus As String
    Dim fileCount As Long
    Dim filterLastRow As Long

    ' Initialize file counter
    fileCount = 0

    ' Paths for saving files
    folderPath = "/Users/kathrynvictoria/desktop/CCLC/Macros/"
    openFolderPath = folderPath & "/OPEN"
    closedFolderPath = folderPath & "/CLOSED"

    ' Create OPEN and CLOSED folders
    On Error Resume Next
    MkDir openFolderPath
    MkDir closedFolderPath
    On Error GoTo 0

    ' Set source workbook and sheets
    Set sourceWorkbook = ThisWorkbook
    Set sourceSheet = sourceWorkbook.Sheets("Trust Ledger Report")
    Set statusSheet = sourceWorkbook.Sheets("Matter Report")
    Set filterSheet = sourceWorkbook.Sheets("List of Matters to Run On")

    ' Initialize collections
    Set uniqueMatters = New Collection
    Set statusCollection = New Collection

    ' Populate the status collection (Matter Number as key, Status as value)
    With statusSheet
        For Each statusRow In .Range("C2:C" & .Cells(.Rows.Count, "C").End(xlUp).Row)
            On Error Resume Next
            statusCollection.Add statusRow.Offset(0, 2).Value, CStr(statusRow.Value) ' Status is in column E
            On Error GoTo 0
        Next statusRow
    End With

    ' Pull matter numbers from the "List of Matters to Run On" tab
    filterLastRow = filterSheet.Cells(filterSheet.Rows.Count, "A").End(xlUp).Row

    On Error Resume Next
    For Each filterCell In filterSheet.Range("A2:A" & filterLastRow)
        If Trim(filterCell.Value) <> "" Then
            uniqueMatters.Add Trim(filterCell.Value), CStr(Trim(filterCell.Value))
        End If
    Next filterCell
    On Error GoTo 0

    ' Find the last row in the source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through unique matters and create files
    For Each matterName In uniqueMatters
        Dim outputRow As Long
        Dim finalBalance As Double
        outputRow = 2

        ' Get the status for the current matter from the collection
        On Error Resume Next
        matterStatus = statusCollection(matterName)
        On Error GoTo 0

        ' Sanitize the matter name
        matterName = Replace(matterName, Chr(160), " ") ' Replace non-breaking space
        matterName = Replace(matterName, "/", "_")
        matterName = Replace(matterName, "\", "_")
        matterName = Replace(matterName, ":", "-")
        matterName = Replace(matterName, "*", "")
        matterName = Replace(matterName, "?", "")
        matterName = Replace(matterName, """", "")
        matterName = Replace(matterName, "<", "")
        matterName = Replace(matterName, ">", "")
        matterName = Replace(matterName, "|", "")

        ' Determine file path based on status
        If matterStatus Like "Open*" Then
            newFilePath = openFolderPath & "/" & matterName & ".xlsx"
        ElseIf matterStatus Like "Closed*" Then
            newFilePath = closedFolderPath & "/" & matterName & ".xlsx"
        Else
            MsgBox "Status for " & matterName & " not found in the Matter Report."
            GoTo NextMatter
        End If

        ' Create new workbook and sheet
        Set newWorkbook = Workbooks.Add
        Set newSheet = newWorkbook.Sheets(1)

        ' Add headers
        newSheet.Range("A1:F1").Value = Array("Date", "Action", "Invoice #", "Check #", "Transaction Amount", "Total")

        ' Write data to the new sheet
        For Each matterCell In sourceSheet.Range("C2:C" & lastRow)
            If Trim(matterCell.Value) = matterName Then
                With newSheet
                    .Cells(outputRow, 1).Value = matterCell.Offset(0, 3).Value ' Date
                    .Cells(outputRow, 2).Value = matterCell.Offset(0, 7).Value ' Action
                    .Cells(outputRow, 3).Value = matterCell.Offset(0, 6).Value ' Invoice #
                    .Cells(outputRow, 4).Value = matterCell.Offset(0, 8).Value ' Check #
                    .Cells(outputRow, 5).Value = matterCell.Offset(0, 10).Value - matterCell.Offset(0, 9).Value ' Transaction Amount
                    .Cells(outputRow, 6).Value = matterCell.Offset(0, 11).Value ' Total
                End With
                finalBalance = matterCell.Offset(0, 11).Value ' Final balance is in column N
                outputRow = outputRow + 1
            End If
        Next matterCell

        ' Apply table formatting
        Dim tableObject As ListObject
        Set tableObject = newSheet.ListObjects.Add(xlSrcRange, newSheet.Range("A1").Resize(outputRow - 1, 6), , xlYes)
        tableObject.Name = "MatterTable"

        ' Format columns E and F as Accounting
        With newSheet.Columns("E:F")
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End With

        ' Save and close the new workbook
        On Error Resume Next
        Kill newFilePath ' Delete if already exists
        On Error GoTo 0
        newWorkbook.SaveAs FileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
        newWorkbook.Close SaveChanges:=False
        fileCount = fileCount + 1

NextMatter:
    Next matterName

    ' Display a success message with the file count
    MsgBox "Great job! " & fileCount & " files created successfully.", vbInformation
End Sub

