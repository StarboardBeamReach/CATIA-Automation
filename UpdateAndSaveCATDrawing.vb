
''' <summary>
''' 
''' This program automates the two common tasks associated with engineering drawings,
''' the first task is creating a revision history block to track changes,
''' the second task is to save the file out as a pdf.
''' 
''' The user inputs the data using prompts and the table is automatically
''' inserted into the background view of the first page with standardized
''' location, features, and text.
''' 
''' This program requires a .CATDrawing be the open and active document
''' in a CATIA session. It was written for CATIA V5R26, but can work
''' with any future versions of CATIA.
''' 
''' </summary>

Sub CATMain()

    On Error Resume Next

    Dim Message1, Message2, Message3, Message4, Message5
    Dim Title, Default
    Dim DwgFolder, Rev, RevDesc, RevDate, RevApprover As String

    Message1 = "Enter drawing revision (ex: -)"
    Message2 = "Enter revision description (ex: RELEASED or INCORPORATED ECN 1234)"
    Message3 = "Enter the drawing release date in YY/MM/DD format (ex: 21/10/14)"
    Message4 = "Enter the approver for the drawing in F. LAST format (ex: J. DOE)"
    Message5 = "Enter folder to save pdf (ex: C:\CATIA\)"

    Title = "InputBox"
    Default = "l"

    Rev = InputBox(Message1, Title, Default)
    RevDesc = InputBox(Message2, Title, Default)
    RevDate = InputBox(Message3, Title, Default)
    RevApprover = InputBox(Message4, Title, Default)
    DwgFolder = InputBox(Message5, Title, Default)

    Call AddRevBlock(Rev, RevDesc, RevDate, RevApprover)
    Call PrintToPDF(DwgFolder)

End Sub


Sub AddRevBlock(newRev As String, newDesc As String, newDate As String, newApprover As String)

    Dim oDrawing As DrawingDocument
    Set oDrawing = CATIA.ActiveDocument

    Dim oSheets As DrawingSheets
    Set oSheets = oDrawing.Sheets

    Dim oSheet As DrawingSheet
    Set oSheet = oSheets.Item(1) 'rev history block is on the first page

    Dim oView As DrawingView
    Set oView = oSheets.Views.Item(2) 'rev history block is in the background view, which is always the 2nd view of any page

    Dim xPos, yPos, numRows, numCols, rowHeight, colWidth As Integer
    Set xPos = 50
    Set yPos = 50
    Set numRows = 2
    Set numCols = 4
    Set rowHeight = 20
    Set colWidth = 50

    Dim table As DrawingTable
    Set table = oView.Tables.Add(xPos, yPos, numRows, numCols, rowHeight, colWidth)

    table.SetCellString 1, 1 "REV"
    table.SetCellString 2, 1 newRev

    table.SetCellString 1, 2 "DESCRIPTION"
    table.SetCellString 2, 2 newDesc

    table.SetCellString 1, 3 "DATE"
    table.SetCellString 2, 3 newDate

    table.SetCellString 1, 4 "APPROVER"
    table.SetCellString 2, 4 newApprover

End Sub

Sub PrintToPDF(SaveFolder As String)

    Dim oDrawing As DrawingDocument
    Set oDrawing = CATIA.ActiveDocument

    Dim SaveFile As String
    Set SaveFile = SaveFolder & oDrawing.Name & ".pdf"
    
    oDrawing.ExportData SaveFile, "pdf"

End Sub
