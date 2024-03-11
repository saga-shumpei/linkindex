Attribute VB_Name = "Module1"
Sub NameLinkLog()
Attribute NameLinkLog.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' Call "NameLink" and record it in the history book.
'
'
    Dim activeName As String
    Dim strPath As String
    Dim fullPath As String
    Dim activeNameRelative As String
    Dim LinkName As String
    Dim wb As Workbook
    Dim LogFilePath As String
    Dim logWb  As Workbook
    Dim logWs As Worksheet
    Dim LastRow As Long
    Dim PathNam As String, FileNam As String, pos As Long, pathLen As Long
       
    activeName = ActiveWorkbook.Name
    strPath = ActiveWorkbook.Path
    fullPath = strPath & "\" & activeName   'Full path of the book being edited

    Call NameLink(LinkName)   'Naming cells
    
    'to retrieve the name of the log file (The log file is set in a cell named "LogFilePath" in this book.)
    Set wb = Workbooks("PERSONAL.XLSB")
    LogFilePath = wb.Sheets("settings").Range("LogFilePath").Value
    'If no log file has been set, the process ends here.
    If IsEmpty(LogFilePath) Then
        MsgBox "LogFilePath is blank. Stopping macro."
        Exit Sub
    End If
    
    'Create a relative path between the log file and the book being edited
    pos = InStrRev(LogFilePath, "\")
    PathNam = Left(LogFilePath, pos)
    pathLen = Len(PathNam)              'Length of the path part (before the file name)
    FileNam = Mid(LogFilePath, pos + 1) 'File name of log file
    activeNameRelative = "." & Mid(fullPath, pathLen)   'Relative path between log file and editing book
    
    'Open the log file and search for the last entry. At this time, 1 is returned for both 0 and 1 descriptions.
    Workbooks.Open FileName:=LogFilePath
    'MyFile = Mid$(LogFilePath, InStrRev(LogFilePath, "\") + 1)
    Set logWb = Workbooks(FileNam)
    Set logWs = logWb.Sheets("latest")
    LastRow = logWs.Cells(Rows.Count, 1).End(xlUp).Row
    'If it is not 0, add 1 to Row to make it a line where you can start writing.
    If Not IsEmpty(logWs.Cells(LastRow, 1)) Then
        LastRow = LastRow + 1
    End If
    
    logWs.Cells(LastRow, 1).Value = Now                 'Write the time in the first column
    logWs.Cells(LastRow, 3).Value = activeNameRelative  'Write the relative path in the third column
    'Create a hyperlink with a relative path
    logWs.Hyperlinks.Add Anchor:=logWs.Cells(LastRow, 4), Address:=activeNameRelative, SubAddress:=LinkName, TextToDisplay:=LinkName
    'save log
    ActiveWorkbook.Save
    'Go to the hyperlink you just created
    logWs.Cells(LastRow, 4).Hyperlinks(1).Follow
    'Save this too
    ActiveWorkbook.Save

End Sub

Function NameLink(ByRef retLinkName As String)
'
' Gives the cell where the cursor is currently located the same name as its contents. If the cell is blank, create a name from the TimeStamp.
' Let the user select any cell in the index sheet and create a link from there to the cell named above.
'
'
    Dim sheetName As String
    Dim cellAdr As String
    Dim LinkName As String
    
    sheetName = ActiveSheet.Name
    LinkName = ActiveCell.Value
    
    If LinkName = "" Then
        LinkName = "_" & Replace(Replace(Replace(Now, "/", "_"), " ", "_"), ":", "_")
        ActiveCell.Value = LinkName
    End If
    
    cellAdr = Selection.Address(, , xlR1C1)

    ActiveWorkbook.Names.Add Name:=LinkName, RefersToR1C1:="=" & sheetName & "!" & cellAdr
    ActiveWorkbook.Names(LinkName).Comment = ""
    
    Selection.Font.Bold = True
    Sheets("index").Select
    
    Set Rng = Application.InputBox(Prompt:="Please select a cell.", Type:=8)
    Rng.Select
    
    ActiveCell.Value = LinkName
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=LinkName, TextToDisplay:=LinkName
    retLinkName = LinkName

End Function


