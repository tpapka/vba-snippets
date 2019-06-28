''--- IMPORTING CLASSES/MODULE FILES
Private VBComp As VBComponent

Sub ImportOD()
    ThisWorkbook.VBProject.VBComponents.Import "G:\z-BPA\ImportClass\OpenThisClass.cls"
    OpenOD
    Set VBComp = ThisWorkbook.VBProject.VBComponents("OpenThisClass")
    ThisWorkbook.VBProject.VBComponents.Remove VBComp
End Sub

Sub OpenOD()
    Dim o As New OpenOnDemand
    o.OpenOnDemand
    Set o = Nothing
End Sub

''- OpenThisClass Class
Private OD As Variant
Private Path As String

Sub OpenOnDemand()
    Path = "C:\Program Files\IBM\OnDemand Clients\V10.1\bin\arsgui.exe"
    OD = Shell(Path, vbNormalFocus)
End Sub



''--- HOW TO GET LAST COLUMNs NUMBER AND NAME ---
Sub ColNumberName()
LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column    'this will get you column number
Debug.Print Split(Cells(, LastColumn).Address, "$")(1)      'this will get you column name
For i = 1 To LastColumn
    Debug.Print i & ". " & Split(Cells(, i).Address, "$")(1)
Next
End Sub



''--- FUNCTION TO CHECK IF AN ELEMENT IS IN AN ARRAY ALREADY
Function IsInArray(SearchFor As String, MyArray As Variant) As Boolean
    Dim st As String
    st = "$" & Join(MyArray, "$") & "$"
    IsInArray = InStr(st, "$" & SearchFor & "$") > 0
End Function



''--- CHECK IF ARRAY IS EMPTY
Function IsArrayEmpty(AnArray As Variant) As Boolean
    On Error GoTo IS_EMPTY
    If (UBound(AnArray) >= 0) Then Exit Function
IS_EMPTY:
    IsArrayEmpty = True
End Function



''--- LAST ROW & COLUMN
LastRow = Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = Sheets("Sheet1").Cells(1, Columns.Count).End(xlToLeft).Column



''--- SELECT CASE
Dim DispType as String
Select Case Range("A1").Value
    Case "01"
        DispType = "One"
    Case "04"
        DispType = "Four"
    Case "06"
        DispType = "Six"
End Select



''--- GET WORKBOOK NAME, PATH. ADD, OPEN NEW WORKBOOK
MainWkb = ActiveWorkbook.Name
CurDir = ActiveWorkbook.Path & "\"
Workbooks.Add
Workbooks.Open
ActiveWorkbook.SaveAs "Test"
FileName = Dir(CurDir & "0*.xl??")



''--- POPULATE DYNAMIC ARRAY
Dim ProjSheets()
PSCount = 0
For c = 1 To Worksheets.Count
    If Worksheets(c).Name <> "Test" Then
        ReDim Preserve ProjSheets(0 To PSCount)
        ProjSheets(PSCount) = Worksheets(c).Name
        PSCount = PSCount + 1
    End If
Next



''--- GET THE LEFT/RIGHT SIDE OF THE STRING
Sub SplitString()
    Dim s As String
    Dim MinArray() As String
    Dim k As Integer

    s = "HE-LL-OO"

    MinArray = Split(s, "-")
    Debug.Print Split(s, "-")(0) ' Left
    Debug.Print Split(s, "-")(1) ' Middle
    Debug.Print Split(s, "-")(2) ' Right
    For k = LBound(MinArray) To UBound(MinArray)
        Debug.Print (MinArray(k))
    Next
    Debug.Print UBound(MinArray) + 1
End Sub



''--- REMOVE UNWANTED CHARACTERS FROM A STRING
Sub RemoveCharacters()
    Dim RemoveChr()
    Dim Character as Variant
    RemoveChr = Array("/", "\", ":", "?", "<", ">", "|", "&", "%", "*", "{", "}", "[", "]", "!")
    MyString = "Hello \/ World!"
    For Each Character In RemoveChr
        MyString = Replace(MyString, Character, " ")
    Next Character
End Sub



''--- ACTIVATE SPECIFIC WORKBOOK
Sub ActivateWkb()
    Workbooks("Book1").Activate
    Workbooks("Book2").Activate
End Sub



''--- GET COLUMN NUMBER FROM A NAME AND VICE VERSA
Sub Sample()
    Dim ColName As String
    Dim ColNo As Integer
    ColName = "C"
    Debug.Print Trim(Range(ColName & 1).Column)
    ColNo = 3
    Debug.Print Split(Cells(, ColNo).Address, "$")(1)
End Sub



''--- FUNCTION TO CHECK IF A FILE IN A FOLDER EXISTS
Function FileOrDirExists(PathName As String) As Boolean
    Dim iTemp As Integer
    On Error Resume Next
    iTemp = GetAttr(PathName)
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
    On Error GoTo 0
End Function

Function FileOrDirExists(PathName As String) As Boolean
    If Dir(PathName) <> "" Then
        FileOrDirExists = True
    Else
        FileOrDirExists = False
    End If
End Function



''--- EXPORT SHEET TO A PDF. IF A FILE WITH THE SAME NAME EXISTS THEN ADD A FILE
''--- NUMBER AT THE END OF THE FILENAME
Sub ExportToPDF()

Dim FileNumber As Integer
Dim FName As String, Path As String
Dim sh As Worksheet

Set sh = Sheets("Sheet1")

Path = Sheets("Script").Range("A1").Value

If Path = "" Then
    MsgBox "Missing Path To A Folder!", _
        vbExclamation, "Missing Path To A Folder"
    Exit Sub
End If

If Left(Path, 1) <> "\" Then Path = Path & "\"

FileNumber = 2
FName = "Report - " & Replace(CStr(Date), "/", "-")
LastRow = sh.Cells(Rows.Count, "A").End(xlUp).Row

sh.Activate
sh.Range("A1:F" & LastRow).WrapText = False
sh.Cells.EntireColumn.AutoFit
sh.PageSetup.Orientation = xlLandscape
sh.PageSetup.Zoom = False
sh.PageSetup.FitToPagesWide = 1
sh.PageSetup.TopMargin = 5
sh.PageSetup.LeftMargin = 2
sh.PageSetup.RightMargin = 2

If FileOrDirExists(Path & FName & ".pdf") Then
    Do While FileOrDirExists(Path & FName & "(" & CStr(FileNumber) & ")" & ".pdf")
        FileNumber = FileNumber + 1
    Loop
    sh.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Path & FName & "(" & CStr(FileNumber) & ")" & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    FileNumber = FileNumber + 1
Else
    sh.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Path & FName & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End If

End Sub



''--- GET ALL COLUMN HEADER NAMES
Sub ColNumberName()
LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column    'this will get you column number
Debug.Print Split(Cells(, LastColumn).Address, "$")(1)      'this will get you column name

j = 2
For i = 1 To LastColumn
    Range("A" & j).Value = "sheet['" & Split(Cells(, i).Address, "$")(1) & "1" & "']=" & "'" & Cells(1, j) & "'"
    j = j + 1
Next

End Sub



''--- COPY FILE TO DIFFERENT FOLDER
Sub CopyingAFile()

Dim FSO
Dim FileName As String
Dim SourceFolder As String
Dim DestinationFolder As String

FileName = "powershell.exe" 'File that needs to be copied
SourceFolder = "C:\WINDOWS\system32\WindowsPowerShell\v1.0\"
DestinationFolder = "C:\Marko\"

Set FSO = CreateObject("Scripting.FileSystemObject")

If Not FSO.FileExists(DestinationFolder & FileName) Then
    FSO.CopyFile (SourceFolder & FileName), DestinationFolder, True
End If
    
End Sub



''--- GET THE FILE IN A FOLDER USING DIR AND ENVIRON FUNCTIONS
Sub EnviFile()
    OnDemandPath = Dir(Environ("AppData") & "\MBI\ODmnd Clint\DATA\" & "*.txt")
    Debug.Print OnDemandPath
End Sub



''--- LOOP THROUGH ALL THE "ENVIRON" OPTIONS
Sub LoopEnviron()
Dim i As Integer
    Dim stEnviron As String
    For i = 1 To 100
        ' get the environment variable
        stEnviron = Environ(i)
        ' see if there is a variable set
        If Len(stEnviron) > 0 Then
            Debug.Print i, Environ(i)
        Else
            Exit For
        End If
    Next
End Sub



''--- REMOVES ALL NONPRINTABLE CHARACTERS FROM A RANGE
Sub RemoveNonPrintables()
LRow = Sheets("Weekly").Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = Sheets("Weekly").Cells(1, Columns.Count).End(xlToLeft).Column
LColumn = Split(Cells(, LastColumn).Address, "$")(1)

Set MyRange = Sheets("Weekly").Range("A2:" & LColumn & LRow)
For Each cell In MyRange
    cell.Value = Application.WorksheetFunction.Clean(cell.Value)
Next
End Sub



''--- OPEN OTHER PROGRAM USING VBA, GET THE PATH USING POWERSHELL
In PowerShell: (gps -name ProcessName).Path

Sub OpenProgram()
Dim x as Variant
Dim Path as String

Path = ""C:\Program Files (x86)\IBM\OnDemand Clients\V9.5\bin\arsgui32.exe"
x = Shell(Path, vbNormalFocus)



''--- READ A TXT FILE AND PASTE ITS CONTENT INTO AN EXCEL
'(First version gets rid off of unprintable characters)
Sub ImportTxtFile1()
    OnDemandFolder = Environ("AppData") & "\IBM\OnDemand Client\DATA\"
    ODFile = Dir(OnDemandFolder & "*.A32")

    LineNum = 1
    Open OnDemandFolder & ODFile For Input As #1
    Do Until EOF(1)
        Line Input #1, TextLine
        TextLine = Replace(TextLine, Chr(0), " ")
        Range("A" & LineNum).Value = TextLine
        LineNum = LineNum + 1
    Loop
    Close #1
End Sub


Sub ImportTxtFile2()
    MyFolder = Environ("AppData") & "\IBM\OnDemand Client\DATA\"
    StrFile = Dir(MyFolder & "*.A32")

    textInput = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(MyFolder & StrFile, 1).ReadAll, vbCrLf)
    Sheets(Worksheets.Count).Activate
        
    With Range("A" & Rows.Count).End(xlUp).Offset(IIf(Cells(1, 1) = vbNullString, 0, 1), 0)
        .Resize(UBound(textInput) + 1, 1).Value = WorksheetFunction.Transpose(textInput)
    End With
End Sub



''--- SEARCH BOTTOM UP
Sub FindFromTheBottom()
    Set a = Range("A:A").Find("Hello", after:=Cells(1, 1), searchdirection:=xlPrevious)
    Set ff = Range("A1:B17")
    NumOfBlankRows = ff.Rows.Count + 1
    Rows(Range(a.Address).Row & ":" & NumOfBlankRows).Insert Shift:=xlShiftDown
    Debug.Print Range(a.Address).Row   
End Sub



''--- DELETE/REMOVE SPECIFIC SHEETS
Sub RemoveSheets()
    Dim ws As Worksheet
    NumOfSheets = Sheets.Count
    If NumOfSheets > 1 Then
        For Each ws In Worksheets
            If ws.Name <> "Script" Then
                ws.Delete
            End If
        Next
    End If
End Sub



''--- CALCULATE EXACT AGE
Sub CalculateExactAge()
    Dim s As Long
    Dim ApplicationDate As Date, BDate As Date
    Dim ExactAge As Integer
    
    s = Range("C5").Value
    ApplicationDate = DateSerial(Left(s, 4), Mid(s, 5, 2), Right(s, 2))

    s = Range("X5").Value
    BDate = DateSerial(Left(s, 4), Mid(s, 5, 2), Right(s, 2))

    ExactAge = DateDiff("yyyy", BDate, ApplicationDate)
    If ApplicationDate < DateSerial(Year(ApplicationDate), Month(BDate), Day(BDate)) Then
        ExactAge = ExactAge - 1
    End If
    Debug.Print ExactAge
End Sub



''--- DELETE CELL VALUE AND MOVE EVERYTHING IN JUST THAT ONE COLUMN UP
For i = 1 To 10
    If ActiveSheet.Cells(i, 1).Value = "MAX" Then 
        ActiveSheet.Cells(i, 1).Resize(1, 1).Delete
    End If
Next i



''--- REORGANIZE COLUMN ORDER
Sub Reorder_Columns()
Dim ColumnOrder As Variant, ndx As Integer
Dim Found As Range, counter As Integer

ColumnOrder = Array("H6", "H2", "H1", "H4", "H5", "H3")
counter = 1

Application.ScreenUpdating = False
   
For ndx = LBound(ColumnOrder) To UBound(ColumnOrder)
    Set Found = Rows("1:1").Find(ColumnOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, MatchCase:=False)
    If Not Found Is Nothing Then
        If Found.Column <> counter Then
            Found.EntireColumn.Cut
            Columns(counter).Insert Shift:=xlToRight
            Application.CutCopyMode = False
        End If
    counter = counter + 1
    End If
Next ndx
Application.ScreenUpdating = True
End Sub



''--- VBA VLOOKUP
Sub VLookupSuxx()

LastRow = Cells(Rows.Count, "E").End(xlUp).Row

For i = 1 To LastRow
    For k = 1 To LastRow
        If CStr(Range("E" & i).Value) = CStr(Range("A" & k).Value) Then
            Range("F" & i).Value = "ok"
'            Exit For
        End If
    Next
Next

End Sub



''--- VBA LOOP OVER ALL FILES IN A FOLDER AND SUBFOLDERS
Option Explicit
Private SubFolders As Folders
Private MyFolder As String, MyFile As String
Private fso As FileSystemObject, MainFolder As Object
Private sf

Sub FieldConversion()

Application.ScreenUpdating = False

MyFolder = Sheets("Reformat").Range("A1").Value
If MyFolder = "" Then
    MsgBox "You forgot to enter path into cell 'A1'", vbExclamation, "Missing Folder Path"
    Exit Sub
End If

If Right(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"

Set fso = New FileSystemObject
Set MainFolder = fso.GetFolder(MyFolder)
Set SubFolders = MainFolder.SubFolders

If SubFolders.Count > 0 Then
    For Each sf In SubFolders
        MyFolder = sf & "\"
        MyFile = Dir(sf & "\" & "*.xlsx")
        ProcessFiles
    Next
End If

Application.ScreenUpdating = True

End Sub


Sub ProcessFiles()

Application.ScreenUpdating = False

Do While Len(MyFile) > 0
    Workbooks.Open (MyFolder & MyFile)
    ActiveWorkbook.Sheets("PDPL").Unprotect ("mypassword")
    Range("D21").NumberFormat = "0.00"
    Range("D22").NumberFormat = "$#,##0.00"
    ActiveWorkbook.Sheets("PDPL").Protect ("mypassword")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    MyFile = Dir
Loop

Application.ScreenUpdating = True

End Sub



''--- MYSEARCH FUNCTION
Function MySearchFunction(FindWhat As String)

Set MySearch = Cells.Find(What:=FindWhat, After:=ActiveCell, _
    LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
End Function



''--- VBA SEARCH WHOLE VALUES
Function MySearchFunction(FindWhat As String, SheetName As String)
    
    Set MySearch = Sheets(SheetName).Cells.Find(What:=FindWhat, After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        
    If MySearch Is Nothing Then
        msf = "zero"
    Else
        msf = MySearch.Address
    End If
        
End Function



''--- CREATING EMPTY TXT FILE (hidden)
'' Check Microsoft Scripting Runtime in References
Sub CreateTxtFile()
Dim fso As New Scripting.FileSystemObject
Dim ofile As Scripting.File

CookieName = Environ("LOCALAPPDATA") & "\" & "12375458888788" & ".txt"
Open CookieName For Output As #1: Close #1

Set ofile = fso.GetFile(Environ("LOCALAPPDATA") & "\" & "12375458888788" & ".txt")

ofile.Attributes = Hidden
End Sub



''--- Collecting Values From A Column and adding it to an Array
''--- Searching for all of them in every cell that hold any value (UsedRange)
Sub tt()

Dim Domains()
Dim sh As Range

LastRow = Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
Count = 0

Set sh = Sheets("user_profiles").UsedRange

For r = 1 To LastRow
    ReDim Preserve Domains(0 To Count)
    Domains(Count) = Sheets("Sheet1").Range("A" & r).Value
    Count = Count + 1
Next

For Each c In sh.Cells
    If IsInArray(c.Value, Domains) Then
        c.Interior.ColorIndex = 6
    End If
Next

End Sub



''--- CLOSING WEBSITE POPUPS
Shell "wscript.exe ""G:\z-BPA\ImportClass\ClosePopUp.vbs"""
'- ClosePopUp.vbs
set wshShell = CreateObject("wscript.shell")

Do
    ret = wshShell.AppActivate("Message from webpage")
Loop Until ret = True

WScript.Sleep 500

ret = wshShell.AppActivate("Message from webpage")
If ret = True Then
    ret = wshshell.AppActivate("Message from webpage")
    WScript.Sleep 10
    wshShell.SendKeys "{enter}"
End If

