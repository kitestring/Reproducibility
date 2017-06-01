VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrecisionTesting 
   Caption         =   "Data Miner"
   ClientHeight    =   4035
   ClientLeft      =   -4950
   ClientTop       =   -4620
   ClientWidth     =   6255
   OleObjectBlob   =   "frmPrecisionTesting.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrecisionTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdRun_Click()
'**Note** All arrays are base 1

Dim Analytes() As String
Dim TotalAnalytesInList As Integer
Dim NoOfAnalytesSelected As Integer
Dim analyte As Integer

Dim Fields() As String
Dim TotalFieldsInList As Integer
Dim NoOfFieldsSelected As Integer
Dim field As Integer

Dim AnalyteData() As String

Dim ExperimentDate As String
Dim ExperimentTitle As String
Dim MS_SN As String
Dim MS_Type As String
Dim Filament As String
Dim GC_SN As String
Dim GC_Type As String
Dim SoftwareVersion As String

Dim SourceCodeWkbk As Workbook
Dim DataAnalysisWkbk As Workbook
Dim RawDataWkbk As Workbook

Dim csvFile As Variant
Dim NoOfcsvFiles As Integer
Dim file As Integer

Const AnalyteMissingReturnValue As String = "Analyte Not Found"
Const FieldMissingReturnValue As String = "Field Missing"
Const NullValueReturnValue As String = "Null"

Dim Row As Integer
Dim WorkingRow As Integer
Dim EvaluationRow As Integer
Dim Column As Integer
Dim GraphDataTopLink As Integer
Dim GraphDataBottomLink As Integer

Const AnalyteTemplateSheetStartingRow As Integer = 2
Const AnalyteTemplateSheetStartingColumn As Integer = 3
Const AnalyteNameHeader As String = "Name"
Const ColumnToFreezeOnSummarySheet As String = "C"
Const NameOfSummarySheet As String = "Summary"
Const ColumnLetterForInjectionNumber As String = "A"
Dim c As String

Dim AverageSummaryLink As String
Dim SDSummaryLink As String
Dim RSDSummaryLink As String
Dim GraphDataLink As String
Dim DataRow As Integer
Dim AverageColumn As Integer
Dim SDColumn As Integer

Dim RSDColumn As Integer
Dim RSDFlagRow As Integer
Const RSDFlagValues As Single = 0.075
Dim RSDFlagLink() As String
Dim RSDFlagColumn As Integer

Dim NotFoundColumnNumber As Integer
Dim NotFoundEquation As String

Dim GraphPrefix As String
Dim L(1) As String
Dim AxisPaths() As String
Dim GraphAnalyteSummary() As Boolean
Dim GraphingElement As Integer
Dim Graph_LeftLocation As Long
Dim Graph_TopLocation As Long
Dim Graph_Location_Intermediate As Double
Dim GraphCounter As Integer
Dim GraphTopOffset As Double

Dim FileSaveName As Variant
Dim DefaultFileSavePath As String

'Define SourceCodeWorkbook, turn off screen updating, & hide frmPrecisionTesting
    Set SourceCodeWkbk = ActiveWorkbook
    Call ProgressReport(SourceCodeWkbk, SourceCodeWkbk, "Extracting data from csv files")
    
'Grab Instrumentation information
    ExperimentDate = txtDate.Value
    ExperimentTitle = txtExpTitle.Value
    MS_SN = txtMS_SN.Value
    MS_Type = cmbMSType.Value
    Filament = cmbFilament.Value
    GC_SN = txtGC_SN.Value
    GC_Type = cmbGCType.Value
    SoftwareVersion = txtSoftwareVersion.Value
    
    'GraphPrefix = ExperimentTitle & " / " & MS_Type & " " & MS_SN & " / EG " & Filament
    GraphPrefix = MS_SN & " " & ExperimentTitle
    
    
'Check for empty analyte list box
    If lstAnalytes.ListCount = 0 Then
        MsgBox "No Analytes Selected", vbCritical, "Missing Entry"
        Exit Sub
    End If
    
'Check for empty field list box
    If lstFields.ListCount = 0 Then
        MsgBox "No Fieldss Selected", vbCritical, "Missing Entry"
        Exit Sub
    End If
    
'Grab values from analyte list box and populate Analyte Array
    NoOfAnalytesSelected = 1
    TotalAnalytesInList = lstAnalytes.ListCount - 1
    ReDim Analytes(TotalAnalytesInList + 1) As String
    
    For analyte = 0 To TotalAnalytesInList
        If lstAnalytes.Selected(analyte) = True Then
            Analytes(NoOfAnalytesSelected) = lstAnalytes.List(analyte)
            NoOfAnalytesSelected = NoOfAnalytesSelected + 1
        End If
    Next analyte
    NoOfAnalytesSelected = NoOfAnalytesSelected - 1
    
'Grab values from field list box and populate Field Array
    NoOfFieldsSelected = 1
    TotalFieldsInList = lstFields.ListCount - 1
    ReDim Fields(TotalFieldsInList + 1) As String
    
    For field = 0 To TotalFieldsInList
        If lstFields.Selected(field) = True Then
            Fields(NoOfFieldsSelected) = lstFields.List(field)
            NoOfFieldsSelected = NoOfFieldsSelected + 1
        End If
    Next field
    NoOfFieldsSelected = NoOfFieldsSelected - 1
    
'Redimension RSDFlagLink array & ColumnHeaders array
    ReDim RSDFlagLink(NoOfFieldsSelected) As String

'Prompt user for csv files to be data mined
    csvFile = Application.GetOpenFilename(Title:="Select csv Files", MultiSelect:=True)
    If VarType(csvFile) = vbBoolean Then Exit Sub
    NoOfcsvFiles = UBound(csvFile)
    frmPrecisionTesting.Hide
    
'Redim AnalytedData array using: AnalyteData(No Of csv Files, No of Analytes, No of data fields)
    ReDim AnalyteData(NoOfcsvFiles, NoOfAnalytesSelected, NoOfFieldsSelected) As String

'Begin data extraction loop
    For file = 1 To NoOfcsvFiles
        Workbooks.Open Filename:=csvFile(file)
        Set RawDataWkbk = ActiveWorkbook
        'Call ProgressReport(SourceCodeWkbk, RawDataWkbk, "Extracting data from file number " & file)
        For analyte = 1 To NoOfAnalytesSelected
            If FindString(Analytes(analyte)) = False Then
                For field = 1 To NoOfFieldsSelected
                    AnalyteData(file, analyte, field) = AnalyteMissingReturnValue
                Next field
            Else
                Row = ActiveCell.Row
                For field = 1 To NoOfFieldsSelected
                    If FindString(Fields(field)) = False Then
                        AnalyteData(file, analyte, field) = FieldMissingReturnValue
                    End If
                    Cells(10, 1).Select
                    If FindString(Fields(field)) = True Then
                        Column = ActiveCell.Column
                        If Cells(Row, Column).Value = "" Then
                            AnalyteData(file, analyte, field) = NullValueReturnValue
                        Else
                            AnalyteData(file, analyte, field) = Cells(Row, Column).Value
                            If chkFieldFilter = True And MS_Type = "Pegasus HRT" Then
                            'If MS_Type = "Pegasus HRT" Then
                                Select Case Fields(field)
                                    Case "Quant Masses"
                                        AnalyteData(file, analyte, field) = MassFieldCorrection(AnalyteData(file, analyte, field), "(", "±")
                                    Case "Actual Masses"
                                        AnalyteData(file, analyte, field) = MassFieldCorrection(AnalyteData(file, analyte, field), "(", ")")
                                End Select
                            End If
                        End If
                    End If
                Next field
            End If
        Next analyte
        RawDataWkbk.Close False
    Next file

'Create and set data analysis workbook
    Workbooks.Add
    Set DataAnalysisWkbk = ActiveWorkbook
    
'***Build analysis workbook based upon fields found during data extraction
'1) Build Temp sheet
    '1a) Format template analyte worksheet
    '1b) Drop Header Names, Equations, & Conditional Formatting
    '1c) Drop in labels and RSD flag values
'2) Copy and rename Sheets("Temp") for each analyte
'3) Build summary worksheet
    '3a) Copy/rename temp sheet, clear it, reformat for RSD summary & freeze column
    '3b) Drop header names
    '3c) Drop chemical names
    '3d) Drop equations & conditional formatting
    '3e) Drop conditional formatting
    '3f) Drop injection numbers
    '3g) Format Count Not Found column
'4) Create links between from Analyte Sheet RSD Flags to summary sheet RSD flags


'1)Build Temp sheet
    '1a)Format worksheet
        Call ProgressReport(SourceCodeWkbk, DataAnalysisWkbk, "Building & Formatting Workbook")
        Call MakeAndRenameNewSheet("Sheet1", "AxisPaths")
        Call FormatAxisNamesSheet
        Sheets("Sheet1").Select
        Row = AnalyteTemplateSheetStartingRow
        Column = AnalyteTemplateSheetStartingColumn
        Call FormatWorkbook(NoOfFieldsSelected + 1, Row, Column, NoOfcsvFiles)
        Cells(Row, Column - 1).Value = AnalyteNameHeader
        Cells(Row, Column + NoOfFieldsSelected).Value = "csv file directory address"
    '1b)Drop Header Names, Equations, & Conditional Formatting
        For field = 1 To NoOfFieldsSelected
            c = DetermineColumnLetter(Column + field - 1)
            Cells(Row, Column + field - 1).Value = Fields(field)
            Call DropEquationsAndFormat(Row, c, NoOfcsvFiles)
            Call ConditionalFormatting(c, Row + NoOfcsvFiles + 4, Row + NoOfcsvFiles + 3)
        Next field
    '1c)Drop in labels and RSD flag values
        Call DropEquationTitles(NoOfFieldsSelected, Row, Column, NoOfcsvFiles)
'2)Copy and rename Sheets("Sheet1") for each analyte
    For analyte = 1 To NoOfAnalytesSelected
        'Call MakeAndRenameNewSheet("Sheet1", Left(Analytes(analyte), 30))
        Call MakeAndRenameNewSheet("Sheet1", Left(Analytes(analyte), 28) & "-" & analyte)
    Next analyte
'3)Build summary worksheet
    '3a)Copy/rename temp sheet, clear it, reformat for RSD summary & freeze column
        Call MakeAndRenameNewSheet("Sheet1", NameOfSummarySheet)
        Call ClearTempSheet("A", DetermineColumnLetter(Column + NoOfFieldsSelected + 3))
        Call FormatWorkbook(NoOfFieldsSelected * 3, Row, Column, NoOfAnalytesSelected + 1, True)
        Cells(Row, Column - 1).Value = "Chemical Name"
        Call FreezeColumn(ColumnToFreezeOnSummarySheet)
    '3b)Drop header names
        For field = 1 To NoOfFieldsSelected
            Cells(Row, Column + ((field - 1) * 3)).Value = Fields(field)
            Cells(Row + 1, Column + ((field - 1) * 3)).Value = "Average"
            Cells(Row + 1, (Column + ((field - 1) * 3)) + 1).Value = "Std. Dev."
            Cells(Row + 1, (Column + ((field - 1) * 3)) + 2).Value = "RSD"
        Next field
    '3c)Drop chemical names
        For analyte = 1 To NoOfAnalytesSelected
            Cells(Row + analyte + 1, Column - 1).Value = Left(Analytes(analyte), 18)
        Next analyte
    '3d&e) Drop equations & conditional formatting
        For analyte = 1 To NoOfAnalytesSelected
            For field = 1 To NoOfFieldsSelected
                c = DetermineColumnLetter(Column + field - 1)
                AverageSummaryLink = DetermineEquationLinks(Row, c, NoOfcsvFiles, _
                    Analytes(analyte), "Average", analyte)
                SDSummaryLink = DetermineEquationLinks(Row, c, NoOfcsvFiles, _
                    Analytes(analyte), "Standard Deviation", analyte)
                RSDSummaryLink = DetermineEquationLinks(Row, c, NoOfcsvFiles, _
                    Analytes(analyte), "Relative Standard Deviation", analyte)
                GraphDataLink = DetermineEquationLinks(Row, c, NoOfcsvFiles, _
                    Analytes(analyte), "Graph Data", analyte)
                
                DataRow = Row + analyte + 1
                AverageColumn = (Column + ((field - 1) * 3)) + 0
                SDColumn = (Column + ((field - 1) * 3)) + 1
                RSDColumn = (Column + ((field - 1) * 3)) + 2
                RSDFlagRow = Row + NoOfAnalytesSelected + 2
                GraphDataTopLink = 7 + NoOfAnalytesSelected
                GraphDataBottomLink = (7 + (NoOfAnalytesSelected * 2)) - 1
                c = DetermineColumnLetter(RSDColumn)
                
                Cells(DataRow, AverageColumn).Value = AverageSummaryLink
                Cells(DataRow, SDColumn).Value = SDSummaryLink
                Cells(DataRow, RSDColumn).Value = RSDSummaryLink
                Cells(DataRow + NoOfAnalytesSelected + 3, RSDColumn).Value = GraphDataLink
                
                Cells(DataRow + NoOfAnalytesSelected + 3, RSDColumn).Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                
                Cells(GraphDataTopLink - 1, RSDColumn).Value = "=IF(COUNTIF(" & c & GraphDataTopLink & ":" & c & GraphDataBottomLink & ",TRUE)>0,TRUE,FALSE)"
                Cells(GraphDataTopLink - 1, RSDColumn).Select
                ActiveWorkbook.Names.Add Name:="Graph_" & field, RefersToR1C1:="=Summary!R" & GraphDataTopLink - 1 & "C" & RSDColumn
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
               
                Call ConditionalFormatting(c, RSDFlagRow, DataRow)
                Cells(RSDFlagRow, RSDColumn).Value = RSDFlagValues
                Call PercentNumberFormat(Row + 2, RSDColumn, RSDFlagRow, RSDColumn)
                RSDFlagLink(field) = "='Summary'!" & c & RSDFlagRow
            Next field
        Next analyte
    '3f)Drop injection numbers
        Call FillSeriesDown(NoOfcsvFiles)
    '3g) Format Count Not Found column
        NotFoundColumnNumber = Formatting_CountNotFounds(Row, Column, NoOfAnalytesSelected, AnalyteMissingReturnValue)
        For analyte = 1 To NoOfAnalytesSelected
            NotFoundEquation = "=COUNTIF('" & Left(Analytes(analyte), 28) & "-" & analyte & "'!1:1048576," & Chr(34) & _
                AnalyteMissingReturnValue & Chr(34) & ")/" & NoOfFieldsSelected
            Cells(Row + analyte + 1, NotFoundColumnNumber).Value = NotFoundEquation
        Next analyte
'4)Create links between from Analyte Sheet RSD Flags to summary sheet RSD flags
    RSDFlagRow = Row + NoOfcsvFiles + 4
    For analyte = 1 To NoOfAnalytesSelected
        Sheets(Left(Analytes(analyte), 28) & "-" & analyte).Select
        For field = 1 To NoOfFieldsSelected
            RSDFlagColumn = Column + field - 1
            Cells(RSDFlagRow, RSDFlagColumn).Value = RSDFlagLink(field)
        Next field
        Call CenterCells
    Next analyte
    
    Erase RSDFlagLink
    
'Begin data dump loop
    Column = Column - 1
    
    For analyte = 1 To NoOfAnalytesSelected
        Call ProgressReport(SourceCodeWkbk, DataAnalysisWkbk, "Dumping " & Analytes(analyte) & " Data")
        Sheets(Left(Analytes(analyte), 28) & "-" & analyte).Select
        For file = 1 To NoOfcsvFiles
            WorkingRow = Row + file
            EvaluationRow = WorkingRow + NoOfcsvFiles + 6
            Cells(WorkingRow, Column).Value = Analytes(analyte)
            Cells(WorkingRow, Column + NoOfFieldsSelected + 1).Value = csvFile(file)
            For field = 1 To NoOfFieldsSelected
                L(0) = DetermineColumnLetter(Column + field)
                Cells(WorkingRow, Column + field).Value = AnalyteData(file, analyte, field)
                Cells(EvaluationRow, Column + field).Value = "=ISNUMBER(" & L(0) & WorkingRow & ")"
                Cells(EvaluationRow, Column + field).Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            Next field
        Next file
        Cells.Select
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
    Next analyte
    
    Erase AnalyteData
    Erase csvFile
    
'Begin graph building loop for each analyte sheet
    For analyte = 1 To NoOfAnalytesSelected
        Call ProgressReport(SourceCodeWkbk, DataAnalysisWkbk, "Graphing " & Analytes(analyte) & " data")
        Sheets(Left(Analytes(analyte), 28) & "-" & analyte).Select
        Call InsertRows
        For field = 1 To NoOfFieldsSelected
            If CreateGraph(Fields(field)) = True Then
                Call FindString(Fields(field))
                L(0) = DetermineColumnLetter(ActiveCell.Column) 'y-Axis column letter
                Call FindString("Name")
                L(1) = DetermineColumnLetter(ActiveCell.Column) 'Name column letter
                Cells(1, 1).Select
                Call GraphManager(1, NameOfSummarySheet, Left(Analytes(analyte), 28) & "-" & analyte, Left(Analytes(analyte), 28) & "-" & analyte, _
                    ColumnLetterForInjectionNumber, L(0), L(1), NoOfcsvFiles, NoOfFieldsSelected, Row + 1, 30, _
                    True, analyte, field, Fields(field), GraphPrefix, Analytes(analyte))
            End If
        Next field
    Next analyte
    
    Erase L

'***Create Graphs and finalize Summary worksheet
'1) Grab axis paths which define the name, x-axis, & y-axis paths for each line on each graph
    'a) Delete AxisPath sheet
'2) Grab GraphAnalyteSummary Boolean values to determine which fields to graph
    'a) Center & customize the width of each column
'3) Create each graph
'4) Build and populate instrument information table
'5) Create last analysis meta data worksheet (History)
'-----------------------------------------------------

    Call ProgressReport(SourceCodeWkbk, DataAnalysisWkbk, "Building Summary Graphs")

'1) Grab axis paths which define the name, x-axis, & y-axis paths for each line on each graph
    ReDim AxisPaths(NoOfFieldsSelected, NoOfAnalytesSelected, 2) As String '0 = Series Name / 1 = x-Axis / 2 = y-Axis
    Sheets("AxisPaths").Select
    
    For field = 1 To NoOfFieldsSelected
        Column = 1 + ((field - 1) * 3)
        For analyte = 1 To NoOfAnalytesSelected
            Row = analyte + 1
            For GraphingElement = 0 To 2
                AxisPaths(field, analyte, GraphingElement) = Cells(Row, Column + GraphingElement).Value
            Next GraphingElement
        Next analyte
    Next field
    
    '1a) Delete AxisPath sheet
        Application.DisplayAlerts = False
        Sheets("AxisPaths").Delete
        Application.DisplayAlerts = True
    
'2) Grab GraphAnalyteSummary Boolean values to determine which fields to graph
    ReDim GraphAnalyteSummary(NoOfFieldsSelected) As Boolean
    DataAnalysisWkbk.Activate
    Sheets(NameOfSummarySheet).Select
    
    For field = 1 To NoOfFieldsSelected
        GraphAnalyteSummary(field) = Range("Graph_" & field).Value
    Next field
    
    '2a) Center & customize the width of each column
        Call CenterCells
        Cells.EntireColumn.AutoFit

'3) Create each graph
    GraphCounter = 0
    GraphTopOffset = (NoOfAnalytesSelected + 6) * (325 / 18.7)
    GraphTopOffset = Application.WorksheetFunction.Round(GraphTopOffset, 0)
    
    For field = 1 To NoOfFieldsSelected
    
        If GraphAnalyteSummary(field) = True Then
            GraphCounter = GraphCounter + 1
            Graph_Location_Intermediate = (GraphCounter - 1) / 3
            Graph_TopLocation = GraphTopOffset + (Int(Graph_Location_Intermediate) * 520)
            Graph_Location_Intermediate = Graph_Location_Intermediate - Int(Graph_Location_Intermediate)
        
            Graph_Location_Intermediate = Application.WorksheetFunction.Round(Graph_Location_Intermediate, 2)
            
            Select Case Graph_Location_Intermediate
                Case 0
                    Graph_LeftLocation = (150 + (780 * 0))
                Case 0.33
                    Graph_LeftLocation = (150 + (780 * 1))
                Case 0.67
                    Graph_LeftLocation = (150 + (780 * 2))
            End Select
        
            Call CreateSummaryChart(Graph_LeftLocation, Graph_TopLocation, 760, 500, AxisPaths(field, 1, 0), _
                AxisPaths(field, 1, 1), AxisPaths(field, 1, 2), GraphPrefix & Chr(10) & Fields(field) & " vs. Injection No.", 1, "Injection No.", Fields(field))
            Call AddMarkersToSeries(GraphCounter, 1)
            
            For analyte = 2 To NoOfAnalytesSelected
                Call AddSeriesToGraph(analyte, AxisPaths(field, analyte, 0), AxisPaths(field, analyte, 1), AxisPaths(field, analyte, 2))
                Call AddMarkersToSeries(GraphCounter, analyte)
            Next analyte
            
        End If
            
        Call ChangeLegendFont(GraphCounter, 14)
        
    Next field
    
'4) Build and populate instrument information table
    Call FormatInstrumentalInformation
    
    Range("N4").Value = ExperimentDate
    Range("D2").Value = ExperimentTitle
    Range("F4").Value = MS_SN
    Range("F3").Value = MS_Type
    Range("F5").Value = Filament
    Range("J4").Value = GC_SN
    Range("J3").Value = GC_Type
    Range("N3").Value = SoftwareVersion
    
'5) Create last analysis meta data worksheet (History)
    SourceCodeWkbk.Activate
    Sheets("MetaData").Select
    
    Cells(3, 2).Value = NoOfAnalytesSelected
    For analyte = 1 To NoOfAnalytesSelected
        Cells(3 + analyte, 2).Value = Analytes(analyte)
    Next analyte
    
    Cells(3, 3).Value = NoOfFieldsSelected
    For field = 1 To NoOfFieldsSelected
        Cells(3 + field, 3).Value = Fields(field)
    Next field
    
    Cells(3, 5).Value = ExperimentDate
    Cells(4, 5).Value = ExperimentTitle
    Cells(5, 5).Value = MS_SN
    Cells(6, 5).Value = MS_Type
    Cells(7, 5).Value = Filament
    Cells(8, 5).Value = GC_SN
    Cells(9, 5).Value = GC_Type
    Cells(10, 5).Value = SoftwareVersion
    
    If opbYes.Value = True Then
        Sheets("Sheet1").Select
        SourceCodeWkbk.Save
    End If
    
    Cells(1, 1).Select
    
    Sheets("MetaData").Copy Before:=DataAnalysisWkbk.Sheets(1)
    Sheets("MetaData").Visible = xlSheetVeryHidden
    
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    Sheets(NameOfSummarySheet).Select
    
'Save final workbook
    'DefaultFileSavePath = DesktopAddress & MS_SN & "_" & ExperimentTitle & "_EG " & Filament
    DefaultFileSavePath = MS_SN & "_" & ExperimentTitle & "_EG " & Filament
    
    Call ProgressReport(SourceCodeWkbk, SourceCodeWkbk, "Completed Building Reproducibility Workbook")
    
SaveAs:
    FileSaveName = Application.GetSaveAsFilename(DefaultFileSavePath, "Excel Files (*.xlsx), *.xlsx", , "Title", "Button Text")
    If FileSaveName = False Then GoTo SaveAs

    DataAnalysisWkbk.SaveAs FileSaveName
    
    Unload frmPrecisionTesting

End Sub

Private Sub cmdImportAnalytesFields_Click()
Dim csvFile As Variant
Dim RawAnalytesFieldsWkbk As Workbook
Const AnalytesHeader As String = "Name"
Const HeadersRowNumber As Integer = 1
Dim WasTheStringFound As Boolean
Dim Row As Integer
Dim Column As Integer
Dim EndRow As Integer
Dim EndColumn As Integer


On Error GoTo Errors
Application.ScreenUpdating = False
WasTheStringFound = True
csvFile = True


    csvFile = Application.GetOpenFilename(Title:="Select *.csv File To Define Analytes & Fields", MultiSelect:=False)
    
    Workbooks.Open Filename:=csvFile
    Set RawAnalytesFieldsWkbk = ActiveWorkbook
    
    lstAnalytes.Clear
    lstFields.Clear
    
    WasTheStringFound = FindString(AnalytesHeader)
    If WasTheStringFound = False Then GoTo Errors
    
    Row = ActiveCell.Row
    Column = ActiveCell.Column
    Selection.End(xlDown).Select
    EndRow = ActiveCell.Row
    
    Cells(1, 1).Select
    Selection.End(xlToRight).Select
    EndColumn = ActiveCell.Column
    
    For Row = (HeadersRowNumber + 1) To EndRow
        lstAnalytes.AddItem Cells(Row, Column).Value
    Next Row
    
    Row = HeadersRowNumber
    For Column = 1 To EndColumn
        If Cells(Row, Column).Value <> AnalytesHeader Then
            lstFields.AddItem Cells(Row, Column).Value
        End If
    Next Column
    
    RawAnalytesFieldsWkbk.Close
    
Exit Sub
Errors:
If csvFile = False Then Exit Sub
If WasTheStringFound = False Then
    MsgBox AnalytesHeader & " Field Missing", vbCritical, "Invalid *.csv File"
    RawAnalytesFieldsWkbk.Close
    Exit Sub
End If
    
End Sub
    
Private Sub cmdClearAnalyteFields_Click()
lstAnalytes.Clear
lstFields.Clear
chkAllAnalytes.Value = False
chkAllFields.Value = False
End Sub

Private Sub chkAllAnalytes_Click()
Dim AnalyteCount As Integer
Dim analyte As Integer

    AnalyteCount = lstAnalytes.ListCount - 1
    For analyte = 0 To AnalyteCount
        lstAnalytes.Selected(analyte) = chkAllAnalytes.Value
    Next analyte
End Sub
    
Private Sub chkAllFields_Click()
Dim FieldCount As Integer
Dim field As Integer

    FieldCount = lstFields.ListCount - 1
    For field = 0 To FieldCount
        lstFields.Selected(field) = chkAllFields.Value
    Next field
End Sub
    
Private Sub SelectAllOptions()

Dim AnalyteCount As Integer
Dim analyte As Integer
Dim FieldCount As Integer
Dim field As Integer

    chkAllAnalytes.Value = True
    chkAllFields.Value = True

    AnalyteCount = lstAnalytes.ListCount - 1
    For analyte = 0 To AnalyteCount
        lstAnalytes.Selected(analyte) = chkAllAnalytes.Value
    Next analyte
    
    FieldCount = lstFields.ListCount - 1
    For field = 0 To FieldCount
        lstFields.Selected(field) = chkAllFields.Value
    Next field

End Sub

    
Private Sub cmdImportConditions_Click()
Dim ImportFile As Variant
Dim ImportWkBk As Workbook

On Error GoTo ErrorCatch
    Application.ScreenUpdating = False
    ImportFile = Application.GetOpenFilename(Title:="Select File To Import From", MultiSelect:=False)
    If VarType(ImportFile) = vbBoolean Then Exit Sub
    Workbooks.Open Filename:=ImportFile
    Set ImportWkBk = ActiveWorkbook
    Sheets("MetaData").Visible = xlSheetVisible
    Sheets("MetaData").Select
    Call GrabMetaData
    Application.DisplayAlerts = False
    ImportWkBk.Close
    Application.DisplayAlerts = True
    Call SelectAllOptions
    Exit Sub
    
ErrorCatch:
    MsgBox "Data Not Found", vbCritical, "NOT FOUND"
        
End Sub


Private Sub UserForm_Activate()
'Define MS Type Comboxbox
    cmbMSType.AddItem "Pegasus HT"
    cmbMSType.AddItem "Pegasus 4D"
    cmbMSType.AddItem "Pegasus HRT"
    cmbMSType.AddItem "Saturn"
'Define Date Text Box
    txtDate.Value = Date
'Define MS Type Comboxbox
    cmbGCType.AddItem "6890N"
    cmbGCType.AddItem "7890A"
    cmbGCType.AddItem "7890B"
'Define Filament Combobox
    cmbFilament.AddItem "1"
    cmbFilament.AddItem "2"
    cmbFilament.AddItem "?"
'Define Default Experiment Title
    txtExpTitle.Value = "Reproducibility"
    
End Sub

Private Function FindString(ByVal StringToFind As String) As Boolean
On Error GoTo StringNotFound

    Cells.Find(What:=StringToFind, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
FindString = True
Exit Function
StringNotFound:
FindString = False
End Function

Private Sub JumpToString(ByVal StringToFind As String)

    Cells.Find(What:=StringToFind, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate

End Sub

Private Sub ProgressReport(ByVal MessageWkbk As Workbook, ByVal WorkingWkbk As Workbook, _
ByVal Message As String)
    MessageWkbk.Activate
    Application.ScreenUpdating = True
    Application.StatusBar = Message
    Application.ScreenUpdating = False
    WorkingWkbk.Activate
End Sub

Private Sub MakeAndRenameNewSheet(ByVal OriginalSheetName As String, ByVal NewSheetName As String)
    
    Sheets(OriginalSheetName).Select
    Sheets(OriginalSheetName).Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = NewSheetName
    
End Sub

Private Sub FormatAxisNamesSheet()
    Cells.Select
    Selection.NumberFormat = "@"
    Range("A1").Select
End Sub

Private Sub FormatWorkbook(ByVal NoOfFields As Integer, ByVal StartingRow As Integer, _
ByVal StartingColumn As Integer, ByVal NoOfFiles As Integer, Optional ByVal SummarySheet As Boolean, _
Optional ByVal GC_Type As String, Optional ByVal GC_LECO_SN As String, Optional Agilent_GC_SN As String)

Dim HeaderRow As Integer
Dim BodyBotRow As Integer
Dim BodyTopRow As Integer
Dim AveRow As Integer
Dim GCInfoRow As Integer
Dim RSDRow As Integer
Dim TableBottomRow As Integer
Dim StartColumn As Integer
Dim EndColumn As Integer
Dim field As Integer

    HeaderRow = StartingRow
    BodyBotRow = StartingRow + 1
    BodyTopRow = StartingRow + NoOfFiles
    AveRow = BodyTopRow + 1
    GCInfoRow = AveRow + 2
    RSDRow = BodyTopRow + 3
    TableBottomRow = BodyTopRow + 5
    StartColumn = StartingColumn - 1
    EndColumn = StartingColumn + NoOfFields - 1

'Headers row height
    Rows(HeaderRow & ":" & HeaderRow).Select
    
    
    If SummarySheet = True Then
        Selection.RowHeight = 30
    Else
        Selection.RowHeight = 60
    End If
    
'Whole document fill white
    Cells.Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'Headers
    If SummarySheet = True Then
        Range(Cells(HeaderRow, StartColumn), Cells(HeaderRow + 1, EndColumn)).Select
    Else
        Range(Cells(HeaderRow, StartColumn), Cells(HeaderRow, EndColumn)).Select
    End If
    
    Selection.NumberFormat = "@"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    
    If SummarySheet = True Then
        For field = 0 To NoOfFields - 1
            Range(Cells(HeaderRow, (StartColumn + 1) + (field * 3)), _
                Cells(HeaderRow, (StartColumn + 3) + (field * 3))).Select
            Selection.Merge
        Next field
    End If
    
'Body
    If SummarySheet = True Then
        Range(Cells(BodyBotRow + 1, StartColumn), Cells(BodyTopRow + 1, EndColumn)).Select
    Else
        Range(Cells(BodyBotRow, StartColumn), Cells(BodyTopRow, EndColumn)).Select
    End If
    
    Selection.NumberFormat = "0.00"
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    If SummarySheet = True Then
        Range(Cells(BodyTopRow + 1, StartColumn), Cells(BodyTopRow + 1, EndColumn)).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThick
        End With
        Cells(BodyTopRow + 1, StartColumn).Value = "RSD Flags"
        
        Exit Sub
    End If

'Equations
    Range(Cells(AveRow, StartColumn), Cells(TableBottomRow, EndColumn)).Select
    
    Selection.NumberFormat = "0.00"
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    
'RSD & RSDFlsg set to %.00
    Range(Cells(RSDRow, StartColumn), Cells(TableBottomRow, EndColumn)).Select
    Selection.NumberFormat = "0.00%"
    
    Cells(1, 1).Select
End Sub

Private Sub DropEquationsAndFormat(ByVal TopRow As Integer, ByVal EqunColumn As String, ByVal NumberOfFiles As Integer)
Dim formula(2) As String
Dim c As Byte
Dim Equation As String
Dim CellRange As String
    
    formula(0) = "AVERAGE(" & EqunColumn & TopRow + 1 & ":" & EqunColumn & TopRow + NumberOfFiles & ")"
    formula(1) = "STDEV(" & EqunColumn & TopRow + 1 & ":" & EqunColumn & TopRow + NumberOfFiles & ")"
    formula(2) = EqunColumn & TopRow + NumberOfFiles + 2 & "/" & EqunColumn & TopRow + NumberOfFiles + 1
    
    For c = 0 To 2
        Equation = "=IF(" & EqunColumn & NumberOfFiles + 7 & "=TRUE," & formula(c) & "," & Chr(34) & Chr(34) & ")"
        CellRange = EqunColumn & TopRow + NumberOfFiles + 1 + c
        Range(CellRange).Value = Equation
    Next c
    
End Sub

Private Sub ConditionalFormatting(ByVal ColumnLetter As String, ByVal RSDFlagRow As Integer, _
ByVal RSDRow As Integer)

Dim RSDCell As String
Dim RSDFlagCell As String

    RSDCell = ColumnLetter & RSDRow
    RSDFlagCell = "=$" & ColumnLetter & "$" & RSDFlagRow

    Range(RSDCell).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:=RSDFlagCell
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

Private Sub DropEquationTitles(ByVal NoOfFields As Integer, ByVal StartingRow As Integer, _
ByVal StartingColumn As Integer, ByVal NoOfFiles As Integer)

Dim AveRow As Integer
Dim SDRow As Integer
Dim RSDRow As Integer
Dim RSDFlagRow As Integer
Dim GraphDataRow As Integer
Dim DataColumn As Integer
Dim field As Integer
Dim L As String
Dim Equation As String
Dim DataRowStart As Integer
Dim DataRowEnd As Integer

    DataRowStart = StartingRow + NoOfFiles + 7
    DataRowEnd = DataRowStart + NoOfFiles - 1
    AveRow = StartingRow + NoOfFiles + 1
    SDRow = AveRow + 1
    RSDRow = SDRow + 1
    RSDFlagRow = RSDRow + 1
    GraphDataRow = RSDFlagRow + 1
    DataColumn = StartingColumn - 1
    
    Cells(AveRow, DataColumn).Value = "Average"
    Cells(SDRow, DataColumn).Value = "Std. Dev."
    Cells(RSDRow, DataColumn).Value = "RSD"
    Cells(RSDFlagRow, DataColumn).Value = "RSD Flags"
    Cells(GraphDataRow, DataColumn).Value = "Graph Data"
    
    For field = 1 To NoOfFields
        L = DetermineColumnLetter(DataColumn + field)
        Equation = "=IF(COUNTIF(" & L & DataRowStart & ":" & L & DataRowEnd & ",TRUE)>0,TRUE,FALSE)"
        Cells(GraphDataRow, DataColumn + field).Value = Equation
    Next field
    
End Sub

Private Sub ClearTempSheet(ByVal StartingColumnLetter As String, ByVal EndingColumnLetter As String)
    Columns(StartingColumnLetter & ":" & EndingColumnLetter).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub

Private Sub FreezeColumn(ByVal ColumnLetter As String)
    Columns(ColumnLetter & ":" & ColumnLetter).Select
    ActiveWindow.FreezePanes = True
    Cells(1, 1).Select
End Sub

Private Function DetermineEquationLinks(ByVal StartingRowNumber As Integer, ByVal ColumnLetter As String, _
ByVal TotalNumberOfFiles As Integer, ByVal AnalyteName As String, ByVal Calculation As String, ByVal analyte As Integer) As String
Dim SourceRow As Integer
Dim SourceColumn As String
Dim OffsetValue As Byte

    Select Case Calculation
        Case "Average"
            OffsetValue = 1
        Case "Standard Deviation"
            OffsetValue = 2
        Case "Relative Standard Deviation"
            OffsetValue = 3
        Case "Graph Data"
            OffsetValue = 5
    End Select
    
    SourceRow = StartingRowNumber + TotalNumberOfFiles + OffsetValue
    SourceColumn = ColumnLetter
    DetermineEquationLinks = "='" & Left(AnalyteName, 28) & "-" & analyte & "'!" & SourceColumn & SourceRow
End Function

Private Sub PercentNumberFormat(ByVal Row1 As Integer, ByVal Column1 As Integer, _
ByVal Row2 As Integer, ByVal Column2 As Integer)
    Range(Cells(Row1, Column1), Cells(Row2, Column2)).Select
    Selection.NumberFormat = "0.00%"
    Cells(1, 1).Select
End Sub

Private Sub FillSeriesDown(ByVal NumberOfFiles As Integer)

    Cells(3, 1).Select
    Cells(3, 1).Value = 1
    
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Stop:=NumberOfFiles, Trend:=False
    Columns("A:A").Select
    
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    Cells(2, 1).Value = "Inj."
    Cells(1, 1).Select
End Sub

Private Function Formatting_CountNotFounds(ByVal StartingRow As Integer, ByVal StartingColumn As Integer, _
ByVal NumberOfAnalytes As Integer, ByVal NotFoundTitle As String) As Integer
Dim ColumnLetter As String

    Cells(StartingRow + 1, StartingColumn).Select
    Selection.End(xlToRight).Select
    ColumnLetter = DetermineColumnLetter(ActiveCell.Column + 1)
    Formatting_CountNotFounds = ActiveCell.Column + 1
    
    Columns(ColumnLetter & ":" & ColumnLetter).Select
    Selection.UnMerge
    Columns(ColumnLetter & ":" & ColumnLetter).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(ColumnLetter & StartingRow) = NotFoundTitle
    
'Header Formatting
    Range(ColumnLetter & StartingRow & ":" & ColumnLetter & StartingRow + 1).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
'Body Formatting
    Range(ColumnLetter & StartingRow + 2 & ":" & ColumnLetter & StartingRow + NumberOfAnalytes + 1).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.NumberFormat = "0"
    Selection.FormatConditions.Delete
End Function

Private Sub CenterCells()
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub

Private Sub InsertRows()
    Columns("A:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, 1).Select
End Sub

Private Function CreateGraph(ByVal HeaderName As String) As Boolean
Dim Column As Integer
Dim Row As Integer
    Call FindString(HeaderName)
    Column = ActiveCell.Column
    Call FindString("Name")
    Selection.End(xlDown).Select
    Row = ActiveCell.Row
    Cells(Row, Column).Select
    If Cells(Row, Column).Value = "" Then
        CreateGraph = False
    Else
        CreateGraph = Cells(Row, Column).Value
    End If
    Cells(1, 1).Select
End Function

Private Function DetermineColumnLetter(ByVal ColumnNumber As Integer) As String
Dim ColumnLetterOutput(416) As String

    ColumnLetterOutput(1) = "A"
    ColumnLetterOutput(2) = "B"
    ColumnLetterOutput(3) = "C"
    ColumnLetterOutput(4) = "D"
    ColumnLetterOutput(5) = "E"
    ColumnLetterOutput(6) = "F"
    ColumnLetterOutput(7) = "G"
    ColumnLetterOutput(8) = "H"
    ColumnLetterOutput(9) = "I"
    ColumnLetterOutput(10) = "J"
    ColumnLetterOutput(11) = "K"
    ColumnLetterOutput(12) = "L"
    ColumnLetterOutput(13) = "M"
    ColumnLetterOutput(14) = "N"
    ColumnLetterOutput(15) = "O"
    ColumnLetterOutput(16) = "P"
    ColumnLetterOutput(17) = "Q"
    ColumnLetterOutput(18) = "R"
    ColumnLetterOutput(19) = "S"
    ColumnLetterOutput(20) = "T"
    ColumnLetterOutput(21) = "U"
    ColumnLetterOutput(22) = "V"
    ColumnLetterOutput(23) = "W"
    ColumnLetterOutput(24) = "X"
    ColumnLetterOutput(25) = "Y"
    ColumnLetterOutput(26) = "Z"
    
    ColumnLetterOutput(27) = "AA"
    ColumnLetterOutput(28) = "AB"
    ColumnLetterOutput(29) = "AC"
    ColumnLetterOutput(30) = "AD"
    ColumnLetterOutput(31) = "AE"
    ColumnLetterOutput(32) = "AF"
    ColumnLetterOutput(33) = "AG"
    ColumnLetterOutput(34) = "AH"
    ColumnLetterOutput(35) = "AI"
    ColumnLetterOutput(36) = "AJ"
    ColumnLetterOutput(37) = "AK"
    ColumnLetterOutput(38) = "AL"
    ColumnLetterOutput(39) = "AM"
    ColumnLetterOutput(40) = "AN"
    ColumnLetterOutput(41) = "AO"
    ColumnLetterOutput(42) = "AP"
    ColumnLetterOutput(43) = "AQ"
    ColumnLetterOutput(44) = "AR"
    ColumnLetterOutput(45) = "AS"
    ColumnLetterOutput(46) = "AT"
    ColumnLetterOutput(47) = "AU"
    ColumnLetterOutput(48) = "AV"
    ColumnLetterOutput(49) = "AW"
    ColumnLetterOutput(50) = "AX"
    ColumnLetterOutput(51) = "AY"
    ColumnLetterOutput(52) = "AZ"
    
    ColumnLetterOutput(53) = "BA"
    ColumnLetterOutput(54) = "BB"
    ColumnLetterOutput(55) = "BC"
    ColumnLetterOutput(56) = "BD"
    ColumnLetterOutput(57) = "BE"
    ColumnLetterOutput(58) = "BF"
    ColumnLetterOutput(59) = "BG"
    ColumnLetterOutput(60) = "BH"
    ColumnLetterOutput(61) = "BI"
    ColumnLetterOutput(62) = "BJ"
    ColumnLetterOutput(63) = "BK"
    ColumnLetterOutput(64) = "BL"
    ColumnLetterOutput(65) = "BM"
    ColumnLetterOutput(66) = "BN"
    ColumnLetterOutput(67) = "BO"
    ColumnLetterOutput(68) = "BP"
    ColumnLetterOutput(69) = "BQ"
    ColumnLetterOutput(70) = "BR"
    ColumnLetterOutput(71) = "BS"
    ColumnLetterOutput(72) = "BT"
    ColumnLetterOutput(73) = "BU"
    ColumnLetterOutput(74) = "BV"
    ColumnLetterOutput(75) = "BW"
    ColumnLetterOutput(76) = "BX"
    ColumnLetterOutput(77) = "BY"
    ColumnLetterOutput(78) = "BZ"
    
    ColumnLetterOutput(79) = "CA"
    ColumnLetterOutput(80) = "CB"
    ColumnLetterOutput(81) = "CC"
    ColumnLetterOutput(82) = "CD"
    ColumnLetterOutput(83) = "CE"
    ColumnLetterOutput(84) = "CF"
    ColumnLetterOutput(85) = "CG"
    ColumnLetterOutput(86) = "CH"
    ColumnLetterOutput(87) = "CI"
    ColumnLetterOutput(88) = "CJ"
    ColumnLetterOutput(89) = "CK"
    ColumnLetterOutput(90) = "CL"
    ColumnLetterOutput(91) = "CM"
    ColumnLetterOutput(92) = "CN"
    ColumnLetterOutput(93) = "CO"
    ColumnLetterOutput(94) = "CP"
    ColumnLetterOutput(95) = "CQ"
    ColumnLetterOutput(96) = "CR"
    ColumnLetterOutput(97) = "CS"
    ColumnLetterOutput(98) = "CT"
    ColumnLetterOutput(99) = "CU"
    ColumnLetterOutput(100) = "CV"
    ColumnLetterOutput(101) = "CW"
    ColumnLetterOutput(102) = "CX"
    ColumnLetterOutput(103) = "CY"
    ColumnLetterOutput(104) = "CZ"
    
    ColumnLetterOutput(105) = "DA"
    ColumnLetterOutput(106) = "DB"
    ColumnLetterOutput(107) = "DC"
    ColumnLetterOutput(108) = "DD"
    ColumnLetterOutput(109) = "DE"
    ColumnLetterOutput(110) = "DF"
    ColumnLetterOutput(111) = "DG"
    ColumnLetterOutput(112) = "DH"
    ColumnLetterOutput(113) = "DI"
    ColumnLetterOutput(114) = "DJ"
    ColumnLetterOutput(115) = "DK"
    ColumnLetterOutput(116) = "DL"
    ColumnLetterOutput(117) = "DM"
    ColumnLetterOutput(118) = "DN"
    ColumnLetterOutput(119) = "DO"
    ColumnLetterOutput(120) = "DP"
    ColumnLetterOutput(121) = "DQ"
    ColumnLetterOutput(122) = "DR"
    ColumnLetterOutput(123) = "DS"
    ColumnLetterOutput(124) = "DT"
    ColumnLetterOutput(125) = "DU"
    ColumnLetterOutput(126) = "DV"
    ColumnLetterOutput(127) = "DW"
    ColumnLetterOutput(128) = "DX"
    ColumnLetterOutput(129) = "DY"
    ColumnLetterOutput(130) = "DZ"
    
    ColumnLetterOutput(131) = "EA"
    ColumnLetterOutput(132) = "EB"
    ColumnLetterOutput(133) = "EC"
    ColumnLetterOutput(134) = "ED"
    ColumnLetterOutput(135) = "EE"
    ColumnLetterOutput(136) = "EF"
    ColumnLetterOutput(137) = "EG"
    ColumnLetterOutput(138) = "EH"
    ColumnLetterOutput(139) = "EI"
    ColumnLetterOutput(140) = "EJ"
    ColumnLetterOutput(141) = "EK"
    ColumnLetterOutput(142) = "EL"
    ColumnLetterOutput(143) = "EM"
    ColumnLetterOutput(144) = "EN"
    ColumnLetterOutput(145) = "EO"
    ColumnLetterOutput(146) = "EP"
    ColumnLetterOutput(147) = "EQ"
    ColumnLetterOutput(148) = "ER"
    ColumnLetterOutput(149) = "ES"
    ColumnLetterOutput(150) = "ET"
    ColumnLetterOutput(151) = "EU"
    ColumnLetterOutput(152) = "EV"
    ColumnLetterOutput(153) = "EW"
    ColumnLetterOutput(154) = "EX"
    ColumnLetterOutput(155) = "EY"
    ColumnLetterOutput(156) = "EZ"
    
    ColumnLetterOutput(157) = "FA"
    ColumnLetterOutput(158) = "FB"
    ColumnLetterOutput(159) = "FC"
    ColumnLetterOutput(160) = "FD"
    ColumnLetterOutput(161) = "FE"
    ColumnLetterOutput(162) = "FF"
    ColumnLetterOutput(163) = "FG"
    ColumnLetterOutput(164) = "FH"
    ColumnLetterOutput(165) = "FI"
    ColumnLetterOutput(166) = "FJ"
    ColumnLetterOutput(167) = "FK"
    ColumnLetterOutput(168) = "FL"
    ColumnLetterOutput(169) = "FM"
    ColumnLetterOutput(170) = "FN"
    ColumnLetterOutput(171) = "FO"
    ColumnLetterOutput(172) = "FP"
    ColumnLetterOutput(173) = "FQ"
    ColumnLetterOutput(174) = "FR"
    ColumnLetterOutput(175) = "FS"
    ColumnLetterOutput(176) = "FT"
    ColumnLetterOutput(177) = "FU"
    ColumnLetterOutput(178) = "FV"
    ColumnLetterOutput(179) = "FW"
    ColumnLetterOutput(180) = "FX"
    ColumnLetterOutput(181) = "FY"
    ColumnLetterOutput(182) = "FZ"
    
    ColumnLetterOutput(183) = "GA"
    ColumnLetterOutput(184) = "GB"
    ColumnLetterOutput(185) = "GC"
    ColumnLetterOutput(186) = "GD"
    ColumnLetterOutput(187) = "GE"
    ColumnLetterOutput(188) = "GF"
    ColumnLetterOutput(189) = "GG"
    ColumnLetterOutput(190) = "GH"
    ColumnLetterOutput(191) = "GI"
    ColumnLetterOutput(192) = "GJ"
    ColumnLetterOutput(193) = "GK"
    ColumnLetterOutput(194) = "GL"
    ColumnLetterOutput(195) = "GM"
    ColumnLetterOutput(196) = "GN"
    ColumnLetterOutput(197) = "GO"
    ColumnLetterOutput(198) = "GP"
    ColumnLetterOutput(199) = "GQ"
    ColumnLetterOutput(200) = "GR"
    ColumnLetterOutput(201) = "GS"
    ColumnLetterOutput(202) = "GT"
    ColumnLetterOutput(203) = "GU"
    ColumnLetterOutput(204) = "GV"
    ColumnLetterOutput(205) = "GW"
    ColumnLetterOutput(206) = "GX"
    ColumnLetterOutput(207) = "GY"
    ColumnLetterOutput(208) = "GZ"
    
    ColumnLetterOutput(209) = "HA"
    ColumnLetterOutput(210) = "HB"
    ColumnLetterOutput(211) = "HC"
    ColumnLetterOutput(212) = "HD"
    ColumnLetterOutput(213) = "HE"
    ColumnLetterOutput(214) = "HF"
    ColumnLetterOutput(215) = "HG"
    ColumnLetterOutput(216) = "HH"
    ColumnLetterOutput(217) = "HI"
    ColumnLetterOutput(218) = "HJ"
    ColumnLetterOutput(219) = "HK"
    ColumnLetterOutput(220) = "HL"
    ColumnLetterOutput(221) = "HM"
    ColumnLetterOutput(222) = "HN"
    ColumnLetterOutput(223) = "HO"
    ColumnLetterOutput(224) = "HP"
    ColumnLetterOutput(225) = "HQ"
    ColumnLetterOutput(226) = "HR"
    ColumnLetterOutput(227) = "HS"
    ColumnLetterOutput(228) = "HT"
    ColumnLetterOutput(229) = "HU"
    ColumnLetterOutput(230) = "HV"
    ColumnLetterOutput(231) = "HW"
    ColumnLetterOutput(232) = "HX"
    ColumnLetterOutput(233) = "HY"
    ColumnLetterOutput(234) = "HZ"
    
    ColumnLetterOutput(235) = "IA"
    ColumnLetterOutput(236) = "IB"
    ColumnLetterOutput(237) = "IC"
    ColumnLetterOutput(238) = "ID"
    ColumnLetterOutput(239) = "IE"
    ColumnLetterOutput(240) = "IF"
    ColumnLetterOutput(241) = "IG"
    ColumnLetterOutput(242) = "IH"
    ColumnLetterOutput(243) = "II"
    ColumnLetterOutput(244) = "IJ"
    ColumnLetterOutput(245) = "IK"
    ColumnLetterOutput(246) = "IL"
    ColumnLetterOutput(247) = "IM"
    ColumnLetterOutput(248) = "IN"
    ColumnLetterOutput(249) = "IO"
    ColumnLetterOutput(250) = "IP"
    ColumnLetterOutput(251) = "IQ"
    ColumnLetterOutput(252) = "IR"
    ColumnLetterOutput(253) = "IS"
    ColumnLetterOutput(254) = "IT"
    ColumnLetterOutput(255) = "IU"
    ColumnLetterOutput(256) = "IV"
    ColumnLetterOutput(257) = "IW"
    ColumnLetterOutput(258) = "IX"
    ColumnLetterOutput(259) = "IY"
    ColumnLetterOutput(260) = "IZ"

    ColumnLetterOutput(261) = "JA"
    ColumnLetterOutput(262) = "JB"
    ColumnLetterOutput(263) = "JC"
    ColumnLetterOutput(264) = "JD"
    ColumnLetterOutput(265) = "JE"
    ColumnLetterOutput(266) = "JF"
    ColumnLetterOutput(267) = "JG"
    ColumnLetterOutput(268) = "JH"
    ColumnLetterOutput(269) = "JI"
    ColumnLetterOutput(270) = "JJ"
    ColumnLetterOutput(271) = "JK"
    ColumnLetterOutput(272) = "JL"
    ColumnLetterOutput(273) = "JM"
    ColumnLetterOutput(274) = "JN"
    ColumnLetterOutput(275) = "JO"
    ColumnLetterOutput(276) = "JP"
    ColumnLetterOutput(277) = "JQ"
    ColumnLetterOutput(278) = "JR"
    ColumnLetterOutput(279) = "JS"
    ColumnLetterOutput(280) = "JT"
    ColumnLetterOutput(281) = "JU"
    ColumnLetterOutput(282) = "JV"
    ColumnLetterOutput(283) = "JW"
    ColumnLetterOutput(284) = "JX"
    ColumnLetterOutput(285) = "JY"
    ColumnLetterOutput(286) = "JZ"
    
    ColumnLetterOutput(287) = "KA"
    ColumnLetterOutput(288) = "KB"
    ColumnLetterOutput(289) = "KC"
    ColumnLetterOutput(290) = "KD"
    ColumnLetterOutput(291) = "KE"
    ColumnLetterOutput(292) = "KF"
    ColumnLetterOutput(293) = "KG"
    ColumnLetterOutput(294) = "KH"
    ColumnLetterOutput(295) = "KI"
    ColumnLetterOutput(296) = "KJ"
    ColumnLetterOutput(297) = "KK"
    ColumnLetterOutput(298) = "KL"
    ColumnLetterOutput(299) = "KM"
    ColumnLetterOutput(300) = "KN"
    ColumnLetterOutput(301) = "KO"
    ColumnLetterOutput(302) = "KP"
    ColumnLetterOutput(303) = "KQ"
    ColumnLetterOutput(304) = "KR"
    ColumnLetterOutput(305) = "KS"
    ColumnLetterOutput(306) = "KT"
    ColumnLetterOutput(307) = "KU"
    ColumnLetterOutput(308) = "KV"
    ColumnLetterOutput(309) = "KW"
    ColumnLetterOutput(310) = "KX"
    ColumnLetterOutput(311) = "KY"
    ColumnLetterOutput(312) = "KZ"
    
    ColumnLetterOutput(313) = "LA"
    ColumnLetterOutput(314) = "LB"
    ColumnLetterOutput(315) = "LC"
    ColumnLetterOutput(316) = "LD"
    ColumnLetterOutput(317) = "LE"
    ColumnLetterOutput(318) = "LF"
    ColumnLetterOutput(319) = "LG"
    ColumnLetterOutput(320) = "LH"
    ColumnLetterOutput(321) = "LI"
    ColumnLetterOutput(322) = "LJ"
    ColumnLetterOutput(323) = "LK"
    ColumnLetterOutput(324) = "LL"
    ColumnLetterOutput(325) = "LM"
    ColumnLetterOutput(326) = "LN"
    ColumnLetterOutput(327) = "LO"
    ColumnLetterOutput(328) = "LP"
    ColumnLetterOutput(329) = "LQ"
    ColumnLetterOutput(330) = "LR"
    ColumnLetterOutput(331) = "LS"
    ColumnLetterOutput(332) = "LT"
    ColumnLetterOutput(333) = "LU"
    ColumnLetterOutput(334) = "LV"
    ColumnLetterOutput(335) = "LW"
    ColumnLetterOutput(336) = "LX"
    ColumnLetterOutput(337) = "LY"
    ColumnLetterOutput(338) = "LZ"
    
    ColumnLetterOutput(339) = "MA"
    ColumnLetterOutput(340) = "MB"
    ColumnLetterOutput(341) = "MC"
    ColumnLetterOutput(342) = "MD"
    ColumnLetterOutput(343) = "ME"
    ColumnLetterOutput(344) = "MF"
    ColumnLetterOutput(345) = "MG"
    ColumnLetterOutput(346) = "MH"
    ColumnLetterOutput(347) = "MI"
    ColumnLetterOutput(348) = "MJ"
    ColumnLetterOutput(349) = "MK"
    ColumnLetterOutput(350) = "ML"
    ColumnLetterOutput(351) = "MM"
    ColumnLetterOutput(352) = "MN"
    ColumnLetterOutput(353) = "MO"
    ColumnLetterOutput(354) = "MP"
    ColumnLetterOutput(355) = "MQ"
    ColumnLetterOutput(356) = "MR"
    ColumnLetterOutput(357) = "MS"
    ColumnLetterOutput(358) = "MT"
    ColumnLetterOutput(359) = "MU"
    ColumnLetterOutput(360) = "MV"
    ColumnLetterOutput(361) = "MW"
    ColumnLetterOutput(362) = "MX"
    ColumnLetterOutput(363) = "MY"
    ColumnLetterOutput(364) = "MZ"
    
    ColumnLetterOutput(365) = "NA"
    ColumnLetterOutput(366) = "NB"
    ColumnLetterOutput(367) = "NC"
    ColumnLetterOutput(368) = "ND"
    ColumnLetterOutput(369) = "NE"
    ColumnLetterOutput(370) = "NF"
    ColumnLetterOutput(371) = "NG"
    ColumnLetterOutput(372) = "NH"
    ColumnLetterOutput(373) = "NI"
    ColumnLetterOutput(374) = "NJ"
    ColumnLetterOutput(375) = "NK"
    ColumnLetterOutput(376) = "NL"
    ColumnLetterOutput(377) = "NM"
    ColumnLetterOutput(378) = "NN"
    ColumnLetterOutput(379) = "NO"
    ColumnLetterOutput(380) = "NP"
    ColumnLetterOutput(381) = "NQ"
    ColumnLetterOutput(382) = "NR"
    ColumnLetterOutput(383) = "NS"
    ColumnLetterOutput(384) = "NT"
    ColumnLetterOutput(385) = "NU"
    ColumnLetterOutput(386) = "NV"
    ColumnLetterOutput(387) = "NW"
    ColumnLetterOutput(388) = "NX"
    ColumnLetterOutput(389) = "NY"
    ColumnLetterOutput(390) = "NZ"
    
    ColumnLetterOutput(391) = "OA"
    ColumnLetterOutput(392) = "OB"
    ColumnLetterOutput(393) = "OC"
    ColumnLetterOutput(394) = "OD"
    ColumnLetterOutput(395) = "OE"
    ColumnLetterOutput(396) = "OF"
    ColumnLetterOutput(397) = "OG"
    ColumnLetterOutput(398) = "OH"
    ColumnLetterOutput(399) = "OI"
    ColumnLetterOutput(400) = "OJ"
    ColumnLetterOutput(401) = "OK"
    ColumnLetterOutput(402) = "OL"
    ColumnLetterOutput(403) = "OM"
    ColumnLetterOutput(404) = "ON"
    ColumnLetterOutput(405) = "OO"
    ColumnLetterOutput(406) = "OP"
    ColumnLetterOutput(407) = "OQ"
    ColumnLetterOutput(408) = "OR"
    ColumnLetterOutput(409) = "OS"
    ColumnLetterOutput(410) = "OT"
    ColumnLetterOutput(411) = "OU"
    ColumnLetterOutput(412) = "OV"
    ColumnLetterOutput(413) = "OW"
    ColumnLetterOutput(414) = "OX"
    ColumnLetterOutput(415) = "OY"
    ColumnLetterOutput(416) = "OZ"
    
    DetermineColumnLetter = ColumnLetterOutput(ColumnNumber)
End Function

'*******************************
'Graphing procedures

Private Sub GraphManager(ByVal SeriesNumber As Integer, ByVal SheetNameX As String, ByVal SheetNameY As String, _
    ByVal SheetNameN As String, ByVal xColumn As String, ByVal yColumn As String, ByVal NameColumn As String, _
    ByVal NumberOfFiles As Integer, ByVal NumberOfFields As Integer, ByVal StartingRowNo As Byte, _
    ByVal StartingColumnNo As Byte, ByVal NewGraph As Boolean, ByVal AnalyteNumber As Integer, _
    ByVal FieldNumber As Integer, ByVal FieldName As String, ByVal GraphTitlePrefix As String, ByVal FullAnalyteName As String)

Dim LeftLocation As Long
Dim TopLocation As Long
Const ChartWidth As Integer = 680
Const ChartHeight As Integer = 450
Dim NameOfSeries As String
Dim TitleOfChart As String
Dim XAxisTitle As String
Dim YAxisTitle As String
Dim XAxisValues As String
Dim YAxisValues As String
Dim xMin As Single
Dim xMax As Single
Dim yMinMinus5Percent As Single
Dim yMaxPlus5Percent As Single
Dim GraphNumber As Integer
Dim CurrentWorksheet As Worksheet

Application.ScreenUpdating = False
Set CurrentWorksheet = ActiveSheet

    If NewGraph = True Then
        GraphNumber = DetermineNextChartNumber()
        LeftLocation = DetermineLeftLocation(GraphNumber)
        TopLocation = DetermineTopLocation(GraphNumber)
        NameOfSeries = DetermineNameOfSeries(SheetNameN, NameColumn)
        XAxisValues = DetermineXAxisValues(SheetNameX, xColumn, NumberOfFiles, StartingRowNo)
        YAxisValues = DetermineYAxisValues(SheetNameY, yColumn, NumberOfFiles, StartingRowNo)
        TitleOfChart = GraphTitlePrefix & Chr(10) & FullAnalyteName & Chr(10) & DetermineTitleOfChart(SheetNameX, SheetNameY, xColumn, yColumn, _
            CurrentWorksheet, StartingRowNo)
        XAxisTitle = DetermineXAxisTitle(SheetNameX, xColumn, StartingRowNo)
        YAxisTitle = DetermineYAxisTitle(SheetNameY, yColumn, StartingRowNo)
        xMin = CalculatexMin(SheetNameX, StartingRowNo, _
            NumberOfFiles + StartingRowNo - 1, xColumn, CurrentWorksheet)
        xMax = CalculatexMax(SheetNameX, StartingRowNo, _
            NumberOfFiles + StartingRowNo - 1, xColumn, CurrentWorksheet)
        yMinMinus5Percent = CalculateyMinMinus5Percent(SheetNameY, StartingRowNo, _
            NumberOfFiles + StartingRowNo - 1, yColumn, CurrentWorksheet)
        yMaxPlus5Percent = CalculateyMaxPlus5Percent(SheetNameY, StartingRowNo, _
            NumberOfFiles + StartingRowNo - 1, yColumn, CurrentWorksheet)
        Call CreateChart(LeftLocation, TopLocation, ChartWidth, ChartHeight, _
            NameOfSeries, XAxisValues, YAxisValues, TitleOfChart, SeriesNumber, XAxisTitle, YAxisTitle)
        Call RescaleGraphs(xMax, xMin, yMaxPlus5Percent, yMinMinus5Percent)
    ElseIf NewGraph = False Then
    End If
    
    Call DumpAxisValue(CurrentWorksheet, AnalyteNumber, FieldNumber, NameOfSeries, _
        XAxisValues, YAxisValues, FieldName)
    
End Sub

Private Sub CreateChart(ByVal LeftLocation As Long, ByVal TopLocation As Long, _
    ByVal ChartWidth As Integer, ByVal ChartHeight As Integer, ByVal NameOfSeries As String, _
    ByVal XAxisValues As String, ByVal YAxisValues As String, ByVal TitleOfChart As String, _
    ByVal SeriesNum As Integer, ByVal XAxisTitle As String, ByVal YAxisTitle As String)
Dim Graph As ChartObject

    Set Graph = ActiveSheet.ChartObjects.Add(Left:=LeftLocation, _
        Top:=TopLocation, Width:=ChartWidth, Height:=ChartHeight)
    Graph.Activate
    ActiveChart.ChartType = xlXYScatter
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(SeriesNum).Name = NameOfSeries
    ActiveChart.SeriesCollection(SeriesNum).XValues = XAxisValues
    ActiveChart.SeriesCollection(SeriesNum).Values = YAxisValues
    
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = TitleOfChart
    ActiveChart.ChartTitle.Select
    ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
    ActiveChart.Legend.Delete
    
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.Axes(1, xlPrimary).AxisTitle.Text = XAxisTitle
    
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(2, xlPrimary).AxisTitle.Text = YAxisTitle
        
End Sub

Private Sub DumpAxisValue(ByVal CurrentWorksheet As Worksheet, ByVal AnalyteNumber As Integer, _
    ByVal FieldNumber As Integer, ByVal TitleAddress As String, ByVal X_Address As String, _
    ByVal Y_Address As String, ByVal NameOfField As String)

Dim Row As Integer
Dim Title_Column As Integer
Dim X_Column As Integer
Dim Y_Column As Integer

    Row = AnalyteNumber + 1
    Title_Column = 1 + ((FieldNumber - 1) * 3)
    X_Column = 2 + ((FieldNumber - 1) * 3)
    Y_Column = 3 + ((FieldNumber - 1) * 3)
    
    Sheets("AxisPaths").Select
    
    Cells(1, Title_Column).Value = NameOfField
    Cells(Row, Title_Column).Value = TitleAddress
    Cells(Row, X_Column).Value = X_Address
    Cells(Row, Y_Column).Value = Y_Address
    
    CurrentWorksheet.Select
    
End Sub

Private Function DetermineLeftLocation(ByVal GraphNumber As Integer) As Integer

Dim x As Single

    x = (GraphNumber / 2) - Int(GraphNumber / 2)
    
    If x = 0 Then
        DetermineLeftLocation = 700
    ElseIf x = 0.5 Then
        DetermineLeftLocation = 10
    End If

End Function

Private Function DetermineTopLocation(ByVal GraphNumber As Integer) As Integer
Dim x As Single
    x = (WorksheetFunction.RoundUp((GraphNumber / 2), 0) - 1)
    DetermineTopLocation = 10 + (x * 460)
End Function

Private Function DetermineNameOfSeries(ByVal SheetWithName As String, _
ByVal ColumnLetterWithName As String) As String
    DetermineNameOfSeries = "='" & SheetWithName & "'!" & ColumnLetterWithName & "3"
End Function

Private Function DetermineTitleOfChart(ByVal xSheetName As String, ByVal ySheetName As String, _
ByVal xColumnLetter As String, ByVal yColumnLetter As String, ByVal SummarySheet As Worksheet, _
ByVal StartingRow As Integer)
Dim XTitle As String
Dim YTitle As String
Dim RowNum As Integer
    RowNum = StartingRow - 1
    Sheets(xSheetName).Select
    XTitle = Range(xColumnLetter & RowNum).Value
    Sheets(ySheetName).Select
    YTitle = Range(yColumnLetter & RowNum).Value
    SummarySheet.Select
    DetermineTitleOfChart = YTitle & " vs " & XTitle
End Function

Private Function DetermineYAxisValues(ByVal ySheetName As String, ByVal yColumnLetter As String, _
ByVal NumberOfcsvFiles As Long, ByVal StartingRow As Integer) As String
    DetermineYAxisValues = "='" & ySheetName & "'!" & yColumnLetter & StartingRow & ":" & _
        yColumnLetter & (NumberOfcsvFiles + (StartingRow - 1))
End Function

Private Function DetermineXAxisValues(ByVal xSheetName As String, ByVal xColumnLetter As String, _
ByVal NumberOfcsvFiles As Long, ByVal StartingRow As Integer) As String
    DetermineXAxisValues = "='" & xSheetName & "'!" & xColumnLetter & StartingRow & ":" & _
        xColumnLetter & (NumberOfcsvFiles + (StartingRow - 1))
End Function

Private Function DetermineNextChartNumber() As Integer
    DetermineNextChartNumber = ActiveSheet.ChartObjects.Count + 1
End Function

Private Function DetermineXAxisTitle(ByVal xSheetName As String, ByVal xColumnLetter As String, _
 ByVal StartingRow As Integer) As String
 Dim RowNum As Integer
    RowNum = StartingRow - 1
    DetermineXAxisTitle = "='" & xSheetName & "'!" & xColumnLetter & RowNum
End Function

Private Function DetermineYAxisTitle(ByVal ySheetName As String, ByVal yColumnLetter As String, _
 ByVal StartingRow As Integer) As String
 Dim RowNum As Integer
    RowNum = StartingRow - 1
    DetermineYAxisTitle = "='" & ySheetName & "'!" & yColumnLetter & RowNum
End Function

Private Sub RescaleGraphs(ByVal xAxisMax As Single, ByVal xAxisMin As Single, _
ByVal yAxisMax As Single, ByVal yAxisMin As Single)
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = yAxisMin
    ActiveChart.Axes(xlValue).MaximumScale = yAxisMax
    
    ActiveChart.Axes(xlCategory).MinimumScale = xAxisMin
    ActiveChart.Axes(xlCategory).MaximumScale = xAxisMax
End Sub

Private Function CalculateyMinMinus5Percent(ByVal ySheetName As String, ByVal StartRow As Integer, _
ByVal EndRow As Integer, ByVal yColumnLetter As String, ByVal GraphSheet As Worksheet) As Single
Dim Equation As String
Dim MinValue As Single
    Sheets(ySheetName).Select
    Equation = "=MIN(" & yColumnLetter & StartRow & ":" & yColumnLetter & EndRow & ")"
    Range(yColumnLetter & "1").Value = Equation
    MinValue = Range(yColumnLetter & "1").Value
    CalculateyMinMinus5Percent = Round(MinValue - (MinValue * 0.03), 2)
    Range(yColumnLetter & "1").ClearContents
    Cells(1, 1).Select
    GraphSheet.Select
End Function

Private Function CalculateyMaxPlus5Percent(ByVal ySheetName As String, ByVal StartRow As Integer, _
ByVal EndRow As Integer, ByVal yColumnLetter As String, ByVal GraphSheet As Worksheet) As Single
Dim Equation As String
Dim MaxValue As Single
    Sheets(ySheetName).Select
    Equation = "=MAX(" & yColumnLetter & StartRow & ":" & yColumnLetter & EndRow & ")"
    Range(yColumnLetter & "1").Value = Equation
    MaxValue = Range(yColumnLetter & "1").Value
    CalculateyMaxPlus5Percent = Round(MaxValue + (MaxValue * 0.03), 2)
    Range(yColumnLetter & "1").ClearContents
    Cells(1, 1).Select
    GraphSheet.Select
End Function

Private Function CalculatexMin(ByVal xSheetName As String, ByVal StartRow As Integer, _
ByVal EndRow As Integer, ByVal xColumnLetter As String, ByVal GraphSheet As Worksheet)
Dim Equation As String
    Sheets(xSheetName).Select
    Equation = "=MIN(" & xColumnLetter & StartRow & ":" & xColumnLetter & EndRow & ")"
    Range(xColumnLetter & "1").Value = Equation
    CalculatexMin = Round(Range(xColumnLetter & "1").Value, 2)
    Range(xColumnLetter & "1").ClearContents
    Cells(1, 1).Select
    GraphSheet.Select
End Function

Private Function CalculatexMax(ByVal xSheetName As String, ByVal StartRow As Integer, _
ByVal EndRow As Integer, ByVal xColumnLetter As String, ByVal GraphSheet As Worksheet)
Dim Equation As String
    Sheets(xSheetName).Select
    Equation = "=MAX(" & xColumnLetter & StartRow & ":" & xColumnLetter & EndRow & ")"
    Range(xColumnLetter & "1").Value = Equation
    CalculatexMax = Round(Range(xColumnLetter & "1").Value, 2)
    Range(xColumnLetter & "1").ClearContents
    Cells(1, 1).Select
    GraphSheet.Select
End Function

Private Sub CreateSummaryChart(ByVal LeftLocation As Long, ByVal TopLocation As Long, _
    ByVal ChartWidth As Integer, ByVal ChartHeight As Integer, ByVal NameOfSeries As String, _
    ByVal XAxisValues As String, ByVal YAxisValues As String, ByVal TitleOfChart As String, _
    ByVal SeriesNum As Integer, ByVal XAxisTitle As String, ByVal YAxisTitle As String)
Dim Graph As ChartObject

    Set Graph = ActiveSheet.ChartObjects.Add(Left:=LeftLocation, _
        Top:=TopLocation, Width:=ChartWidth, Height:=ChartHeight)
    Graph.Activate
    'ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
    ActiveChart.ChartType = xlXYScatterLines
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(SeriesNum).Name = NameOfSeries
    ActiveChart.SeriesCollection(SeriesNum).XValues = XAxisValues
    ActiveChart.SeriesCollection(SeriesNum).Values = YAxisValues
    
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = TitleOfChart
    
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.Axes(1, xlPrimary).AxisTitle.Text = XAxisTitle
    
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(2, xlPrimary).AxisTitle.Text = YAxisTitle
        
End Sub

Private Sub AddSeriesToGraph(ByVal SeriesCollectionNumber As Integer, ByVal NameAddress As String, _
    ByVal XAddress As String, ByVal YAddress As String)
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(SeriesCollectionNumber).Name = NameAddress
    ActiveChart.SeriesCollection(SeriesCollectionNumber).XValues = XAddress
    ActiveChart.SeriesCollection(SeriesCollectionNumber).Values = YAddress
End Sub

'*****************************************
'Formatting for instrument information table

Private Sub FormatInstrumentalInformation()

'Insert rows at the top of the Summary sheet to make room for the instrumental information
Dim x As Byte
    For x = 0 To 4
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next x

'Format Experiment Title Cells

    Range("D2:O2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
'Merge cells
    Range("D3:E3,D4:E4,D5:E5,F3:G3,F4:G4,F5:G5,H3:I3,H4:I4,J3:K3,J4:K4,L3:M3,L4:M4,N3:O3,N4:O4").Select
    Selection.Merge True

'Format data section of the table
    Range("D3:O5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H5:O5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D3:E5,H3:I4,L3:M4").Select
    Range("L3").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
'Drop labels to instrument information fields
    Range("D3").Value = "Instrument Type:"
    Range("D4").Value = "MS S/N:"
    Range("D5").Value = "Filament:"
    
    Range("H3").Value = "GC Type:"
    Range("H4").Value = "GC S/N:"
    
    Range("L3").Value = "Software Version:"
    Range("L4").Value = "Analysis Date:"
    
    Cells(1, 1).Select
    
End Sub

Private Sub GrabMetaData()
Dim NoOfFields As Integer
Dim NoOfAnalytes As Integer
Dim Row As Integer
    
    lstAnalytes.Clear
    lstFields.Clear
    
    txtDate.Value = Range("ExperimentDate").Value
    txtExpTitle.Value = Range("ExperimentTitle").Value
    txtMS_SN.Value = Range("MS_SN").Value
    cmbMSType.Value = Range("MS_Type").Value
    cmbFilament.Value = Range("Filament").Value
    txtGC_SN.Value = Range("GC_SN").Value
    cmbGCType.Value = Range("GC_Type").Value
    txtSoftwareVersion.Value = Range("SoftwareVersion").Value
    
    NoOfAnalytes = Range("NoOfAnalytes").Value
    NoOfFields = Range("NoOfFields").Value
    
    For Row = 4 To NoOfAnalytes + 3
        lstAnalytes.AddItem Cells(Row, 2).Value
    Next Row
    
    For Row = 4 To NoOfFields + 3
        lstFields.AddItem Cells(Row, 3).Value
    Next Row
    
End Sub

Private Sub cmdLastConditions_Click()
    Application.ScreenUpdating = False
    Sheets("MetaData").Select
    Call GrabMetaData
    Call SelectAllOptions
    Sheets("Sheet1").Select
End Sub

Private Function DesktopAddress() As String
    DesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator
End Function

Private Function MassFieldCorrection(ByVal StartingStringValue As String, ByVal FrontEndCharacter As String, ByVal BackEndCharacter As String) As String
Dim BackEndRemoved As String
Dim BothEndsRemoved As String
Dim CharacterPositionOfBackEndCharacter As Integer
Dim LengthOfStringOnceBackEndIsRemoved As Integer
Dim CharacterPositionOfFrontEndCharacter As Integer
Dim LengthOfStringOnceFrontEndIsRemoved As Integer

    
    CharacterPositionOfBackEndCharacter = InStr(StartingStringValue, BackEndCharacter)
    LengthOfStringOnceBackEndIsRemoved = CharacterPositionOfBackEndCharacter - 1
    BackEndRemoved = Left(StartingStringValue, LengthOfStringOnceBackEndIsRemoved)
    
    CharacterPositionOfFrontEndCharacter = InStr(BackEndRemoved, FrontEndCharacter)
    LengthOfStringOnceFrontEndIsRemoved = LengthOfStringOnceBackEndIsRemoved - CharacterPositionOfFrontEndCharacter
    BothEndsRemoved = Right(BackEndRemoved, LengthOfStringOnceFrontEndIsRemoved)
    
    MassFieldCorrection = BothEndsRemoved
    
End Function

Private Sub AddMarkersToSeries(ByVal ChartNumber As Long, ByVal SeriesNumber As Long)
    
Dim cht As Chart
Dim ser As Series

    
    ActiveSheet.ChartObjects("Chart " & ChartNumber).Activate
    Set cht = ActiveChart
    
        Set ser = cht.SeriesCollection(SeriesNumber)
        ser.MarkerStyle = xlMarkerStyleSquare

End Sub

Private Sub ChangeLegendFont(ByVal ChartNumber As Byte, ByVal FontSize As Integer)

    ActiveSheet.ChartObjects("Chart " & ChartNumber).Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementLegendBottom)
    
    ActiveChart.Legend.Font.Size = FontSize
    

End Sub
