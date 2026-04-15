Attribute VB_Name = "Main"
'@Folder("HCM_AB_VALIDATION")
Option Explicit

'@Ignore EncapsulatePublicField
Public Sesh As Session

'@VariableDescription("Stores the Unix Timestamp at runtime, set in the Main method.")
Private UnixTimestamp As LongLong
Attribute UnixTimestamp.VB_VarDescription = "Stores the Unix Timestamp at runtime, set in the Main method."

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method.")
Public Function GetTimestamp() As LongLong
Attribute GetTimestamp.VB_Description = "Returns the Unix Timestamp recorded at runtime in the Main method."
    GetTimestamp = UnixTimestamp
End Function

'@Description("Returns the Unix Timestamp recorded at runtime in the Main method as a string.")
Public Function GetTimestampStr() As String
Attribute GetTimestampStr.VB_Description = "Returns the Unix Timestamp recorded at runtime in the Main method as a string."
    GetTimestampStr = Trim$(Str$(GetTimestamp()))
End Function

Public Function UnixTime() As LongLong
    UnixTime = DateDiff("s", "1/1/1970 00:00:00", Now)
End Function

'@EntryPoint
Public Sub Main()
    UnixTimestamp = UnixTime()
    
    Dim wbOutput As Workbook
    Dim wbJobData As Workbook
    Dim wbABValidation As Workbook
    Dim ws As Worksheet
    Dim wsRepOut As Worksheet
    Dim wsNonRepOut As Worksheet
    Dim wsJobData As Worksheet
    Dim wsABValidation As Worksheet

    Set Sesh = New Session
    form_fileselect.Show
    ' if file selection form was closed without clicking the button to run this application then terminate
    If Sesh.FormClosedWithoutRunning Then Exit Sub
    
    Set wbOutput = Workbooks.Add
    With wbOutput
      '  .Name = "Absence Validation Output"
        .SaveAs Filename:="ABValidation_" & GetTimestampStr()
        
        ' assign original sheet1 to a variable
        Set wsRepOut = .Sheets.Item(1)
        wsRepOut.Name = "Rep Output"
        
        ' create non-rep sheet
        wsRepOut.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsNonRepOut = .Sheets.Item(.Sheets.Count)
        wsNonRepOut.Name = "NonRep Output"
        
    End With
    
    Set wbJobData = Workbooks.Open( _
        Filename:=Sesh.fpathJobData _
        , ReadOnly:=True _
    )
    
    Set ws = wbJobData.Sheets.Item(1)
    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsJobData = .Sheets.Item(.Sheets.Count)
    End With
    Set ws = Nothing
    wbJobData.Close
    
    With wsJobData
        .Name = "Job Data"
        .Rows.Item(1).EntireRow.Delete
        
        ' A1:AZ1, reverse order, check for values
        ' AZ is overkill, as the columns currently stop earlier
        ' But for future-proofing, I doubt the number of cols
        ' will ever exceed AZ
        ' and if they do, I guess this will have to be updated
        Dim col As Long
        Dim Value As String
        For col = 26 * 2 To 1 Step -1
            Value = .Cells.Item(1, col).Value2
            If Not (Value = "Employee ID" _
                Or Value = "Employee Primary Name" _
                Or Value = "Employee Class" _
                Or Value = "Lv Accrual Dt" _
                Or Value = "Union Member" _
                Or Value = "Full/Part" _
                ) Then
                .Cells.Item(1, col).EntireColumn.Delete
            End If
            Next col
    End With
    
    ' Remove non-classified employees
    AddAutoFilter wsJobData, wsJobData.Range("$A$1:$G$1"), 3, "<>CLA"
    DeleteUnfilteredRows wsJobData
    wsJobData.AutoFilterMode = False
    
    Set wbABValidation = Workbooks.Open( _
        Filename:=Sesh.fpathABValidation _
        , ReadOnly:=True _
    )
    
    Set ws = wbABValidation.Sheets.Item(1)
    With wbOutput
        ws.Copy After:=.Sheets.Item(.Sheets.Count)
        Set wsABValidation = .Sheets.Item(.Sheets.Count)
    End With
    Set ws = Nothing
    wbABValidation.Close
    
    With wsABValidation
        .Name = "AB Validation"
        .Rows.Item(1).EntireRow.Delete
        
        For col = 26 To 1 Step -1
            Value = .Cells.Item(1, col).Value2
            If Not (Value = "Name" _
                Or Value = "ID" _
                Or Value = "PIN Name" _
                Or Value = "Slice Begin Date" _
                Or Value = "Slice End Date" _
                Or Value = "Leave Accrual" _
                Or Value = "Leave Balance" _
                ) Then
                .Cells.Item(1, col).EntireColumn.Delete
            End If
            Next col
    End With
    
    With wsJobData
        wsJobData.Range("G1") _
            .Value2 = "Years of Service"
    
        .Range("G2:G" & wsJobData.Range("A1").CurrentRegion.Rows.Count) _
            .Formula = "=IF(ISBLANK(E2),"""",DATEDIF(E2,TODAY(),""Y""))"
        
        Dim CutCopyMode As Boolean
        CutCopyMode = Application.CutCopyMode
        
        .Range("A:A").EntireColumn.Insert (XlDirection.xlToRight)
        .Range("A:A").Value = .Range("C:C").Value
        .Range("C:C").EntireColumn.Delete
        
        Application.CutCopyMode = CutCopyMode
    End With
    
    
    With wsRepOut
        ' Filter to only Union employees and copy to Rep Output sheet
        AddAutoFilter wsJobData, wsJobData.Range("A:G"), 6, "Y"
        wsJobData.Range("A:A") _
            .SpecialCells(xlCellTypeVisible) _
            .Copy .Range("A1")
            
        Dim RepOutRowCount As Long
        RepOutRowCount = .Range("A1").CurrentRegion.Rows.Count
        
        .Range("B1").Value2 = "Name"
        .Range("B2:B" & RepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 2, FALSE)"
        
        .Range("C1").Value2 = "FT/PT"
        .Range("C2:C" & RepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 4, FALSE)"
        
        .Range("D1").Value2 = "Years of Service"
        .Range("D2:D" & RepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 7, FALSE)"
        
        .Range("E1").Value2 = "Leave Accrual"
        .Range("E2:E" & RepOutRowCount) _
            .Formula = "=VLOOKUP(TEXT(A2, ""0""), '" & wsABValidation.Name & "'!B:G, 5, FALSE)"
    End With
    
    With wsNonRepOut
        ' Filter to only non-union employees and copy to non-rep output sheet
        AddAutoFilter wsJobData, wsJobData.Range("A:G"), 6, "N"
        wsJobData.Range("A:A") _
            .SpecialCells(xlCellTypeVisible) _
            .Copy .Range("A1")
            
        Dim NonRepOutRowCount As Long
        NonRepOutRowCount = .Range("A1").CurrentRegion.Rows.Count
        
        .Range("B1").Value2 = "Name"
        .Range("B2:B" & NonRepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 2, FALSE)"
        
        .Range("C1").Value2 = "FT/PT"
        .Range("C2:C" & NonRepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 4, FALSE)"
        
        .Range("D1").Value2 = "Years of Service"
        .Range("D2:D" & NonRepOutRowCount) _
            .Formula = "=VLOOKUP(A2, '" & wsJobData.Name & "'!A:G, 7, FALSE)"
        
        .Range("E1").Value2 = "Leave Accrual"
        .Range("E2:E" & NonRepOutRowCount) _
            .Formula = "=VLOOKUP(TEXT(A2, ""0""), '" & wsABValidation.Name & "'!B:G, 5, FALSE)"
    End With
    
    AddAutoFilter wsRepOut, wsRepOut.Range("A1").CurrentRegion, 5, "#N/A"
    DeleteUnfilteredRows wsRepOut
    wsRepOut.AutoFilterMode = False
    
    AddAutoFilter wsNonRepOut, wsNonRepOut.Range("A1").CurrentRegion, 5, "#N/A"
    DeleteUnfilteredRows wsNonRepOut
    wsNonRepOut.AutoFilterMode = False
    
    wbOutput.Close SaveChanges:=True
End Sub

'@Ignore UseMeaningfulName
Private Sub AddAutoFilter( _
        ByVal ws As Worksheet, _
        ByVal rg As Range, _
        Optional ByVal ColOffset As Long, _
        Optional ByVal Criteria As String _
    )
    ws.AutoFilterMode = False
    rg.AutoFilter Field:=ColOffset, Criteria1:=Criteria
End Sub

' I could dynamically determine the number of rows
' and columns
' But I don't think I need to worry about this data
' having more than 10k rows
' If this ends up being too intense for my CPU then
' I will change it
' Time is Money
'@Ignore UseMeaningfulName
Private Sub DeleteUnfilteredRows( _
        ByVal ws As Worksheet, _
        Optional ByVal IncludeHeader As Boolean = False _
    )
    If IncludeHeader Then
        ' Get range A1:AZ10k,
        ' then delete anything not filtered out from that range
        ws.Range("$A$1:$AZ$10000") _
            .SpecialCells(xlCellTypeVisible) _
            .EntireRow _
            .Delete
    Else
        ' Get range A1:A10k,
        ' Then delete everything not filtered
        ' Excluding the header
        ws.Range("$A$1:$AZ$10000") _
            .Offset(RowOffset:=1) _
            .SpecialCells(xlCellTypeVisible) _
            .EntireRow _
            .Delete
    End If
End Sub

' TODO: separate into a separate, referenceable package so that it can be used in multiple projects.
'@Description "Open a workbook and copy a worksheet from it to another workbook, returning the copy."
Private Function CopyWorkbookSheet( _
        ByVal wbFromPath As String, _
        ByVal wbTo As Workbook, _
        ByVal SheetIndex As Long, _
        ByVal Name As String _
    ) As Worksheet
    Dim wbFrom As Workbook
    Set wbFrom = Workbooks.Open( _
        Filename:=wbFromPath _
        , ReadOnly:=True _
    )
    With wbTo
        wbFrom.Sheets.Item(SheetIndex).Copy After:=.Sheets.Item(.Sheets.Count)
        Set CopyWorkbookSheet = .Sheets.Item(.Sheets.Count)
        CopyWorkbookSheet.Name = Name
    End With
    wbFrom.Close
End Function
