VBA MACRO modToolsDataCleansing.bas 
in file: temp_extract/xl/vbaProject.bin - OLE stream: 'VBA/modToolsDataCleansing'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
' ------------------------------------------------------
' Name: modToolsDataCleansing
' Kind: Module
' Purpose: Series of procedures to control the "Data Cleansing Tools" section on the Ribbon UI
' Author: Corey Brosam
' Date: 01/12/2021
' ------------------------------------------------------
Option Explicit
Private Addr As String
Private rng As Range
Private cell As Range
Private Arr() As Variant
Private rowIndex As Long
Private colIndex As Long
Private lrow As Long
Private lcol As Long
Private MyRow As Long
Private MyCol As Long

Sub Autofit_Columns()

    'Autofits the columns of the selected range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Autofits the columns of the range selected
    
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    rng.Columns.AutoFit
    
TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Autofit_Columns", "modToolsDataCleansing")

    Resume TidyUp
    
End Sub

Sub Blank_Cells()

    'Selects, highlights, and fills blank cells in the selected range

    Dim SelectRange As Range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Assigns address from the "Cell Selection" on the Ribbon
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column
    
    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    
    'Calls methods to launch the progressbar
    Run "reset_Vasr"
    Run "ShowUserForm"
    
    'Sets value of total number of iterations on the progressbar
    TotalToPro = lrow * lcol
    
    'Loops through each cell in the range and concatenates the addresses of each blank cell
    For rowIndex = 1 To lrow
        
        For colIndex = 1 To lcol
            
            If Cells(MyRow, MyCol) = "" Then
                
                If Not SelectRange Is Nothing Then
                    
                    Set SelectRange = Union(SelectRange, Range(Cells(MyRow, MyCol).Address))
                    
                Else
                    
                    Set SelectRange = Range(Cells(MyRow, MyCol).Address)
                    
                End If
                
            End If
            
            MyCol = MyCol + 1
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next
        
        MyRow = MyRow + 1
        MyCol = Range(Addr).Column
        
    Next
    
    'Selects, applies highlighting (if requested) and populates values (if requested) to the blank cells
    If SelectRange Is Nothing Then
        
        MsgBox ("No Blank Cells Found")
        
    Else
        
        SelectRange.Select
        
        If ThisWorkbook.Cur_Blank_Checked = True Then
            
            Selection.Interior.Color = vbYellow
            
        End If
        
        If ThisWorkbook.Cur_Blank_Val <> "" Then
            
            Selection = ThisWorkbook.Cur_Blank_Val
            
        End If
        
    End If
    
    Application.Calculation = xlCalculationAutomatic
    
TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Blank_Cells", "modToolsDataCleansing")

    Resume TidyUp

End Sub

Sub Change_Case(id As String)

    'Changes the casing on every cell in the selected range

    Dim MyString As String
    
    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler
    
    'Assigns address from the "Cell Selection" on the Ribbon
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column
    
    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    
    'Calls function to test if a filter is in place
    If TestTheFilter(rng) = False Then
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = lrow
        
        'Loops through each cell in the range and applies the case formatting
        ReDim Arr(1 To lrow, 1 To lcol)
        
        For rowIndex = 1 To lrow
            
            For colIndex = 1 To lcol
                
                If id = "Item1" Then
                    
                    Arr(rowIndex, colIndex) = LCase(Cells(MyRow, MyCol))
                    
                ElseIf id = "Item2" Then
                    
                    Arr(rowIndex, colIndex) = UCase(Cells(MyRow, MyCol))
                    
                ElseIf id = "Item3" Then
                    
                    Arr(rowIndex, colIndex) = Application.WorksheetFunction.Proper(Cells(MyRow, MyCol))
                    
                Else
                    
                    Arr(rowIndex, colIndex) = ProperCaps(Cells(MyRow, MyCol).Value)
                    
                End If
                
                MyCol = MyCol + 1
            Next
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
            MyRow = MyRow + 1
            MyCol = Range(Addr).Column
        Next
        
        'Applies the values in the array to the originally selected range
        rng.Value = Arr
        
    Else
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = rng.Cells.Count
        
        'Loops through each cell in the range and applies the case formatting
        For Each cell In rng
            
            If cell.EntireRow.Hidden = False Then
                
                If id = "Item1" Then
                    
                    cell.Value = LCase(cell.Value)
                    
                ElseIf id = "Item2" Then
                    
                    cell.Value = UCase(cell.Value)
                    
                ElseIf id = "Item3" Then
                    
                    cell.Value = Application.WorksheetFunction.Proper(cell.Value)
                    
                Else
                    
                    MyString = ProperCaps(cell.Value)
                    
                End If
                
            End If
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next cell
        
    End If

TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Change_Case", "modToolsDataCleansing")

    Resume TidyUp

End Sub

Sub Flip_Signs()

    'Flips the number signage of every cell in the selected range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Assigns address from the "Cell Selection" on the Ribbon

    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column
    
    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    
    'Calls function to test if a filter is in place
    If TestTheFilter(rng) = False Then
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = lrow
        
        ReDim Arr(1 To lrow, 1 To lcol)
        
        'Loops through each cell in the range and flips its signage
        For rowIndex = 1 To lrow
        
            For colIndex = 1 To lcol
            
                If Cells(MyRow, MyCol) = "" Then
                
                    Arr(rowIndex, colIndex) = ""
                
                Else
            
                    Arr(rowIndex, colIndex) = -Cells(MyRow, MyCol).Value

                End If

                MyCol = MyCol + 1
                
            Next
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
            MyRow = MyRow + 1
            MyCol = Range(Addr).Column
            
        Next
        
        rng.Value = Arr
        
    Else
        
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = rng.Cells.Count
        
        'Loops through each cell in the range and flips its signage
        For Each cell In rng
            
            If cell.EntireRow.Hidden = True Or cell.Value = "" Then
                
                'Do Nothing
                
            Else
                
                cell.Value = -cell.Value
                
            End If
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next cell
        
    End If
    
    Application.Calculation = xlCalculationAutomatic

TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Flip_Signs", "modToolsDataCleansing")

    Resume TidyUp

End Sub

Function ProperCaps(strIn As String) As String
    
    'Function to apply sentence casing. For use with "Sub Change_Case" method
    
    Dim objRegex As Object
    Dim objRegMC As Object
    Dim objRegM As Object
    
    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler
    
    Set objRegex = CreateObject("vbscript.regexp")
    strIn = LCase$(strIn)
    
    With objRegex
        .Global = True
        .ignoreCase = True
        .Pattern = "(^|[\.\?\!\r\t]\s?)([a-z])"
        
        If .test(strIn) Then
            
            Set objRegMC = .Execute(strIn)
            
            For Each objRegM In objRegMC
                
                Mid$(strIn, objRegM.firstindex + 1, objRegM.length) = UCase$(objRegM)
                
            Next
            
        End If
        
    End With
    
    ProperCaps = strIn
    
TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    Exit Function

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "ProperCaps", "modToolsDataCleansing")

    Resume TidyUp
    
End Function

Sub Reset_Variables_Data_Cleansing_Tools()

    'Resets variables for use in other procedures

    Addr = ""
    Set rng = Nothing
    Set cell = Nothing
    rowIndex = 0
    colIndex = 0
    lrow = 0
    lcol = 0
    MyRow = 0
    MyCol = 0

End Sub

Public Function TestTheFilter(MyRng As Range) As Boolean
    
    'Function to test if the range the user has selected is filtered
    
    Dim l As Long
    Dim k As Long
    
    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler
    
    l = MyRng.Rows.Count
    k = MyRng.SpecialCells(xlCellTypeVisible).Count
    
    If l > k Then TestTheFilter = True
    
TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    Exit Function

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "TestTheFilter", "modToolsDataCleansing")

    Resume TidyUp
    
End Function

Sub Text_To_Value()

    'Converts all numbers stored as text to values for cells in the selected range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Assigns address from the "Cell Selection" on the Ribbon
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column
    
    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub

    Application.Calculation = xlCalculationManual
    
    'Calls function to test if a filter is in place
    If TestTheFilter(rng) = False Then
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = lrow
        
        
        ReDim Arr(1 To lrow, 1 To lcol)
        
        'Loops through each cell in the range to convert to values
        For rowIndex = 1 To lrow
            
            For colIndex = 1 To lcol
                Arr(rowIndex, colIndex) = Cells(MyRow, MyCol).Value
                MyCol = MyCol + 1
            Next
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
            MyRow = MyRow + 1
            MyCol = Range(Addr).Column
        Next
        
        'Converts the entire array to "General" formatting
        rng.NumberFormat = "General"
        rng.Value = Arr
        
    Else
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = rng.Cells.Count
        
        'Loops through each cell in the range to convert to values
        For Each cell In rng
            
            If cell.EntireRow.Hidden = False Then
                
                cell.Value = cell.Value
                cell.NumberFormat = "General"
                
            End If
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next cell
        
    End If
    
    Application.Calculation = xlCalculationAutomatic

TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Text_To_Value", "modToolsDataCleansing")

    Resume TidyUp

End Sub

Public Function Too_Many_Cells(lrow As Long, lcol As Long) As Boolean
    
    'Function to test if the number of cells selected is over 100,000. If so, it gives the user the option to exit
    
    If (lrow * lcol) > 100000 Then
        
        'Tests if error 6 appears, which is the error if Excel doesn't have enough resources to complete the operation
        'Resets the error handler even if error is not occured so as to not leave an open error statement
        On Error Resume Next
        
        If Err.Number = 6 Then
            
            MsgBox ("You've selected too many cells for the operation. Please select a smaller range and try again")
            Too_Many_Cells = False
            On Error GoTo 0
            Exit Function
            
        Else
            
            On Error GoTo 0
            
        End If
        
        If MsgBox("You have selected more than 100,000 cells to convert, which will take a while to process. " & "Are you sure you would like to proceed?", vbQuestion + vbYesNo, "Range Threshold Reached") = vbNo Then
            
            Too_Many_Cells = False
            Exit Function
            
        End If
        
        Too_Many_Cells = True
        
    End If
    
    Too_Many_Cells = True
    
End Function

Sub Trim_Range()

    'Removes the leading and trailing blank spaces for all cells in the selected range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Assigns address from the "Cell Selection" on the Ribbon
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column

    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub

    Application.Calculation = xlCalculationManual
    
    'Calls function to test if a filter is in place
    If TestTheFilter(rng) = False Then
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = lrow
        
        ReDim Arr(1 To lrow, 1 To lcol)
        
        'Loops through each cell in the range to trim the leading and trailing spaces
        For rowIndex = 1 To lrow
            For colIndex = 1 To lcol
                Arr(rowIndex, colIndex) = Trim(Cells(MyRow, MyCol))
                MyCol = MyCol + 1
            Next
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
            MyRow = MyRow + 1
            MyCol = Range(Addr).Column
        Next
        
        rng.Value = Arr
        
    Else
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = rng.Cells.Count
        
        'Loops through each cell in the range to trim the leading and trailing spaces
        For Each cell In rng
            
            If cell.EntireRow.Hidden = False Then
                
                cell.Value = Trim(cell.Value)
                
            End If
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next cell
        
    End If
    
    Application.Calculation = xlCalculationAutomatic

TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Trim_Range", "modToolsDataCleansing")

    Resume TidyUp

End Sub

Sub Value_To_Text()

    'Converts all numbers stored as general format to text for each cell in the selected range

    'Code dumps off to the error handler if error is encountered
    On Error GoTo MyHandler

    'Assigns address from the "Cell Selection" on the Ribbon
    Addr = ThisWorkbook.Cur_Cell_Sel_Address
    
    Set rng = Range(Addr)
    
    'Setting variables to set up array
    lrow = Range(Addr).Rows.Count
    lcol = Range(Addr).Columns.Count
    MyRow = Range(Addr).Row
    MyCol = Range(Addr).Column
    
    'Calls function to test if too many cells have been selected
    If Too_Many_Cells(lrow, lcol) = False Then Exit Sub
    
    'Calls function to test if a filter is in place
    If TestTheFilter(rng) = False Then
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = lrow
        
        'Loops through each cell in the range to convert to text
        ReDim Arr(1 To lrow, 1 To lcol)
        
        Application.Calculation = xlCalculationManual
        
        For rowIndex = 1 To lrow
            
            For colIndex = 1 To lcol
                
                Arr(rowIndex, colIndex) = CStr(Cells(MyRow, MyCol))
                
                MyCol = MyCol + 1
                
            Next
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
            MyRow = MyRow + 1
            MyCol = Range(Addr).Column
            
        Next
        
        'Applies text formatting to the entire array
        rng.NumberFormat = "@"
        rng.Value = Arr
        
    Else
        
        'Calls methods to launch the progressbar
        Run "reset_Vasr"
        Run "ShowUserForm"
        TotalToPro = rng.Cells.Count
        
        'Loops through each cell in the range to convert to text
        For Each cell In rng
            
            If cell.EntireRow.Hidden = False Then
                
                cell.Value = CStr(cell.Value)
                cell.NumberFormat = "@"
                
            End If
            
            'Tests if the "Stop" button has been clicked on the progressbar. Exits the subroutine if so.
            If BarNext = False Then
                
                Unload UF_ProgressBar_2
                Exit Sub
                
            End If
            
        Next cell
        
    End If
    
    Application.Calculation = xlCalculationAutomatic

TidyUp:

    'Resetting error handler

    On Error GoTo 0
    
    'Reset variables
    
    Reset_Variables_Data_Cleansing_Tools
    
    'Closes the progressbar Userform
    closefr
    
    Exit Sub

MyHandler:

    'Calls error handler function if error is encountered

    Call ReportError(Err.Description, Err.Number, "Value_To_Text", "modToolsDataCleansing")

    Resume TidyUp

End Sub
-------------------------------------------------------------------------------
