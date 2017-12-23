Attribute VB_Name = "DATA_RETRIEVAL"
Dim symbol As String
Dim number As String

Dim system As Object
Dim session As Object
Dim screen As Object

Dim logOnScreen As String
Dim GMENtest As String
Dim errorCount As Integer




Sub Data_Retrieval()
'Created by Adam Gill
'Date: 16/12/09  -  Updated: 17/03/10
'Basic PROD Extract of the last three modules and related data from PROD/PMS
    
    exitOption = False
       
       
    With Application
        .ScreenUpdating = False
        .StatusBar = "Loading macro..."
        .DisplayAlerts = False
    End With

    'Prepares sheet for split of policy numbers
    Call prepareSheet
    
    'Actions the split of policy numbers
    Application.StatusBar = "Seperating policy numbers..."
    Call TextToColumn

    'Confirms that PMS/Prod is open and active
    Application.StatusBar = "Initiating PMS..."
    Call startPROD
    If exitOption = True Then
        Exit Sub
    End If
    
    'Retreives policy information from PMS/PROD
    Application.StatusBar = "Commencing data retrieval..."
    Call PROD_Validate
    
    'Finalises spreadsheet formatting and restores access to Excel/etc
    Call endRoutine
    
End Sub



Private Sub prepareSheet()
'This function prepares the activework book for split of policy numbers
    
    Sheets("Extract").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Selection.ClearFormats
    
    'Prepares columns, by moving existing data across
    Sheets("Extract").Select
    Columns("A:Z").Select
    Selection.NumberFormat = "@"
    
    Range("I1:J1,O1:P1,U1:V1").Select
    Selection.NumberFormat = "d/m/yy;@"
    
    Sheets("Start").Select
    Columns("A").Select
    Selection.Copy
    Sheets("Extract").Select
    Columns("A").Select
    ActiveSheet.Paste
    
    'Makes copy of policy number in spare column 'V'
    Columns("A").Select
    Selection.Copy
    Columns("B").Select
    ActiveSheet.Paste
    
    ActiveSheet.AutoFilterMode = False
    
    nameHeaders
    
End Sub

    
Private Sub resetFilter()

    Application.ActiveSheet.UsedRange
    
    'Turns off Filter
    ActiveSheet.AutoFilterMode = False
    Range("A1").Select
    
End Sub


Private Sub nameHeaders()

    
    'General policy details
    Range("A1").FormulaR1C1 = "Policy Number"
    Range("B1").FormulaR1C1 = "Symbol"
    Range("C1").FormulaR1C1 = "Number"
    Range("D1").FormulaR1C1 = "Agent #"
    Range("E1").FormulaR1C1 = "P/C"
    Range("F1").FormulaR1C1 = "Insp Dist"
    Range("G1").FormulaR1C1 = "Branch"
    Range("A1:G1").Select
    With Selection
        .Interior.Pattern = xlSolid
        .Interior.Color = 5287936
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    'Latest PMS module      (0)
    Range("H1").FormulaR1C1 = "MOD: Current"
    Range("I1").FormulaR1C1 = "Start"
    Range("J1").FormulaR1C1 = "End"
    Range("K1").FormulaR1C1 = "Predebit"
    Range("L1").FormulaR1C1 = "U/W Code"
    Range("M1").FormulaR1C1 = "EDI"
    Range("H1:M1").Select
    With Selection
        .Interior.Pattern = xlSolid
        .Interior.Color = 15773696
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    
    'Previous PMS module    (-1)
    Range("N1").FormulaR1C1 = "MOD: -1"
    Range("O1").FormulaR1C1 = "Start"
    Range("P1").FormulaR1C1 = "End"
    Range("Q1").FormulaR1C1 = "Predebit"
    Range("R1").FormulaR1C1 = "U/W Code"
    Range("S1").FormulaR1C1 = "EDI"
    Range("N1:S1").Select
    With Selection
        .Interior.Pattern = xlSolid
        .Interior.Color = 12611584
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    '2x Previous PMS module      (-2)
    Range("T1").FormulaR1C1 = "MOD: -2"
    Range("U1").FormulaR1C1 = "Start"
    Range("V1").FormulaR1C1 = "End"
    Range("W1").FormulaR1C1 = "Predebit"
    Range("X1").FormulaR1C1 = "U/W Code"
    Range("Y1").FormulaR1C1 = "EDI"
    Range("T1:Y1").Select
    With Selection
        .Interior.Pattern = xlSolid
        .Interior.Color = 49407
        .Font.ColorIndex = 1
        .Font.Bold = True
        .Font.Italic = True
    End With
    
        
End Sub



Private Sub TextToColumn()
Attribute TextToColumn.VB_ProcData.VB_Invoke_Func = " \n14"
' TextToColumn Macro
Dim counterPecentage As Integer
Dim numberOfPolicies As Integer
Dim counter As Integer
    
    
    Range("b2").Select
    Do
        Selection.TextToColumns Destination:=ActiveCell, DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(1, 1), Array(4, 2), Array(11, 9), Array(12, 2), Array(14, 1)), TrailingMinusNumbers:=True
        
    ActiveCell.Offset(1, 0).Select
    Loop Until IsEmpty(ActiveCell) = True

End Sub




Private Sub endRoutine()

Cells.Select
Cells.EntireColumn.AutoFit
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Range("H:M,T:Y").Select
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With

'Returns screen updating, status bar, filtering to default states
With Application
    .ScreenUpdating = True
    .StatusBar = False
    .DisplayAlerts = True
End With

ActiveSheet.FilterMode = True

MsgBox "Completed macro.", vbInformation
Sheets("Extract").Select
Range("A1").Select

End Sub


'This macro is designed to confirm the module information from PMS PROD, such as U/W code & EPF indicator (ie "BC")
Private Sub PROD_Validate()
Dim policyErrorTest As String
Dim counterPecentage As Integer
Dim counter As Integer

counter = 0
errorCount = 0

'Calculates & counts number of records in excel spreadsheet
Range("A1", ActiveCell.SpecialCells(xlLastCell)).Select
lastRow = Selection.Rows.Count
numberOfPolicies = lastRow


'Reset used range in Excel
Application.ActiveSheet.UsedRange

For currentRow = 2 To lastRow 'start at row 2 and continue to lastrow
    counter = counter + 1
    counterPercentage = (counter / numberOfPolicies) * 100
    
    'Updates the status bar to provide visual feedback of progress
    Application.StatusBar = "Retrieving " & counter & " of " & numberOfPolicies
    
    symbol = ActiveSheet.Cells(currentRow, 2)
    number = ActiveSheet.Cells(currentRow, 3)
    
    'Prepares PMS for data entry  ---------------------------   Excel > PMS
    Call commandScreen(whichKey:="<home>")
    Call commandScreen(whichKey:="<clear>")

    'Sends current module into PIBC to confirm stats on MOD listed in ARL
    Call commandScreen(screenType:="pibc", symbol:=symbol, number:=number, whichKey:="<Enter>")

    'Tests to see if policy is in PROD
    policyErrorTest = screen.getstring(1, 54, 6)
    If policyErrorTest = "POLICY" Then
        policyError (currentRow)
    Else
        session.screen.SendKeys ("<Enter>")
        Do While system.ActiveSession.screen.OIA.Xstatus <> 0
                DoEvents
        Loop
        
        prodExtract (currentRow)
    
    End If

Next currentRow

End Sub



Private Sub policyError(currentRow As Integer)

errorCount = errorCount + 1

Call commandScreen(whichKey:="<home>")
Call commandScreen(whichKey:="<clear>")
Call commandScreen(screenType:="eibc", symbol:=symbol, number:=number, whichKey:="<Enter>")

policyErrorTest = screen.getstring(1, 54, 6)
If policyErrorTest = "POLICY" Then
    ActiveSheet.Cells(currentRow, 1).Interior.ColorIndex = 27
    ActiveSheet.Cells(currentRow, 2).Interior.ColorIndex = 27
    ActiveSheet.Cells(currentRow, 3).Interior.ColorIndex = 27
Else
    Call prodExtract(currentRow)
End If

Do While system.ActiveSession.screen.OIA.Xstatus <> 0
    DoEvents
Loop

End Sub



Private Sub prodExtract(currentRow As Integer)
Dim startDate As Date, endDate As Date

MOD_COUNT = 0

'Start data retrieval from column 'D'
Cells(currentRow, 3).Select

'Basic policy information - as per current module
ActiveCell.Offset(0, 1).Select
Agent = screen.getstring(3, 17, 7)
     ActiveCell = Agent
    
ActiveCell.Offset(0, 1).Select
PC = screen.getstring(3, 57, 2)
    ActiveCell = PC
    
ActiveCell.Offset(0, 1).Select
inspDist = screen.getstring(3, 48, 3)
    ActiveCell = inspDist

ActiveCell.Offset(0, 1).Select
branch = screen.getstring(3, 39, 2)
    ActiveCell = branch

Do
    'Sepecific module information
    ActiveCell.Offset(0, 1).Select
    PROD_MOD = screen.getstring(1, 19, 2)
        ActiveCell = PROD_MOD
    
    ActiveCell.Offset(0, 1).Select
    startDate = screen.getstring(2, 5, 6)
        ActiveCell = startDate
    
    ActiveCell.Offset(0, 1).Select
    endDate = screen.getstring(2, 5, 6)
        ActiveCell = endDate
    
    ActiveCell.Offset(0, 1).Select
    PDeb = screen.getstring(5, 66, 1)
        ActiveCell = PDeb
    
    ActiveCell.Offset(0, 1).Select
    UWCode = screen.getstring(3, 9, 1)
        ActiveCell = UWCode
    
    ActiveCell.Offset(0, 1).Select
    EDI = screen.getstring(4, 78, 2)
        ActiveCell = EDI
    
    PROD_MOD = PROD_MOD - 1
    MOD_COUNT = MOD_COUNT + 1
Loop Until PROD_MOD = "00" Or MOD_COUNT = 3

End Sub



Private Sub startPROD()
Dim logOnScreen As String, lockedScreen As String, PMS_MENU As String, PMS_MENU_DATE As String
Dim applicationScreen As String, loop_counter As Integer


Set system = CreateObject("EXTRA.System")
Set session = system.ActiveSession
Set screen = session.screen

    exitOption = False
    If ((system Is Nothing) Or (session Is Nothing) Or (screen Is Nothing)) Then
        If MsgBox("Could not start PMS. Login to PMS/PROD and when ready press 'OK'.", vbOKCancel + vbInformation, "PMS not ready") = vbCancel Then
            exitOption = True
        Else
            Call startPROD
        End If
    End If
    
    logOnScreen = screen.getstring(3, 16, 6)
    If (logOnScreen = "GGGGGG") Then
        If MsgBox("Login to PMS/PROD and when ready press 'OK'.", vbOKCancel + vbInformation, "PMS not ready") = vbCancel Then
            exitOption = True
        Else
            Call startPROD
        End If
    End If
    
    lockedScreen = screen.getstring(2, 25, 4)
    If (lockedScreen = "Term") Then
        If MsgBox("Login to PMS/PROD and when ready press 'OK'.", vbOKCancel + vbInformation, "PMS screen locked") = vbCancel Then
            exitOption = True
        Else
            Call startPROD
        End If
    End If
    
    PMS_MENU = screen.getstring(3, 27, 5)
    If PMS_MENU <> "CL/SU" Then
        Do
            PMS_MENU = screen.getstring(3, 43, 9)
            PMS_MENU_DATE = screen.getstring(4, 66, 4)
            logOnScreen = screen.getstring(1, 1, 4)
            
            If (PMS_MENU <> "Main Menu") And (PMS_MENU_DATE <> "Date") Then
                Call commandScreen(whichKey:="<Pf3>")
            End If
            
            loop_counter = loop_counter + 1
        Loop Until (PMS_MENU = "Main Menu" And PMS_MENU_DATE = "Date") Or loop_counter = 15 Or logOnScreen = "DFHA"
    End If
    
        Call commandScreen(screenKey:="S", screenRow:=11, screenCol:=2) 'Selects PROD environment
        Call commandScreen(whichKey:="<enter>")
        Call commandScreen(whichKey:="<clear>") 'Clears PROD screen, awaiting dataGrab module (see below)

End Sub


'====================================================================================
'This routine controls the interaction with PMS
Private Sub commandScreen(Optional screenType As String, Optional symbol As String, _
                          Optional number As String, Optional module As String, _
                          Optional whichKey As String, Optional screenKey As String, _
                          Optional screenRow As Integer, Optional screenCol As Integer)
        
        'Tells PROD which screen to go to, if required (ie PIBC, PISA, ERRU, etc)
        If screenType <> "" Then
            screen.putstring screenType, 1, 1
        Else 'do nothing
        End If
        
        'Tests is a policy number has been entered
        If symbol <> "" And number <> "" Then
            screen.putstring " ", 1, 5
                screen.putstring symbol, 1, 6
            screen.putstring " ", 1, 9
                screen.putstring number, 1, 10
            screen.putstring " ", 1, 17
                screen.putstring module, 1, 18
        Else 'do nothing
        End If
        
        'Takes row/column and a key command if provided
        If screenRow <> 0 And screenCol <> 0 Then
            screen.putstring screenKey, screenRow, screenCol
        Else 'do nothing
        End If
        
        'Submits required key press to PROD (ie Enter, Home, Clear, F keys)
        If whichKey <> "" Then
            session.screen.SendKeys (whichKey)
        Else 'do nothing
        End If
        
        'Loops until PMS is ready
        Do While system.ActiveSession.screen.OIA.Xstatus <> 0
            DoEvents
        Loop

End Sub
