'*************************************************************************
'  Program:  FleetRelGrowth                                              *
'   Author:  Robert Andrew Stevens                                       *
'     Date:  02/28/97                                                    *
'  Purpose:  Analize Fleet Reliability Growth data                       *
'*************************************************************************
 
Const MaxData As Integer = 3100 ' Maximum number of Reports
Const MaxVec As Integer = 750 ' Maximum number of elements in vectors
 
Type Report_Struct ' Report Data Structure
    Dat As String  ' failure Date (yy/mm/dd)
    Num As String  ' Report #
    Odo As Long    ' Odometer
    Veh As String  ' Veh #/VIN
    Own As String  ' Natural Owner
    Team As String ' Team: 1, 3, 4, 5, 7, 9, 20, 24, 27, 29
    Rev As String  ' Review Category: 1, 2, or 3
    Cod As String  ' Impact Code: 1, 2, 3, 4, 5 or NA
    Sum As String  ' Problem Summary
    Sta As String  ' Status: "O", "R", "C", "I"  or "N"
    Sub As String  ' Subsystem
    Com As String  ' Component
    Cau As String  ' Root Cause
    FEF As Double  ' FEF (0.XX)
    Fle As Boolean ' In fleet?
    Inc As Boolean ' failure?
    Rea As String  ' Reason why or why not an failure
    Acc As Double  ' Acceleration Factor (X.X)
End Type
Dim Report(1 To MaxData) As Report_Struct ' Report data
 
Type Proj_Struct   ' Projection Data Structure
    Dat As String  ' failure Date (yy/mm/dd)
    Num As String  ' Report # (allow for up to 10 same causes)
    Own As String  ' Natural Owner
    Team As String ' Team: 1, 3, 4, 5, 7, 9, 20, 24, 27, 29 and 209 (???)
    Sum As String  ' Problem Summary
    Sta As String  ' Status: "O", "R", "C", "I"  or "N"
    Sub As String  ' Subsystem
    Com As String  ' Component
    Cau As String  ' Root Cause
    FEF As Double  ' FEF (0.XX)
    Mil As Long    ' Fleet mileage on date
    N_i As Integer ' Number of times B-mode occurred
End Type
Dim Proj(1 To MaxData) As Proj_Struct ' Projection data
 
Type Uniq_Struct   ' Uniq Team and Subsystem Projection Data Structure
    Team As String ' Unique Teams
    Sub As String  ' Unique Subsystems
    Cur As Double  ' Current FP1000 for unique Team and Subsystem
    Pro As Double  ' Projected FP1000 for unique Team and Subsystem
End Type
Dim Uniq(1 To MaxVec) As Uniq_Struct ' Uniq data
 
Dim Num_Report As Integer ' Number of Report data records
Dim Num_Proj As Integer   ' Number of projection data records
Dim Num_Uniq As Integer   ' Number of unique Teams and Subsystems
Dim Num_Date As Integer   ' Number of unique dates
Dim Num_Vehs As Integer   ' Number of unique vehicles
Dim Num_Teams As Integer  ' Number of unique Teams
Dim Num_Subs As Integer   ' Number of unique Subsystems
Dim Num_Coms As Integer   ' Number of unique Components
'Dim Num_Caus As Integer   ' Number of unique Causes
 
Dim Inc_Vec(1 To MaxVec) As Integer ' # failures by Date
Dim Dat_Vec(1 To MaxVec) As String  ' Unique failure Dates (yy/mm/dd)
Dim Veh_Vec(1 To MaxVec) As String  ' Unique Vehicle names
Dim Team_Vec(1 To MaxVec) As String ' Unique Teams
Dim Sub_Vec(1 To MaxVec) As String  ' Unique Subsystems
Dim Com_Vec(1 To MaxVec) As String  ' Unique Components
'Dim Cau_Vec(1 To MaxVec) As String ' Unique Causes
Dim Mil_Vec(1 To MaxVec) As Long    ' Matrix of vehicle odometer readings on a given date
 
Dim ProjSheet As Worksheet
Dim Row As Integer ' Row Counter
 
Sub Main()
 
    Dim DBoxOK As Boolean
    Dim IntrDlg As DialogSheet
    Set IntrDlg = ThisWorkbook.DialogSheets("IntrDlg")
    Dim MainDlg As DialogSheet
    Set MainDlg = ThisWorkbook.DialogSheets("MainDlg")
 
    DBoxOK = IntrDlg.Show
    If Not DBoxOK Then Exit Sub
 
'   Loop until the main dialog is cancelled
    DBoxOK = True
    Do While DBoxOK
        DBoxOK = MainDlg.Show
        If Not DBoxOK Then End
    Loop
End Sub
 
'**************************************************************************
' function: Read_Raw                                                      *
'  purpose: read raw data spreadsheet and store in array                  *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Read_Raw()
 
    Set Raw = ThisWorkbook.Sheets("Raw")
    Set Raw1 = ThisWorkbook.Sheets("Raw1")
    Dim i As Integer ' Loop counter
 
    Raw1.Select
    Cells.Select
    Selection.ClearContents
 
'    Raw.Select
    Num_Report = Application.CountA(Raw.Range("A:A"))
 
    For i = 1 To Num_Report
        Report(i).Dat = Format(Raw.Cells(i, 1), "yy/mm/dd")
'        Raw1.Cells(i, 1) = Report(i).Dat
        Report(i).Num = Raw.Cells(i, 2)
'        Raw1.Cells(i, 2) = Report(i).Num
 
        If VarType(Raw.Cells(i, 3)) = 0 Then ' Empty
            Report(i).Odo = 0
        Else
            If VarType(Raw.Cells(i, 3)) = 5 Then ' Double
                Report(i).Odo = Raw.Cells(i, 3)
            Else
                If VarType(Raw.Cells(i, 3)) = 8 Then ' String
                    If Raw.Cells(i, 3) <> "" Then
                        Report(i).Odo = Val(Raw.Cells(i, 3))
                    Else
                        Report(i).Odo = 0
                    End If
                End If
            End If
        End If
'        Raw1.Cells(i, 3) = Report(i).Odo
 
        Report(i).Veh = Raw.Cells(i, 4)
'        Raw1.Cells(i, 4) = Report(i).Veh
        Report(i).Own = Left(Raw.Cells(i, 5), 15)
'        Raw1.Cells(i, 5) = Report(i).Own
        Report(i).Team = Raw.Cells(i, 6)
'        Raw1.Cells(i, 6) = Report(i).Team
        Report(i).Rev = Raw.Cells(i, 7)
'        Raw1.Cells(i, 7) = Report(i).Rev
        Report(i).Cod = Raw.Cells(i, 8)
'        Raw1.Cells(i, 8) = Report(i).Cod
        Report(i).Sum = Left(Raw.Cells(i, 9), 30)
'        Raw1.Cells(i, 9) = Report(i).Sum
        Report(i).Sta = Raw.Cells(i, 10)
'        Raw1.Cells(i, 10) = Report(i).Sta
        Report(i).Sub = Left(Raw.Cells(i, 11), 15)
'        Raw1.Cells(i, 11) = Report(i).Sub
        Report(i).Com = Left(Raw.Cells(i, 12), 15)
'        Raw1.Cells(i, 12) = Report(i).Com
        Report(i).Cau = Left(Raw.Cells(i, 13), 20)
'        Raw1.Cells(i, 13) = Report(i).Cau
        Report(i).FEF = Val(Raw.Cells(i, 14))
'        Raw1.Cells(i, 14) = Report(i).FEF
        Report(i).Fle = True ' initialize
'        Raw1.Cells(i, 15) = Report(i).Fle
        Report(i).Inc = False ' initialize
'        Raw1.Cells(i, 16) = Report(i).Inc
        Report(i).Acc = 1# ' initialize
'        Raw1.Cells(i, 17) = Report(i).Acc
    Next i
 
'    Print_Raw
 
    MsgBox ("Number of records read = " & Num_Report)
End Sub
 
'**************************************************************************
' function: Calc_Mile                                                     *
'  purpose: Calculate fleet mileage on a given day                        *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Calc_Mile()
 
    Dim i, j, k As Integer
    Dim Odometer(1 To MaxVec, 1 To MaxVec) As Long ' Matrix of vehicle odometer readings on a given date
 
    Set OutOrder = ThisWorkbook.Sheets("OutOrder")
    Set DateMile = ThisWorkbook.Sheets("DateMile")
 
    OutOrder.Select
    Cells.Select
    Selection.ClearContents
 
    DateMile.Select
    Cells.Select
    Selection.ClearContents
 
'   Create vectors of unique values of dates and vehicles (& Team, Subsystem, Component & Cause)
    Make_Uniq
 
'   Make Odometer matrix and fill it
 
'   Initialize Odometer matrix
    For i = 1 To Num_Date
        For j = 1 To Num_Vehs
            Odometer(i, j) = 0
            For k = 1 To Num_Report
                If Dat_Vec(i) = Report(k).Dat And Veh_Vec(j) = Report(k).Veh Then
                    Odometer(i, j) = Report(k).Odo
                End If
            Next k
        Next j
    Next i
 
'   Fill Odometer matrix
    For j = 1 To Num_Vehs
        For i = 1 To Num_Date - 1
            If Odometer(i + 1, j) = 0 Then
                Odometer(i + 1, j) = Odometer(i, j)
            End If
        Next i
    Next j
 
'   Sum across Odometer matrix
    For i = 1 To Num_Date
        Mil_Vec(i) = 0
        For j = 1 To Num_Vehs
            Mil_Vec(i) = Mil_Vec(i) + Odometer(i, j)
        Next j
    Next i
 
'   Check that odometer readings are in ascending order
    k = 1
    For j = 1 To Num_Vehs
        For i = 1 To Num_Date - 1
            If Odometer(i, j) > Odometer(i + 1, j) Then
                OutOrder.Cells(k, 1) = "Vehicle " & Veh_Vec(j) & " mileage is out-of-order on " & Dat_Vec(i + 1)
                k = k + 1
            End If
        Next i
    Next j
 
'   Write date and mileage to spreadsheet
    For i = 1 To Num_Date
        DateMile.Cells(i, 1) = Dat_Vec(i)
        DateMile.Cells(i, 2) = Mil_Vec(i)
    Next i
 
    MsgBox ("Fleet Mileage = " & Mil_Vec(Num_Date) & " on " & Dat_Vec(Num_Date))
End Sub
 
'**************************************************************************
' function: filter                                                        *
'  purpose: Read in fleet vehicles and acceleration factor                *
'           and set "fleet" and adjust mileage for each failure           *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Filter()
 
    Set Fleet = ThisWorkbook.Sheets("Fleet")
 
    Dim i, j As Integer ' Loop counters
    Dim Num_Flt As Integer ' Number of vehicles in fleet
    Dim Veh_Lst(1 To MaxVec) As String ' Vehicle list
    Dim Veh_Acc(1 To MaxVec) As Double ' Vehicle acceleration factor
 
'   Read Vehicle list and acceleration factors
    Num_Flt = Application.CountA(Fleet.Range("A:A"))
 
    For i = 1 To Num_Flt
        Veh_Lst(i) = Fleet.Cells(i, 1)
        Veh_Acc(i) = Val(Fleet.Cells(i, 2))
    Next i
 
    For i = 1 To Num_Report
        Report(i).Fle = False ' Initialize
        For j = 1 To Num_Flt
            If Report(i).Veh = Veh_Lst(j) Then
                Report(i).Fle = True
                Report(i).Odo = Report(i).Odo * Veh_Acc(j)
                Exit For
            End If
        Next j
    Next i
 
'    Print_Raw
    MsgBox ("Number of vehicles read = " & Num_Flt)
End Sub
 
'**************************************************************************
' function: calc_FP1000                                                   *
'  purpose: Calculate FP1000                                              *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Calc_FP1000()
 
    Dim i As Integer ' Loop counter
    Dim Num_Inc As Integer ' Number of failures
    Dim Cur_FP1000 As Double ' Current FP1000
 
    Det_Inc ' Determine whether Report is an "failure"
 
'   Determine number of failures
    Num_Inc = 0
    For i = 1 To Num_Report
        If Report(i).Fle And Report(i).Inc Then
            Num_Inc = Num_Inc + 1
        End If
    Next i
 
    Cur_FP1000 = FP1000((Num_Inc), Mil_Vec(Num_Date))
 
    MsgBox (Num_Inc & " failures occurred in " & Mil_Vec(Num_Date) & _
        " miles" & Chr(13) & "=> FP1000 = " & Application.Round(Cur_FP1000, 0))
End Sub
 
'**************************************************************************
' function: det_inc                                                       *
'  purpose: Determine whether Report counts as an failure                 *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Det_Inc()
 
    Dim i As Integer
 
    For i = 1 To Num_Report
        If Left(Report(i).Sta, 1) = "I" Or Left(Report(i).Sta, 1) = "X" Then
            Report(i).Inc = False
        Else
            If Left(Report(i).Cod, 1) = "5" Then
                Report(i).Inc = False
            Else
                If Report(i).Odo = 0 Then
                    Report(i).Inc = False
                Else: Report(i).Inc = True
                End If
            End If
        End If
    Next i
 
    Count_Inc
 
End Sub
 
'**************************************************************************
' function: Count_Inc                                                     *
'  purpose: Count # of failures by date                                   *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Count_Inc()
 
    Dim i, j As Integer
    Dim Miles(1 To MaxVec) As Long
    Dim Num_Inc As Integer
'    Set DateMile = ThisWorkbook.Sheets("DateMile")
    Set failure = ThisWorkbook.Sheets("failure")
 
'   Initialize # failures vector
'    For i = 1 To Num_Date
'        Inc_Vec(i) = 0
'    Next i
 
    Num_Inc = 0
    For i = 1 To Num_Report
        If Report(i).Inc = True Then
            For j = 1 To Num_Date
                If Report(i).Dat = Dat_Vec(j) Then
'                   Inc_Vec(j) = Inc_Vec(j) + 1
 
                    Num_Inc = Num_Inc + 1
                    Miles(Num_Inc) = Mil_Vec(j)
                End If
            Next j
        End If
    Next i
 
'   Write # failures to spreadsheet
    For i = 1 To Num_Inc
        failure.Cells(i, 1) = Miles(i)
        failure.Cells(i, 2) = 1
    Next i
 
 
End Sub
 
'**************************************************************************
' function: proj_FP1000                                                   *
'  purpose: Project FP1000                                                *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Proj_FP1000()
 
    Set ProjSheet = ThisWorkbook.Sheets("Project")
 
    Fill_Proj
    Group_Proj
 
    Row = 0 ' Initilize Row Counter and clean-up ProjSheet
    ProjSheet.Select
    Cells.Select
    Selection.ClearContents
    Selection.PageBreak = xlNone
 
    Proj_Veh
    Proj_Sub
 
    MsgBox ("Projection data are printed in tab 'Project'" & Chr(13) & _
            "FP1000 Summary is printed in tab 'Summary'    " & Chr(13) & _
            "Sorted FP1000 List is printed in tab 'Ranking'")
End Sub
 
'**************************************************************************
' function: Fill_Proj                                                     *
'  purpose: Match Report to Team, Subsystem & Component and fill Proj     *
'           data structure                                                *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Fill_Proj()
 
    Dim a As Integer ' Loop counter for Report data structure
    Dim i As Integer ' Loop counter for Team
    Dim j As Integer ' Loop counter for Subsystem
    Dim k As Integer ' Loop counter for Component
 
    Num_Proj = 0
    For a = 1 To Num_Report
        If Report(a).Fle And Report(a).Inc Then
            For i = 1 To Num_Teams
                For j = 1 To Num_Subs
                    For k = 1 To Num_Coms
                        If Team_Vec(i) = Report(a).Team And _
                           Sub_Vec(j) = Report(a).Sub And _
                           Com_Vec(k) = Report(a).Com Then
 
                            Num_Proj = Num_Proj + 1
                            Call Copy_Proj(a, Num_Proj)
                        End If
                    Next k
                Next j
            Next i
        End If
    Next a
End Sub
 
'**************************************************************************
' function: Copy_Proj                                                     *
'  purpose: Copy Report record to Proj data structure                     *
'    input: a = Report record postion, b = Proj record postion            *
'   output:                                                               *
'**************************************************************************
Sub Copy_Proj(a As Integer, b As Integer)
    Proj(b).Dat = Report(a).Dat
    Proj(b).Num = Report(a).Num
    Proj(b).Own = Report(a).Own
    Proj(b).Team = Report(a).Team
    Proj(b).Sum = Report(a).Sum
    Proj(b).Sta = Report(a).Sta
    Proj(b).Sub = Report(a).Sub
    Proj(b).Com = Report(a).Com
    Proj(b).Cau = Report(a).Cau
    Proj(b).FEF = Report(a).FEF
    Proj(b).Mil = Mile(Report(a).Dat)
    Proj(b).N_i = 1 ' Initialize
End Sub
 
'**************************************************************************
' function: Group_Proj                                                    *
'  purpose: Group together same root causes                               *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Group_Proj()
 
    Dim i, j, k As Integer ' Loop counters
    Dim Tmp_Proj As Integer ' Tempory storage of Num_Proj
 
    Tmp_Proj = Num_Proj
 
    For i = 1 To Num_Proj - 1
        j = i + 1
        Do While j <= Num_Proj
            If Proj(i).Team = Proj(j).Team And _
               Proj(i).Sub = Proj(j).Sub And _
               Proj(i).Com = Proj(j).Com And _
               Proj(i).Cau = Proj(j).Cau And _
               Proj(i).Sta = Proj(j).Sta And _
               Proj(i).FEF = Proj(j).FEF And _
               Not Proj(i).Cau = "" Then
 
'               Combine Report numbers
                Proj(i).Num = Proj(i).Num & " " & Proj(j).Num
 
                If Proj(i).Mil > Proj(j).Mil Then ' Go with earlier date and mileage
                    Proj(i).Dat = Proj(j).Dat
                    Proj(i).Mil = Proj(j).Mil
                End If
 
                Proj(i).N_i = Proj(i).N_i + 1 ' Increment N_i
 
                For k = j To Num_Proj
                    Call Move_Proj(k) ' Move data up 1 postion to replace repeat cause
                Next k
                Tmp_Proj = Tmp_Proj - 1 ' Decrement number of elements in proj
            Else: j = j + 1
            End If
        Loop
    Next i
 
    Num_Proj = Tmp_Proj
End Sub
 
'**************************************************************************
' function: Move_Proj                                                     *
'  purpose: Move Proj record up one postion (overwritting previous record)*
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Move_Proj(j As Integer)
    Proj(j).Num = Proj(j + 1).Num
    Proj(j).Dat = Proj(j + 1).Dat
    Proj(j).Own = Proj(j + 1).Own
    Proj(j).Team = Proj(j + 1).Team
    Proj(j).Sum = Proj(j + 1).Sum
    Proj(j).Sta = Proj(j + 1).Sta
    Proj(j).Sub = Proj(j + 1).Sub
    Proj(j).Com = Proj(j + 1).Com
    Proj(j).Cau = Proj(j + 1).Cau
    Proj(j).Mil = Proj(j + 1).Mil
    Proj(j).FEF = Proj(j + 1).FEF
    Proj(j).N_i = Proj(j + 1).N_i
End Sub
 
'**************************************************************************
' function: Veh_Proj                                                      *
'  purpose: Calculate and print vehicle projection                        *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Proj_Veh()
 
    Dim Num_A As Integer ' Number of A-modes failures
    Dim Num_B As Integer ' Number of B-modes failures
    Dim Mod_B As Integer ' Number of distinct B-modes
    Dim Adj_Inc As Double ' Adjusted number of failures
    Dim Sum_FEF As Double ' Sum of FEFs
    Dim Sum_lnX As Double ' Sum of log(failure time)
    Dim Project As Double ' Vehicle FP1000 Projection
 
    Set ProjSheet = ThisWorkbook.Sheets("Project")
 
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Vehicle Current and Projected FP1000"
    Row = Row + 1
 
    For i = 1 To Num_Proj
        If Proj(i).FEF > 0# Then
            Num_B = Num_B + Proj(i).N_i
            Sum_FEF = Sum_FEF + (Proj(i).N_i * Proj(i).FEF)
            Sum_lnX = Sum_lnX + Log(Proj(i).Mil)
            Adj_Inc = Adj_Inc + Proj(i).N_i * (1 - Proj(i).FEF)
            Mod_B = Mod_B + 1
        Else: Num_A = Num_A + Proj(i).N_i
        End If
    Next i
 
    Project = Calc_Proj(Num_A, Num_B, Adj_Inc, Sum_FEF, Sum_lnX, Mod_B)
End Sub
 
'**************************************************************************
' function: Sub_Proj                                                      *
'  purpose: Calculate and print subsystem projections                     *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Proj_Sub()
 
    Dim i As Integer ' Loop counter
    Dim Num_A As Integer ' Number of A-modes failures
    Dim Num_B As Integer ' Number of B-modes failures
    Dim Mod_B As Integer ' Number of distinct B-modes
    Dim Sum_FEF As Double ' Sum of FEFs
    Dim Adj_Inc As Double ' Adjusted number of failures
    Dim Sum_lnX As Double ' Sum of log(failure time)
 
    Set ProjSheet = ThisWorkbook.Sheets("Project")
 
    Uni_Comb
 
    For i = 1 To Num_Uniq
 
        Num_A = 0 ' Initialize
        Num_B = 0 ' Initialize
        Mod_B = 0 ' Initialize
        Sum_FEF = 0# ' Initialize
        Sum_lnX = 0# ' Initialize
        Adj_Inc = 0# ' Initialize
 
        Row = Row + 1
        ProjSheet.Cells(Row, 2) = "Summary for Team " & Uniq(i).Team & " Subsystem " & Uniq(i).Sub
 
        'Insert page break
        ProjSheet.Select
        Rows(Row).Select
        ActiveCell.PageBreak = xlManual
 
        Row = Row + 1
 
        Row = Row + 1
        ProjSheet.Cells(Row, 1) = "Report"
        ProjSheet.Cells(Row, 2) = "Date"
        ProjSheet.Cells(Row, 3) = "Status"
        ProjSheet.Cells(Row, 4) = "FEF"
        ProjSheet.Cells(Row, 5) = "#"
        ProjSheet.Cells(Row, 6) = "Component"
        ProjSheet.Cells(Row, 7) = "Cause"
        ProjSheet.Cells(Row, 8) = "Owner"
        ProjSheet.Cells(Row, 9) = "Problem"
 
        For j = 1 To Num_Proj
            If Uniq(i).Team = Proj(j).Team And _
               Uniq(i).Sub = Proj(j).Sub Then
 
                Row = Row + 1
                ProjSheet.Cells(Row, 1) = Proj(j).Num
                ProjSheet.Cells(Row, 2) = Proj(j).Dat
                ProjSheet.Cells(Row, 3) = Proj(j).Sta
                ProjSheet.Cells(Row, 4) = Proj(j).FEF
                ProjSheet.Cells(Row, 5) = Proj(j).N_i
                ProjSheet.Cells(Row, 6) = Proj(j).Com
                ProjSheet.Cells(Row, 7) = Proj(j).Cau
                ProjSheet.Cells(Row, 8) = Proj(j).Own
                ProjSheet.Cells(Row, 9) = Proj(j).Sum
 
                If Proj(j).FEF > 0# Then
                    Num_B = Num_B + Proj(j).N_i
                    Sum_FEF = Sum_FEF + (Proj(j).N_i * Proj(j).FEF)
                    Sum_lnX = Sum_lnX + Log(Proj(j).Mil)
                    Adj_Inc = Adj_Inc + Proj(j).N_i * (1 - Proj(j).FEF)
                    Mod_B = Mod_B + 1
                Else: Num_A = Num_A + Proj(j).N_i
                End If
            End If
        Next j
 
        Row = Row + 1
 
        Uniq(i).Cur = FP1000((Num_A + Num_B), Mil_Vec(Num_Date))
        Uniq(i).Pro = Calc_Proj(Num_A, Num_B, Adj_Inc, Sum_FEF, Sum_lnX, Mod_B)
    Next i
 
    Call Print_Proj("Summary")
    Sort_Proj
    Call Print_Proj("Ranking")
End Sub
 
'**************************************************************************
' function: Uni_Comb                                                      *
'  purpose: Determine unique combinations of Team and Subsystem           *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Uni_Comb()
 
    Dim i, j As Integer ' Loop counters
    Dim Tmp_Team As String
    Dim Tmp_Sub As String
 
'   Fill Uniq data structure
    Num_Uniq = 0
    For i = 1 To Num_Proj
        Num_Uniq = Num_Uniq + 1
        Uniq(i).Team = Proj(i).Team
        Uniq(i).Sub = Proj(i).Sub
    Next i
 
'   Sort Uniq data structure by Team
    For i = 1 To Num_Uniq - 1
        For j = i + 1 To Num_Uniq
            If Uniq(i).Team > Uniq(j).Team Then
                Tmp_Team = Uniq(i).Team
                Tmp_Sub = Uniq(i).Sub
                Uniq(i).Team = Uniq(j).Team
                Uniq(i).Sub = Uniq(j).Sub
                Uniq(j).Team = Tmp_Team
                Uniq(j).Sub = Tmp_Sub
            End If
        Next j
    Next i
 
'   Sort Uniq data structure by Subsystem (within Team)
    For i = 1 To Num_Uniq - 1
        For j = i + 1 To Num_Uniq
            If Uniq(i).Team = Uniq(j).Team And Uniq(i).Sub > Uniq(j).Sub Then
                Tmp_Team = Uniq(i).Team
                Tmp_Sub = Uniq(i).Sub
                Uniq(i).Team = Uniq(j).Team
                Uniq(i).Sub = Uniq(j).Sub
                Uniq(j).Team = Tmp_Team
                Uniq(j).Sub = Tmp_Sub
            End If
        Next j
    Next i
 
'   Delete duplicate entries
    i = 2
    Do While i <= Num_Uniq
        If Uniq(i).Team = Uniq(i - 1).Team And Uniq(i).Sub = Uniq(i - 1).Sub Then
            If Not i = Num_Uniq Then
                For j = i To Num_Uniq
                    Uniq(j).Team = Uniq(j + 1).Team
                    Uniq(j).Sub = Uniq(j + 1).Sub
                Next j
            End If
            Num_Uniq = Num_Uniq - 1
        Else: i = i + 1
        End If
    Loop
End Sub
 
'**************************************************************************
' function: Sort_Proj                                                     *
'  purpose: Sort Subsystem Projection Summary by Projected FP1000         *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Sort_Proj()
 
    Dim i, j As Integer ' Loop counters
    Dim Tmp_Team As String ' Tempory Team
    Dim Tmp_Sub As String ' Tempory Subsystem
    Dim Tmp_Cur As Double ' Tempory current FP1000
    Dim Tmp_Pro As Double ' Tempory projected FP1000
 
    For i = 1 To Num_Uniq - 1
        For j = i + 1 To Num_Uniq
            If Uniq(i).Pro < Uniq(j).Pro Then ' Swap postions
 
                Tmp_Team = Uniq(j).Team
                Tmp_Sub = Uniq(j).Sub
                Tmp_Cur = Uniq(j).Cur
                Tmp_Pro = Uniq(j).Pro
 
                Uniq(j).Team = Uniq(i).Team
                Uniq(j).Sub = Uniq(i).Sub
                Uniq(j).Cur = Uniq(i).Cur
                Uniq(j).Pro = Uniq(i).Pro
 
                Uniq(i).Team = Tmp_Team
                Uniq(i).Sub = Tmp_Sub
                Uniq(i).Cur = Tmp_Cur
                Uniq(i).Pro = Tmp_Pro
            End If
        Next j
    Next i
End Sub
 
'**************************************************************************
' function: calc_proj                                                     *
'  purpose: calculate projection data                                     *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Function Calc_Proj(Num_A As Integer, Num_B As Integer, Adj_Inc As Double, _
    Sum_FEF As Double, Sum_lnX As Double, Mod_B As Integer) As Double
 
    Dim AvgFEF As Double ' Average FEF
    Dim Beta_b As Double ' Biased Beta
    Dim Beta_u As Double ' Unbiased Beta
    Dim HazRat As Double ' Hazard Rate
    Dim A_FP1000 As Double ' A-mode FP1000
    Dim B_FP1000 As Double ' B-mode FP1000
    Dim Adj_FP1000 As Double ' Adjusted FP1000 (for B-modes)
    Dim Cor_FP1000 As Double ' Corrected FP1000 (based on correction term)
 
    If Mod_B > 0 Then
        AvgFEF = Sum_FEF / Num_B
        If (Mod_B * Log(Mil_Vec(Num_Date)) - Sum_lnX) <> 0 Then
            Beta_b = Mod_B / (Mod_B * Log(Mil_Vec(Num_Date)) - Sum_lnX)
            Beta_u = (Mod_B - 1) * Beta_b / Mod_B
            HazRat = Mod_B * Beta_u / Mil_Vec(Num_Date)
        Else
            Beta_b = 0#
            Beta_u = 0#
            HazRat = 0#
        End If
    Else
        AvgFEF = 0#
        Beta_b = 0#
        Beta_u = 0#
        HazRat = 0#
    End If
 
    Call Print_Data(Num_A, Num_B, AvgFEF, Beta_b, Beta_u, HazRat)
 
    A_FP1000 = FP1000((Num_A), Mil_Vec(Num_Date))
    B_FP1000 = FP1000((Num_B), Mil_Vec(Num_Date))
    Adj_FP1000 = FP1000((Adj_Inc), Mil_Vec(Num_Date))
    Cor_FP1000 = 1000 * AvgFEF * HazRat * 10000
 
    Call Print_FP1000(A_FP1000, B_FP1000, Adj_FP1000, Cor_FP1000)
 
    Calc_Proj = A_FP1000 + Adj_FP1000 + Cor_FP1000
End Function
 
'*************************************************************************
'  Function:  FP1000()                                                   *
'   Purpose:  Calculate FP1000 at 1 Year given failures and time         *
'    Inputs:  failures, time                                             *
'    Return:  FP1000                                                     *
'*************************************************************************
Function FP1000(failures As Double, Time As Long) As Double
    FP1000 = 1000 * (failures / Time) * 10000
End Function
 
'**************************************************************************
' function: Print_Data                                                    *
'  purpose: Print Data for Current & Projected FP1000                     *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Print_Data(Num_A As Integer, Num_B As Integer, AvgFEF As Double, _
    Beta_b As Double, Beta_u As Double, HazRat As Double)
 
    Set ProjSheet = ThisWorkbook.Sheets("Project")
 
    Row = Row + 2
    ProjSheet.Cells(Row, 2) = "Data Used for Projections"
    Row = Row + 2
    ProjSheet.Cells(Row, 2) = "No. failures  = " & Num_A + Num_B
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Number A-Modes = " & Num_A
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Number B-Modes = " & Num_B
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Avg FEF        = " & Application.Round(AvgFEF, 6)
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Biased Beta    = " & Application.Round(Beta_b, 6)
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Unbiased Beta  = " & Application.Round(Beta_u, 6)
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "Hazard Rate    = " & Application.Round(HazRat, 6)
End Sub
 
'**************************************************************************
' function: Prnt_FP1000                                                   *
'  purpose: Print FP1000                                                  *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Print_FP1000(A_Mode As Double, B_Mode As Double, Adjust As Double, _
    Correct As Double)
 
    Set ProjSheet = ThisWorkbook.Sheets("Project")
 
    Row = Row + 2
    ProjSheet.Cells(Row, 3) = "FP1000 Summary"
    Row = Row + 2
    ProjSheet.Cells(Row, 3) = "Current  "
    ProjSheet.Cells(Row, 4) = "Projected"
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "A-Modes "
    ProjSheet.Cells(Row, 3) = Application.Round(A_Mode, 0)
    ProjSheet.Cells(Row, 4) = Application.Round(A_Mode, 0)
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "B-Modes "
    ProjSheet.Cells(Row, 3) = Application.Round(B_Mode, 0)
    ProjSheet.Cells(Row, 4) = Application.Round(Adjust, 0)
    Row = Row + 1
    ProjSheet.Cells(Row, 2) = "C-Modes "
    ProjSheet.Cells(Row, 3) = 0
    ProjSheet.Cells(Row, 4) = Application.Round(Correct, 0)
    Row = Row + 1
    ProjSheet.Cells(Row, 3) = "---------"
    ProjSheet.Cells(Row, 4) = "---------"
    Row = Row + 1
    ProjSheet.Cells(Row, 3) = Application.Round(A_Mode + B_Mode, 0)
    ProjSheet.Cells(Row, 4) = Application.Round(A_Mode + Adjust + Correct, 0)
End Sub
 
'*************************************************************************
'  Function:  mile()                                                     *
'   Purpose:  Find fleet mileage on a given date                         *
'    Inputs:  date                                                       *
'    Return:  mileage                                                    *
'*************************************************************************
Function Mile(DateStr As String) As Long
 
    Dim i As Integer ' Loop counter
    Dim pos As Integer ' Postion of desired date
 
    For i = 1 To Num_Date
        If Dat_Vec(i) = DateStr Then
            pos = i
            Exit For
        End If
    Next i
 
    Mile = Mil_Vec(pos)
End Function
 
'**************************************************************************
' function: Make_Uniq                                                     *
'  purpose: Create vectors of unique values                               *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Make_Uniq()
 
'   Initialize counter variables to 0
    Num_Date = 0
    Num_Vehs = 0
    Num_Teams = 0
    Num_Subs = 0
    Num_Coms = 0
    Num_Caus = 0
 
'   Get values from Report database and increment counter variables
    For i = 1 To Num_Report
        If Report(i).Fle Then
            Num_Date = Num_Date + 1
            Num_Vehs = Num_Vehs + 1
            Num_Teams = Num_Teams + 1
            Num_Subs = Num_Subs + 1
            Num_Coms = Num_Coms + 1
'            Num_Caus = Num_Caus + 1
            Dat_Vec(Num_Date) = Report(i).Dat
            Veh_Vec(Num_Vehs) = Report(i).Veh
            Team_Vec(Num_Teams) = Report(i).Team
            Sub_Vec(Num_Subs) = Report(i).Sub
            Com_Vec(Num_Coms) = Report(i).Com
'            Cau_Vec(Num_Caus) = Report(i).Cau
        End If
    Next i
 
    Call Uniq_List(Dat_Vec, Num_Date, "Dates")
    Call Uniq_List(Veh_Vec, Num_Vehs, "Vehicles")
    Call Uniq_List(Team_Vec, Num_Teams, "Teams")
    Call Uniq_List(Sub_Vec, Num_Subs, "Subsystems")
    Call Uniq_List(Com_Vec, Num_Coms, "Components")
'    Call Uniq_List(Cau_Vec, Num_Caus, "Causes")
 
End Sub
 
'**************************************************************************
' function: Uniq_List                                                     *
'  purpose: Create list of unique values                                  *
'    input: List name, number of elements, and location to print          *
'   output:                                                               *
'**************************************************************************
Sub Uniq_List(List() As String, NList As Integer, Location As String)
 
    Dim i As Integer
 
    Call Sort_List(List, NList)
    Call Del_Rep(List, NList)
    Call Print_List(List, NList, Location)
End Sub
 
'*************************************************************************
'  Function:  sort_list                                                  *
'   Purpose:  sort list of strings                                       *
'    Inputs:  list name and number of elements                           *
'    Return:                                                             *
'*************************************************************************
Sub Sort_List(List() As String, NList As Integer)
 
    Dim i, j As Integer ' Loop counters
    Dim Tmp_Str As String
 
    For i = 1 To NList - 1
        For j = i + 1 To NList
            If List(i) > List(j) Then
                Tmp_Str = List(i)
                List(i) = List(j)
                List(j) = Tmp_Str
            End If
        Next j
    Next i
End Sub
 
'*************************************************************************
'  Function:  del_rep                                                    *
'   Purpose:  delete repeats in a list of strings                        *
'    Inputs:  list name and number of elements                           *
'    Return:                                                             *
'*************************************************************************
Sub Del_Rep(List() As String, NList As Integer)
 
    Dim i, j As Integer ' Loop counters
 
    i = 2
    Do While i <= NList
        If List(i) = List(i - 1) Then
            If Not i = NList Then
                For j = i To NList
                    List(j) = List(j + 1)
                Next j
            End If
            NList = NList - 1
        Else: i = i + 1
        End If
    Loop
End Sub
 
'*************************************************************************
'  Function:  Print_list                                                 *
'   Purpose:  Print list to a sheet                                      *
'    Inputs:  list name, number of elements, and location to print       *
'    Return:                                                             *
'*************************************************************************
Sub Print_List(List() As String, NList As Integer, Location As String)
 
    Dim i As Integer ' Loop counter
 
    Sheets(Location).Select
    Cells.Select
    Selection.ClearContents
    For i = 1 To NList
        Sheets(Location).Cells(i, 1) = List(i)
    Next i
End Sub
 
'**************************************************************************
' function: Print_Proj                                                    *
'  purpose: Print Subsystem Projection Summary                            *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Print_Proj(Location As String)
 
    Dim i As Integer ' Loop counter
 
    Set OutSheet = ThisWorkbook.Sheets(Location)
    OutSheet.Select
    Cells.Select
    Selection.ClearContents
 
    Row = 1
    OutSheet.Cells(Row, 1) = "#"
    OutSheet.Cells(Row, 2) = "Team"
    OutSheet.Cells(Row, 3) = "Subsystem"
    OutSheet.Cells(Row, 4) = "Current"
    OutSheet.Cells(Row, 5) = "Project"
 
    Row = 2
    OutSheet.Cells(Row, 1) = "--"
    OutSheet.Cells(Row, 2) = "---"
    OutSheet.Cells(Row, 3) = "---------------"
    OutSheet.Cells(Row, 4) = "-------"
    OutSheet.Cells(Row, 5) = "-------"
 
    For i = 1 To Num_Uniq
        Row = Row + 1
        OutSheet.Cells(Row, 1) = i
        OutSheet.Cells(Row, 2) = Uniq(i).Team
        OutSheet.Cells(Row, 3) = Uniq(i).Sub
        OutSheet.Cells(Row, 4) = Application.Round(Uniq(i).Cur, 0)
        OutSheet.Cells(Row, 5) = Application.Round(Uniq(i).Pro, 0)
    Next i
End Sub
 
Sub Prnt_Proj()
    Set Output = ThisWorkbook.Sheets("Proj")
    Dim b As Integer ' Loop counter
 
    Output.Select
    Cells.Select
    Selection.ClearContents
 
    For b = 1 To Num_Proj
        Output.Cells(b, 1) = Proj(b).Dat
        Output.Cells(b, 2) = Proj(b).Num
        Output.Cells(b, 3) = Proj(b).Own
        Output.Cells(b, 4) = Proj(b).Team
        Output.Cells(b, 5) = Proj(b).Sum
        Output.Cells(b, 6) = Proj(b).Sta
        Output.Cells(b, 7) = Proj(b).Sub
        Output.Cells(b, 8) = Proj(b).Com
        Output.Cells(b, 9) = Proj(b).Cau
        Output.Cells(b, 10) = Proj(b).FEF
        Output.Cells(b, 11) = Proj(b).Mil
        Output.Cells(b, 12) = Proj(b).N_i
    Next b
End Sub
 
'*************************************************************************
'  Function:  Print_Raw                                                  *
'   Purpose:  Print list to a sheet                                      *
'    Inputs:  list name, number of elements, and location to print       *
'    Return:                                                             *
'*************************************************************************
Sub Print_Raw()
 
    Set Raw = ThisWorkbook.Sheets("Raw1")
    Dim i As Integer ' Loop counter
 
    Raw.Select
    Cells.Select
    Selection.ClearContents
 
    For i = 1 To Num_Report
        Raw.Cells(i, 1) = Report(i).Dat
        Raw.Cells(i, 2) = Report(i).Num
        Raw.Cells(i, 3) = Report(i).Odo
        Raw.Cells(i, 4) = Report(i).Veh
        Raw.Cells(i, 5) = Report(i).Own
        Raw.Cells(i, 6) = Report(i).Team
        Raw.Cells(i, 7) = Report(i).Rev
        Raw.Cells(i, 8) = Report(i).Cod
        Raw.Cells(i, 9) = Report(i).Sum
        Raw.Cells(i, 10) = Report(i).Sta
        Raw.Cells(i, 11) = Report(i).Sub
        Raw.Cells(i, 12) = Report(i).Com
        Raw.Cells(i, 13) = Report(i).Cau
        Raw.Cells(i, 14) = Report(i).FEF
        Raw.Cells(i, 15) = Report(i).Fle
        Raw.Cells(i, 16) = Report(i).Inc
        Raw.Cells(i, 17) = Report(i).Acc
    Next i
End Sub
 
Sub HelpButton()
    Dim HelpDlg As DialogSheet
    Set HelpDlg = ThisWorkbook.DialogSheets("HelpDlg")
 
    HelpDlg.Show
End Sub
