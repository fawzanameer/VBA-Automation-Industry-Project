# Excel VBA Automation - Industry Project

## Project Description

This project represents a real-world industry task completed as part of my VBA module. The challenge involved efficiently copying and pasting a single asterisk cell into a designated blue range situated just above it.

## Demonstration
You can watch a demonstration of this project on YouTube: [Click here](https://www.youtube.com/watch?v=LORn-SDDDoY)

## VBA Code Used


    Sub haridwar()

    Worksheets("Test").Activate

    Range("b1").Select

    Dim i As Integer


    actvcell: ActiveCell.Select

    For i = 1 To 15
    

        If Left(ActiveCell.Offset(i, 0), 11) = "*     Total" Then
            Exit For
        
        ElseIf Left(ActiveCell.Offset(i, 0), 2) = "* " Then
            
            'Selects the cell to copy
            ActiveCell.Offset(i, 0).Select
            ActiveCell.Copy
            
            'Count of Blue Range
            x = WorksheetFunction.CountIf(Range(ActiveCell, ActiveCell.Offset(-i, 0)), "      4*") + _
                WorksheetFunction.CountIf(Range(ActiveCell, ActiveCell.Offset(-i, 0)), "      6*")
            
            ActiveCell.Offset(-1, 1).Select
            Range(ActiveCell, ActiveCell.Offset(-x + 1, 0)).PasteSpecial xlPasteAll
            ActiveCell.Offset(x, -1).Select
            GoTo actvcell
            
        End If
        Next i

    Application.CutCopyMode = False


    End Sub

