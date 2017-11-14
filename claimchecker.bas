Sub ClaimChecker()

'
' ClaimChecker Macro
'Last updated Josephine Choi [Mar 17th 2015]
'Use this to format the Claim Checker (downloaded from EBSCONET)

Dim LR As Long

LR = Cells(Rows.Count, 1).End(xlUp).Row

'SAFETY: Just in case the file is not a claimchecker. Basically, checking to see if field E1 is "Claim Date"

If Range("E1").Value = "Claim Date" Then
'
    'Hide Column A and C
    
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    
    ' Change the width
    Columns("B:B").ColumnWidth = 35
    Columns("H:H").ColumnWidth = 18.14
    Columns("G:G").ColumnWidth = 14.71
    Columns("E:E").ColumnWidth = 14.57
    Columns("I:I").ColumnWidth = 15
    
    'Add Text Wrap
    
    Range("B1:I" & LR).Select
    
    
    With Selection
     
        .WrapText = True
      
    End With
    
    ' Add Sort
    
    
    Range("B1").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B2:B" & LR) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:I" & LR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Change the view to Print Layout View
    ActiveWindow.View = xlPageLayoutView

       

'
  '
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Claim Checker" & Date
        
       
        .RightHeader = "Page &P/&N"
        '
       .Orientation = xlLandscape

        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
  
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
     
    End With
    Application.PrintCommunication = True
    
Else
Exit Sub

End If


End Sub

