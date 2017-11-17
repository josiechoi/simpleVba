Sub ForeignUnitAndCurr()
'
' ForeignUnitAndCurr Macro

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'

' Use TextToColumn to Split Foreign Unit and Foreign Curr from Note

    Columns("K:K").Select
    Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="\", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
'Use TextToColumn to separate Foreign Curr from Foreign Unit
  
    Columns("L:L").Select
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlFixedWidth, _
        OtherChar:="\", FieldInfo:=Array(Array(0, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    
    
'Name cell L1 as "Foreign Unit"
  
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Foreign Unit"
    
'Name cell M1 as "Foreign Curr"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Foreign Curr"
    
'Format Column M (Foreign Curr) as Currency
    Columns("M:M").Select
    Selection.NumberFormat = "#,##0.00"
    
'Format Column F (Amount Paid) as Currency
    
    Columns("F:F").Select
    Selection.NumberFormat = "$#,##0.00"
    
'Add Column N for fiscal year

    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Fiscal Year"
    

'Add 2017-18 as Fiscal Year
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "2017-18"
    
     Range("N2").Select
    Selection.Copy
    
    Range("N2:N" & LastRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    
End Sub
