Sub aresDeletionJO()
'
' aresDeletion Macro; modified for JO.

'
If Range("A1").Value = "Item ID" Then

'The below makes sure that the page number is recognized as text format


    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove


Columns("AI:AI").Select
    Selection.NumberFormat = "@"
    
    'This is to delete the columns not needed
    
'
    Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:AS").Select
    Selection.Delete Shift:=xlToLeft
    Range("I5").Select
    ActiveWindow.ScrollColumn = 1
    
    'this is the part that merge author, title etc.
    
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[2],"" | "",RC[3],"" | "",RC[4],"" | "",RC[5],"" | "",RC[6],"" | "",RC[7],"" | "",RC[8],"" | "",RC[9])"
        
        
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:= _
        xlFillDefault
        
        
     Do
       
    ActiveCell.Range("A1:A2").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:= _
        xlFillDefault
        
        Loop Until ActiveCell.Offset(1, -1) = ""
        
        'there's the end of the loop. Below is to do the paste value
        
       
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-6
    
    'below is to clean up the author, title field etc
    
    Columns("D:K").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    'final step: to select the fields needed and to copy
    
    
    Range("A2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
        
Else
Exit Sub


        
        End If
        
        
End Sub


 


