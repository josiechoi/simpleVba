Sub fixthedate()
Dim FinalCol As Long
Dim FinalRow As Long
Dim c As Long
Dim T As Long


' Use this macro to fixthedate. Use this to fix when the invoice is lumped together with the note field


FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

FinalCol = Cells(1, Columns.Count).End(xlToLeft).Column

Cells(3, 2).Offset(0, 10).Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.Select
    Selection.TextToColumns Destination:=ActiveCell.Offset(0, 0), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 3)), TrailingMinusNumbers:=True
        
        
       T = FinalCol / 10
       
Do Until T = c
On Error GoTo Skip
    ActiveCell.Offset(0, c + 10).Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.Select
  Selection.TextToColumns Destination:=ActiveCell.Offset(0, 0), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", FieldInfo:=Array(Array(1, 1), Array(2, 3)), TrailingMinusNumbers:=True


Loop

Skip:
Dim i As Long


Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(2, 1).Select


'
' Order2 Macro (This is the original working)

'i=i+1


'
For i = 2 To LastRow

'If Range("L" & i).Value <> blank Then


'LOOP SHOULD BEGIN HERE

 Do While Range("L" & i).Value <> blank

'This is to add a row below

Cells(i, 1).Offset(1, 0).Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'Cells(i, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    Cells(i, 1).Select
    
    
'This is to cut
    
    ActiveCell.Offset(0, 11).Range("A1:I1").Select
    Selection.Cut
    
    
'This is to paste
    
     ActiveCell.Offset(1, -9).Range("A1").Select
    ActiveSheet.Paste
    
 'This is to copy the Title and order record information
    
'and paste it
    
    ActiveCell.Offset(-1, -2).Range("A1:B1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
'Turn off the paste thing and then delete the cells that don't need

    ActiveCell.Offset(-1, 11).Range("A1:I1").Select
    Application.CutCopyMode = False
    
    Selection.Delete Shift:=xlToLeft
    
     ' ActiveCell.Offset(-1, 11).Range("A1").Select
    ' Application.CutCopyMode = False
    
'Go back to the original point
    ActiveCell.Offset(0, -11).Select
    
    Loop

    
'LOOP END HERE

'End If

Next i


  
   
End Sub


   








