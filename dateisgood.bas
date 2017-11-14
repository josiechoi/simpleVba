Sub DateIsGood()

Dim i As Long
Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(2, 1).Select
'starting point is cell A2

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


