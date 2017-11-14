Sub AresPrintReserve()
'
'This macro is created at the request of MC to facilitate print reserve. Created Nov 14,2017
'
Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'This is the safety. This macro would only work if Cell A1 is "Item ID"


If Range("A1").Value = "Item ID" Then

'Deleting fields that are not needed

    
    
    Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:AR").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:H").Select
    Selection.Delete Shift:=xlToLeft
    
    
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
Else

  MsgBox "Are you sure this file is from Ares? Please check again. "
  Exit Sub
  
End If

    
End Sub
