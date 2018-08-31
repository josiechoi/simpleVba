Sub AresPrintReserve()
'
'This macro is created at the request of Michelle Chen to facilitate print reserve.
'
Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'This is the safety. This macro would only work if Cell A1 is "Item ID"


If Range("A1").Value = "Item ID" Then

'Deleting fields that are not needed

    
Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:AM").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:Q").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
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
