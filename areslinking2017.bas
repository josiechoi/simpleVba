Sub AresLinking2017()
'
' Macro6 Macro
'This macro is created at the request of Ann Ludbrook to facilitate e-reserves staff for linking
'
Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'This is the safety. This macro would only work if Cell A1 is "Item ID"


If Range("A1").Value = "Item ID" Then

'Deleting fields that are not needed


    Columns("B:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:AE").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:M").Select
    Selection.Delete Shift:=xlToLeft
    
  ' Filtering out to "Weblink" only
  
    ActiveSheet.Range("$A$1:$H$" & LastRow).AutoFilter Field:=5, Criteria1:="WebLink"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-21
    
    
    'Paste it to a new column (K)
    
 
    Range("K1").Select
    ActiveSheet.Paste
    
    
    'Deleting the old table
    
    
    Columns("A:J").Select
    Range("J1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    
    'Sorting the result by Item Id
    
    
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:= _
        Range("A1:A" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Else

  MsgBox "Are you sure this file is from Ares? Please check again. "
  Exit Sub
  
End If

    
End Sub
