Attribute VB_Name = "ZZ_JUNK_Z1_Main_Ribbon"
Option Explicit





Global rib As IRibbonUI
Global IndexOfSelectedItem As Integer
Global ItemCount As Integer
Global ListItemsRg As Range

''Excel calls this went it loads our workbook because our RibbonX
'' specified it:  onLoad="ribbonLoaded"
Sub ribbonLoaded(ribbon As IRibbonUI)
    Set rib = ribbon ''We capture the ribbon variable for later use, specifically to invalidate it.  When you invalidate the ribbon Excel recreates it.
End Sub


Sub RedoRib()
    rib.Invalidate
End Sub

''=========Drop Down Code =========

''Callback for Dropdown getItemCount.
''Tells Excel how many items in the drop down.
Sub DDItemCount(control As IRibbonControl, ByRef returnedVal)
    With INTERNALS.ListObjects("have_several_tabs").ListColumns("have_several_tabs")
        Set ListItemsRg = .DataBodyRange
        ItemCount = ListItemsRg.Rows.Count
        returnedVal = ItemCount
    End With
End Sub

''Callback for dropdown getItemLabel.
''Called once for each item in drop down.
''If DDItemCount tells Excel there are 10 items in the drop down
''Excel calls this sub 10 times with an increased "index" argument each time.
''We use "index" to know which item to return to Excel.
Sub DDListItem(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = ListItemsRg.Cells(index + 1).value ''index is 0-based, our list is 1-based so we add 1.
End Sub

''Drop down change handler.
''Called when a drop down item is selected.
Sub DDOnAction(control As IRibbonControl, ID As String, index As Integer)
    ''All we do is note the index number of the item selected.
    ''We use this in sub DDItemSelectedIndex below to reselect the current
    ''item, if possible, after an invalidate.
    IndexOfSelectedItem = index
End Sub

''Returns index of item to display.
''To display current item after the drop down is invalidated.
Sub DDItemSelectedIndex(control As IRibbonControl, ByRef returnedVal)
    If IndexOfSelectedItem > ItemCount - 1 Then IndexOfSelectedItem = ItemCount - 1 ''In case list was shortened
    returnedVal = IndexOfSelectedItem
End Sub

''------- End DD Code --------


