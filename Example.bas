Sub test()

    Dim j As js
    Set j = New js
    
    'so that not overwritten
    j.appendMode = True
    
    ' load underscore.js
    Call j.loadLib("js/underscore.js")
    
    ' load custom js code
    Call j.loadLib("js/main.js")
    
    ' no callback
    j.code = ""
    
    ' invoke the myzipp function in main.js
    arr = j.js.run("myzipp", Array(1, 2, 3), Array("a", "b", "c"))
    
    ' write the result in the worksheet
    Selection = arr
    
    'loop through the result
    For Each i In arr
        Debug.Print i
    Next i

End Sub
