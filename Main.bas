Attribute VB_Name = "Main"


Public Sub main()
    Dim ds As C_Datasources
    Dim d As C_Datasource
    
    Set ds = New C_Datasources
    
    Set d = New C_Datasource
    d.name = "first"
    Call ds.Add(d)
    
    Set d = New C_Datasource
    d.name = "second"
    Call ds.Add(d)
    
    Set d = New C_Datasource
    d.name = "third"
    Call ds.Add(d)
    
    
    Dim iterator As C_Iterator
    Set iterator = ds.I_ObjectList_GetIterator
    
    Do While iterator.HasNext()
        Set d = iterator.GetNext()
        Debug.Print d.name
    Loop
    
    Dim fs As C_Filters
    Dim f As C_Filter
    
    Set fs = New C_Filters
    Set f = New C_Filter
    f.name = "first"
    Call fs.Add(f)
    
    Set f = New C_Filter
    f.name = "second"
    Call fs.Add(f)
    
    Set f = New C_Filter
    f.name = "third"
    Call fs.Add(f)
    
    Set iterator = fs.I_ObjectList_GetIterator
    
    Do While iterator.HasNext()
        Set f = iterator.GetNext()
        Debug.Print f.name
        
        Dim iterator2 As C_Iterator
        Set iterator2 = fs.I_ObjectList_GetIterator
        
        Do While iterator2.HasNext()
            Dim f2 As C_Filter
            Set f2 = iterator2.GetNext()
            Debug.Print ("NESTED LOOP :: " & f2.name)
        Loop
    Loop
    
    
    Set d = ds.GetElementByName("first")
    Debug.Print "–¼‘O‚ÅŽæ“¾::" & d.name
    
    Debug.Print "END"
End Sub
