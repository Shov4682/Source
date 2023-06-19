'// TODO Creat Add_Jumper_Temp_Table Template for all entries
'// TODO Populate path queries

'// Scenerio 2  
Private Sub Add_One_Jumper_w_Continue_Button_Click()
Dim New_Jumper_ID As Long
Dim New_Path_ID As Long
Dim New_Path_Order As Long
Dim Next_Jumper_ID As Long  '//Next Jumper to be evalutated
Dim Next_End_A As Long
Dim Next_End_B As Long
Dim Next_Path_ID As Long
Dim Next_Path_Order As Long
Dim End_A_Link_Jumper_ID As Long
Dim End_A_Link_Path_ID As Long
Dim End_A_Link_Path_Order As Long
Dim End_B_Link_Jumper_ID As Long
Dim End_B_Link_Path_ID As Long
Dim End_B_Link_Path_Order As Long
Dim Transpose_First_Ascending_Jumper_ID As Long
Dim Transpose_Next_Ascending_Jumper_ID As Long
Dim Transpose_Descending_Jumper_ID As Long


DoCmd.SetWarnings False
DoCmd.OpenQuery "New_Jumper_Append_Query", acViewNormal, acEdit

New_Jumper_ID = DMax("Fiber_Jumper_Table_PKey", "Fiber_Jumpers_Table")
New_Jumper_ID_Combo.Value = New_Jumper_ID + 1
Me.Requery
DoCmd.OpenQuery "New_Jumper_Number_Table_Update_Query", acViewNormal, acEdit
DoCmd.SetWarnings True

End_A_Link_Path_ID = DLookup("Fiber_Path_ID", "End_A_Link_Path_Detail_Query", 1 = 1)
End_A_Link_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "End_A_Link_Path_Detail_Query", 1 = 1)

If Not IsNull(End_A_Link_Jumper_ID) Then
  If End_A_Link_Path_ID = 0 Then
    New_Jumper_ID = DMax("Fiber_Jumper_Table_PKey", "Fiber_Jumpers_Table")
    New_Path_ID = DMax("Fiber_Path_ID", "Fiber_Jumpers_Table")
    
    New_Jumper_ID_Combo.Value = New_Jumper_ID
    New_Path_ID_Combo.Value = New_Path_ID + 1

    End_A_Link_Jumper_ID_Combo.Value = End_A_Link_Jumper_ID
    End_A_Link_Path_ID_Combo.Value = New_Path_ID + 1



    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Update_End_A_Path_ID_If_Blank", acViewNormal, acEdit
    DoCmd.OpenQuery "Update_New_Path_ID_To_End_A_Path_ID", acViewNormal, acEdit
    DoCmd.SetWarnings True

  Else  '// if ("Fiber_Path_ID", "End_A_Link_Path_Detail_Query") > 0 then

    New_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "New_Jumper_Path_Detail_Query", 1 = 1)
    Next_Path_ID = DLookup("Fiber_Path_ID", "End_A_Link_Path_Detail_Query", 1 = 1)
    New_Path_ID = DLookup("Fiber_Path_ID", "New_Jumper_Path_Detail_Query", 1 = 1)

    New_Jumper_ID_Combo.Value = New_Jumper_ID
    New_Path_ID_Combo.Value = New_Path_ID
    Next_Path_ID_Combo.Value = Next_Path_ID

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Update_New_Path_ID_To_Next_Path_ID_Query", acViewNormal, acEdit
    DoCmd.SetWarnings True

  End If

DoCmd.SetWarnings False
DoCmd.OpenQuery "New_Fiber_Path_Append_Query", acViewNormal, acEdit
DoCmd.SetWarnings True

End If

Transpose_First_Ascending_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "Transpose_First_Ascending_Jumper_Query", 1 = 1)

If Not IsNull(Transpose_First_Ascending_Jumper_ID) Then
  Next_Jumper_ID_Combo.Value = Transpose_First_Ascending_Jumper_ID
        
  Next_End_A = DLookup("End_A", "Transpose_First_Ascending_Jumper_Query", 1 = 1)
  Next_End_B = DLookup("End_B", "Transpose_First_Ascending_Jumper_Query", 1 = 1)
    
  Next_End_A_Combo.Value = Next_End_A
  Next_End_B_Combo.Value = Next_End_B
    
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "Transpose_Next_Jumper_Update_Query", acViewNormal, acEdit 
  DoCmd.OpenQuery "Next_Transposed_Jumper_Update_Query", acViewNormal, acEdit
  DoCmd.SetWarnings True
    
  Do
    Next_Jumper_ID = Nz(DLookup("Fiber_Jumper_Table_PKey", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
    Next_End_A = Nz(DLookup("End_A", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
    Next_End_B = Nz(DLookup("End_B", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
    Next_End_A_Switch_Port = Nz(Dlookup("End_A_Switch_Port", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
    Next_End_B_Switch_Port = Nz(Dlookup("End_B_Switch_Port", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
    Next_Path_ID = Nz(Dlookup("Fiber_Path_ID", "Transpose_Next_Ascending_Jumper_Query", 1 = 1))
        
    Next_Jumper_ID_Combo.Value = Next_Jumper_ID
    Next_End_A_Combo.Value = Next_End_A
    Next_End_B_Combo.Value = Next_End_B
        
    If Next_Jumper_ID = 0 Then Exit Do
        
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Transpose_Next_Jumper_Update_Query", acViewNormal, acEdit
    DoCmd.OpenQuery "Next_Transposed_Jumper_Update_Query", acViewNormal, acEdit
    DoCmd.SetWarnings True




    IF End_A_Switch_Port > 0 then
      DoCmd.SetWarnings False
      DoCmd.OpenQuery "End_A_Switch_Port_Update_Query", acViewNormal, acEdit
      DoCmd.SetWarnings True
    End If
  Loop Until (IsNull(Next_Jumper_ID) Or Next_Jumper_ID = 0)
End If

Next_Path_ID = Nz(DLookup("Fiber_Path_ID", "Next_End_A_Link_Path_Detail_query", 1 = 1))
New_Path_ID = DMax("Fiber_Path_ID", "Fiber_Jumpers_Table")
  
Next_Path_ID_Combo.Value = Next_Path_ID
New_Path_ID_Combo.Value = New_Path_ID

DoCmd.SetWarnings False
If (Next_Path_ID) > (New_Path_ID) Then
  DoCmd.OpenQuery "Update_Next_Path_ID_To_New_Path_ID_Query", acViewNormal, acEdit
Else
  DoCmd.OpenQuery "Update_New_Path_ID_To_Next_Path_ID_Query", acViewNormal, acEdit
End If
DoCmd.SetWarnings True

New_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "New_Jumper_Path_Detail_Query", 1 = 1)
New_Path_Order = DLookup("Fiber_Path_Order", "New_Jumper_Path_Detail_Query", 1 = 1)
Next_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "End_A_Link_Path_Detail_Query", 1 = 1)
Next_Path_Order = DLookup("Fiber_Path_Order", "End_A_Link_Path_Detail_Query", 1 = 1)

New_Jumper_ID_Combo.Value = New_Jumper_ID
New_Path_Order_Combo.Value = Next_Path_Order - 1

DoCmd.SetWarnings False
DoCmd.OpenQuery "Update_New_Path_Order_Query", acViewNormal, acEdit
DoCmd.SetWarnings True

Next_Jumper_ID = DLookup("Fiber_Jumper_Table_PKey", "End_B_Link_Path_Detail_Query", 1 = 1)
Next_End_A = DLookup("End_A", "End_B_Link_Path_Detail_Query", 1 = 1)
Next_End_B = DLookup("End_B", "End_B_Link_Path_Detail_Query", 1 = 1)

Next_Jumper_ID_Combo.Value = Next_Jumper_ID
Next_End_A_Combo.Value = Next_End_A
Next_End_B_Combo.Value = Next_End_B

DoCmd.SetWarnings False
DoCmd.OpenQuery "Next_Jumper_Update_Query", acViewNormal, acEdit
DoCmd.SetWarnings True


New_Jumper_ID = Nz(DLookup("Fiber_Jumper_Table_PKey", "New_Jumper_Path_Detail_Query", 1 = 1))
New_Path_Order = Nz(DLookup("Fiber_Path_Order", "New_Jumper_Path_Detail_Query", 1 = 1))
Next_Jumper_ID = Nz(DLookup("Next_Jumper_ID", "Next_Jumper_Table", 1 = 1))
Next_Path_Order = Nz(DLookup("Fiber_Path_Order", "Next_End_B_Link_Path_Detail_Query", 1 = 1))

Next_Jumper_ID_Combo.Value = Next_Jumper_ID
Next_Path_Order_Combo.Value = New_Path_Order - 1

DoCmd.SetWarnings False
DoCmd.OpenQuery "Update_Next_Path_Order_Query", acViewNormal, acEdit
DoCmd.SetWarnings True

Do
  Next_Jumper_ID = Nz(DLookup("Fiber_Jumper_Table_PKey", "Next_End_B_Link_Path_Detail_Query", 1 = 1))
If Next_Jumper_ID = 0 Then Exit Do
       Next_End_A = Nz(DLookup("End_A", "Next_End_B_Link_Path_Detail_Query", 1 = 1))
       Next_End_B = Nz(DLookup("End_B", "Next_End_B_Link_Path_Detail_Query", 1 = 1))

  Next_Jumper_ID_Combo.Value = Next_Jumper_ID
  Next_End_A_Combo.Value = Next_End_A
  Next_End_B_Combo.Value = Next_End_B

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "Next_Jumper_Update_Query", acViewNormal, acEdit
  DoCmd.SetWarnings True


  Next_Jumper_ID = Nz(DLookup("Next_Jumper_ID", "Next_Jumper_Table", 1 = 1))
  New_Path_Order = Nz(DLookup("Fiber_Path_Order", "Next_End_A_Link_Path_Detail_Query", 1 = 1))

  Next_Jumper_ID_Combo.Value = Next_Jumper_ID
  Next_Path_Order_Combo.Value = New_Path_Order - 1

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "Update_Next_Path_Order_Query", acViewNormal, acEdit
  DoCmd.SetWarnings True
Loop Until (IsNull(Next_Jumper_ID) Or Next_Jumper_ID = 0)

'=======================
End_A_Switch_Port = Nz(Dlookup("End_A_Switch_Port", "Add_Jumper_Temp_Table", 1 = 1))
End_B_Switch_Port = Nz(Dlookup("End_B_Switch_Port", "Add_Jumper_Temp_Table", 1 = 1))
Next_Path_ID = Nz(Dlookup("Fiber_Path_ID", "Next_Jumper_Table", 1 = 1))

IF End_A_Switch_Port > 0 then
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "End_A_Switch_Port_Update_Query", acViewNormal, acEdit
    DoCmd.SetWarnings True

End If

IF End_B_Switch_Port > 0 Then 
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "End_B_Switch_Port_Update_Query", acViewNormal, acEdit
    DoCmd.SetWarnings True
End If
test 2





    '---------------------------------------------------------------------------------------------
    "Update_to_End_A_Path_ID_Query" 
    End_A_Link_Jumper_ID = Dlookup("Fiber_Jumper_Table_PKey","End_A_Link_Path_Detail_Query", 1 = 1)
    End_A_Link_Path_ID = Dlookup("Fiber_Path_ID","End_A_Link_Path_Detail_Query", 1 = 1)
    End_A_Link_Path_Order = Dlookup("Fiber_Path_Order","End_A_Link_Path_Detail_Query", 1 = 1)
    End_B_Link_Jumper_ID = Dlookup("Fiber_Jumper_Table_PKey", "End_B_Link_Path_Detail_Query", 1 = 1)
    End_B_Link_Path_ID = Dlookup("Fiber_Path_ID","End_B_Link_Path_Detail_Query", 1 = 1)
    End_B_Link_Path_Order = Dlookup("Fiber_Path_Order","End_B_Link_Path_Detail_Query", 1 = 1)
    '----------------------------------------------------------------------------------------------

    End_A_Link_Jumper_ID_Combo.value = End_A_Link_Jumper_ID
    End_A_Link_Path_ID_Combo.value = End_A_Link_Path_ID
    End_A_Link_Path_Order_Combo.value = End_A_Link_Path_Order
    End_B_Link_Jumper_ID_Combo.value = End_B_Link_Jumper_ID
    End_B_Link_Path_ID_Combo.value = End_B_Link_Path_ID
    End_B_Link_Path_Order_Combo.value = End_B_Link_Path_Order

    End_A_Link_Jumper_ID_Combo.value = ""
    End_A_Link_Path_ID_Combo.value = ""
    End_A_Link_Path_Order_Combo.value = ""
    End_B_Link_Jumper_ID_Combo.value = ""
    End_B_Link_Path_ID_Combo.value = ""
    End_B_Link_Path_Order_Combo.value = ""
    New_Jumper_ID_Combo.value = ""
    New_Path_ID_Combo.value = ""
    New_Path_Order_Combo.value = ""
    Next_Jumper_ID_Combo.value = ""
    Next_Path_ID_Combo.value = ""
    Next_Path_Order_Combo.value = ""




'''
''' Returns the boolean value of !isNull
'''
Protected Func MyLookup(columnName As String, queryName As String, criteria As String = "1=1") As Boolean
    '// TODO: Remember to come back and set this up as a handler to 
    '//       return both the !isNull.val && DLookup.val
    Dim returnObject = {
        lookupVal = DLookup(columnName, queryName, criteria),
        notIsNull = null
    }
    returnObject.notIsNull = Not IsNull(DLookup(returnObject.lookupVal), 1 = 1)

    return returnObject.notIsNull
End Func
'Dim asdf As Integer = DLookup("Fiber_Jumper_Table_PKey", "End_A_Link_Path_Detail_Query", 1 = 1)
' if MyLookup("Fiber_Jumper_Table_PKey", "End_A_Link_Path_Detail_Query") Then



If (true) Then
    '// do shit
    If (!false) Then
        Dim yerMomma As Boolean = False
        Do {
            '// Do endless loop forever
            If (yerMomma = True)
                yerMomma = False
            End If
        } while (yerMomma != False)
    End If
Else
    '// do less shit
End If


{
    [
        {

        },
        {
            {
                function () => ({
                    do {
                        if (true !== false && false == 0) {

                        }
                    } while (true);
                })
            }
        }
    ]
}
