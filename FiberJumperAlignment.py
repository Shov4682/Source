AJTT = Add Jumper Temp Table
FJT = Fiber Jumper Table
NPNQ = Newest Path Number Query

if End_A of AJTT = Far_End of End_B in FJT Then
    DoCmd.OpenQuery "New_Jumper_Append_Query", acViewNormal, acEdit
    Update New Jumper Fiber_Path_ID to (Matched Jumper[Fiber_Path_ID])
    Update New Jumper Path_Order_ID to (Matched jumper [Path_Order_ID] + 1 )
    Else
    End if
    If Far_End of End_B = End_A of AJTT Then
    Update Matched Path_Order_ID