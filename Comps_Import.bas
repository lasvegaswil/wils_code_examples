Attribute VB_Name = "Comps_Import"
Option Compare Database

Public Sub Import_Comps()

'Clear up tlb_Temp_Table_Comps_01

    DoCmd.RunSQL "DELETE tbl_Temp_Table_Comps.* " & _
                                "FROM tbl_Temp_Table_Comps;"



'qry_Comps_By_Player_Upload

    DoCmd.RunSQL "INSERT INTO tbl_Comps_By_Player ( Player, PlayerID, Comp_Number, Location, Comp_Bucket_Type, Type, Dept, Authorizer, Origin, Status, Issue_Date, Covers, Amount, Internal_Comment )" & _
                                " SELECT Comps_By_Player_IGT_Report.Player, Comps_By_Player_IGT_Report.PlayerID, Comps_By_Player_IGT_Report.Comp_Number, Comps_By_Player_IGT_Report.Location, Comps_By_Player_IGT_Report.Comp_Bucket_Type, Comps_By_Player_IGT_Report.Type, Comps_By_Player_IGT_Report.Dept, Comps_By_Player_IGT_Report.Authorizer, Comps_By_Player_IGT_Report.Origin, Comps_By_Player_IGT_Report.Status, Comps_By_Player_IGT_Report.Issue_Date, Comps_By_Player_IGT_Report.Covers, Comps_By_Player_IGT_Report.Amount, Comps_By_Player_IGT_Report.Internal_Comment" & _
                                "  FROM Comps_By_Player_IGT_Report" & _
                                "   WHERE (((Comps_By_Player_IGT_Report.Player) Not Like 'textbox70' And (Comps_By_Player_IGT_Report.Player) Not Like 'TEST,*' And (Comps_By_Player_IGT_Report.Player)<>'DUMPTY, HUMPTY' And (Comps_By_Player_IGT_Report.Player)<>'Dealer, Dealer' And (Comps_By_Player_IGT_Report.Player)<>'JACOBY, DAVE' And (Comps_By_Player_IGT_Report.Player)<>'SEAGER, JORDAN') AND ((Comps_By_Player_IGT_Report.PlayerID)>79999999));"
    
    
    
    
    'qry_comps_upload_for_updates
    
    DoCmd.RunSQL "INSERT INTO tbl_Temp_Table_Comps ( Player, PlayerID, Comp_Number, Location, Comp_Bucket_Type, Type, Dept, Authorizer, Origin, Status, Issue_Date, Covers, Amount, Internal_Comment )" & _
                                " SELECT Comps_By_Player_IGT_Report.Player, Comps_By_Player_IGT_Report.PlayerID, Comps_By_Player_IGT_Report.Comp_Number, Comps_By_Player_IGT_Report.Location, Comps_By_Player_IGT_Report.Comp_Bucket_Type, Comps_By_Player_IGT_Report.Type, Comps_By_Player_IGT_Report.Dept, Comps_By_Player_IGT_Report.Authorizer, Comps_By_Player_IGT_Report.Origin, Comps_By_Player_IGT_Report.Status, Comps_By_Player_IGT_Report.Issue_Date, Comps_By_Player_IGT_Report.Covers, Comps_By_Player_IGT_Report.Amount, Comps_By_Player_IGT_Report.Internal_Comment" & _
                                " FROM Comps_By_Player_IGT_Report" & _
                               "  WHERE (((Comps_By_Player_IGT_Report.Player) Not Like 'textbox70' And (Comps_By_Player_IGT_Report.Player) Not Like 'TEST,*' And (Comps_By_Player_IGT_Report.Player)<>'DUMPTY, HUMPTY' And (Comps_By_Player_IGT_Report.Player)<>'Dealer, Dealer' And (Comps_By_Player_IGT_Report.Player)<>'JACOBY, DAVE' And (Comps_By_Player_IGT_Report.Player)<>'SEAGER, JORDAN') AND ((Comps_By_Player_IGT_Report.PlayerID)>79999999));"
                               

    
    
'qry_Comps_By_Player_Update

   DoCmd.RunSQL "UPDATE tbl_Temp_Table_Comps LEFT JOIN tbl_Comps_By_Player ON tbl_Temp_Table_Comps.Comp_Number = tbl_Comps_By_Player.Comp_Number SET tbl_Comps_By_Player.Status = [tbl_Temp_Table_Comps]![Status]" & _
                            " WHERE ((([tbl_Comps_By_Player]![Status])<>[tbl_Temp_Table_Comps]![Status]));"


'Clear up tlb_Temp_Table_Comps_02

    DoCmd.RunSQL "DELETE tbl_Temp_Table_Comps.* " & _
                                "FROM tbl_Temp_Table_Comps;"




End Sub
