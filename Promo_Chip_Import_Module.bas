Attribute VB_Name = "Promo_Chip_Import_Module"
Option Compare Database

Public Sub Promo_Chip_Import()



'Bring in data from Cage Excel Sheet

DoCmd.RunSQL "INSERT INTO tbl_Promo_Chips ( Date_Issued, Guest_name, Club_Number, Amount, Type, Reason, Issued_From, Requester_1, Requester_2, Issued_by, Comment ) " & _
                            " SELECT Cage_Promo_Chip_Log.Date_Issued, Cage_Promo_Chip_Log.Guest_name, Cage_Promo_Chip_Log.Club_Number, Cage_Promo_Chip_Log.Amount, Cage_Promo_Chip_Log.Type, Cage_Promo_Chip_Log.Reason, Cage_Promo_Chip_Log.Issued_From, Cage_Promo_Chip_Log.Requester_1, Cage_Promo_Chip_Log.Requester_2, Cage_Promo_Chip_Log.Issued_by, Cage_Promo_Chip_Log.Comment " & _
                            " FROM Cage_Promo_Chip_Log " & _
                            " WHERE (((Cage_Promo_Chip_Log.Date_Issued) Is Not Null));"


DoCmd.RunSQL "INSERT INTO tbl_Allowances ( Patron_Name, Dragon_Club, Type, Amount, Allowance_Date )" & _
                            " SELECT Player_Info_Table_Main.Name, tbl_Promo_Chips.Club_Number, StrConv([tbl_promo_chips]![type],3) AS Types, tbl_Promo_Chips.Amount, Format([tbl_Promo_Chips]![Date_Issued],'mm/dd/yyyy') AS Date_Chooser" & _
                            " FROM tbl_Promo_Chips LEFT JOIN Player_Info_Table_Main ON tbl_Promo_Chips.Club_Number = Player_Info_Table_Main.Dragon_Club;"



End Sub
