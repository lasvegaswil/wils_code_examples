Attribute VB_Name = "Discount_Import"
Option Compare Database


Public Sub Import_Discounts()


'qry_Import_Discounts_From_CSV
    DoCmd.RunSQL "INSERT INTO tbl_Discounts_and_Rebates ( Doc_ID, Date_Created, Status, Doc_Type, Created_At, Second_Date, Player_Name, Created_By, Player_ID, Amount, Balance ) SELECT Discounts_N_Rebates.DocumentID, Discounts_N_Rebates.Created, Discounts_N_Rebates.Status, Discounts_N_Rebates.Doc_Type, Discounts_N_Rebates.CreatedAt, Discounts_N_Rebates.Date_Complete, Discounts_N_Rebates.Player, Discounts_N_Rebates.CreatedBy, Discounts_N_Rebates.PlayerID, Discounts_N_Rebates.Amount, Discounts_N_Rebates.Balance FROM Discounts_N_Rebates WHERE (((Discounts_N_Rebates.DocumentID) Is Not Null) AND ((Discounts_N_Rebates.Doc_Type)='Discount' Or (Discounts_N_Rebates.Doc_Type)='Rebate') AND ((Discounts_N_Rebates.Player) Is Not Null));"

    'qry_Update_Discount_Allowances
    DoCmd.RunSQL "INSERT INTO tbl_Allowances ( Patron_Name, Dragon_Club, Allowance_Date, Type, Amount ) SELECT Player_Info_Table_Main.Name, tbl_Discounts_and_Rebates.Player_ID, Format([tbl_discounts_and_rebates]![Date_Created], 'mm/dd/yyyy') AS Date_Chooser, 'Discount' AS Disc, tbl_Discounts_and_Rebates.Amount FROM Player_Info_Table_Main RIGHT JOIN tbl_Discounts_and_Rebates ON Player_Info_Table_Main.Dragon_Club = tbl_Discounts_and_Rebates.Player_ID;"





End Sub

