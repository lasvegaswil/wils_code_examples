Attribute VB_Name = "Airfare_Import"
Option Compare Database

Public Sub Import_AF()





'Airfare Section


    
     
     On Error Resume Next
     
     
     
    
    'qry_Airfare_Import_01_Clear_Table
    
        DoCmd.RunSQL "DELETE tbl_Rebates_Airfare.* " & _
                                    "FROM tbl_Rebates_Airfare;"

    
    'Import Cage's sheet
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_Airfare_Import", "\\supfs01.ldlv.local\deptdata\Dept_Cage\Logs\Discounts, Airfare, Rebates.xlsx", True
         
              
         
         
    ' qry_Airfare_Import_02_DataImport
    
        DoCmd.RunSQL "INSERT INTO tbl_Rebates_Airfare ( [Agent/ Host], Type, [FORM OF PAYMENT], [NAME OF PATRON], [PATRON ACCOUNT NUMBER], AMOUNT, [Date], [TRIP DATES], [TRIP START], [TRIP END], INITIALS, COMMENT, UNIQUE_ID )" & _
                                   "SELECT tbl_Airfare_Import.[Agent/ Host], tbl_Airfare_Import.Type, tbl_Airfare_Import.[FORM OF PAYMENT (Apply To Markers /Cash/Chips/ MailChekc/FM/SK)] AS Form, tbl_Airfare_Import.[NAME OF PATRON], tbl_Airfare_Import.[PATRON ACCOUNT NUMBER], tbl_Airfare_Import.AMOUNT, tbl_Airfare_Import.Date, tbl_Airfare_Import.[TRIP DATES], tbl_Airfare_Import.[TRIP START], tbl_Airfare_Import.[TRIP END], tbl_Airfare_Import.INITIALS, tbl_Airfare_Import.COMMENT, [type] &  '-' & [patron account number] & '-' & [trip start] AS Unique_ID " & _
                                    "FROM tbl_Airfare_Import " & _
                                  "WHERE (((tbl_Airfare_Import.Type)<>'Discount') AND ((tbl_Airfare_Import.[FORM OF PAYMENT (Apply To Markers /Cash/Chips/ MailChekc/FM/SK)]) Not Like 'app*') AND ((tbl_Airfare_Import.[PATRON ACCOUNT NUMBER]) Is Not Null));"
        
        
    'qry_Airfare_Import_03_Update_Names
    
        DoCmd.RunSQL "UPDATE tbl_Rebates_Airfare LEFT JOIN Player_Info_Table_Main ON tbl_Rebates_Airfare.[PATRON ACCOUNT NUMBER] = Player_Info_Table_Main.Dragon_Club SET Player_Info_Table_Main.Name = [tbl_Rebates_Airfare]![NAME OF PATRON], tbl_Rebates_Airfare.[TRIP START] = Format([tbl_rebates_airfare]![TRIP START],'mm/dd/yyyy');"
        


    'qry_Airfare_Import_04_To_Allowances
    
        DoCmd.RunSQL "INSERT INTO tbl_Allowances ( Patron_Name, Dragon_Club, Type, AMOUNT, Allowance_Date ) " & _
                                    "SELECT tbl_Rebates_Airfare.[NAME OF PATRON], tbl_Rebates_Airfare.[PATRON ACCOUNT NUMBER], tbl_Rebates_Airfare.Type, tbl_Rebates_Airfare.AMOUNT, tbl_Rebates_Airfare.[TRIP START]" & _
                                    "FROM tbl_Rebates_Airfare;"




    'Get rid of the temp table Justin Case
    
      DoCmd.DeleteObject acTable, "tbl_airfare_import"
                

    

End Sub




