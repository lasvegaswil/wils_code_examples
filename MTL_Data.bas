Attribute VB_Name = "MTL_Data"
Option Compare Database

Public Sub MTL_Data_Import()

On Error Resume Next

DoCmd.SetWarnings False


'qry_Clearn_MTL_Temp

    DoCmd.RunSQL "DELETE tbl_Temp_MTL_Data.*" & _
                                "FROM tbl_Temp_MTL_Data;"


'qry_MTL_Update_Table

    DoCmd.RunSQL "INSERT INTO tbl_Temp_MTL_Data ( GamingDate, PatronNumber, LastName, FirstName, MiddleInitial, SSN, Sex, Amount, AggregateIn, AggregateOut, Flow, TransactionChoice, CheckNumber, CheckChoice, Area, Shift, LegalName, Height, Weight, Age, Hair, Race, Address, City, State, Zip, BankInfo, DetailDate, Comments, Detail_Date_Convert )" & _
                                " SELECT MTL_Log_Data.GamingDate, MTL_Log_Data.PatronNumber, MTL_Log_Data.LastName, MTL_Log_Data.FirstName, MTL_Log_Data.MiddleInitial, MTL_Log_Data.SSN, MTL_Log_Data.Sex, MTL_Log_Data.Amount, MTL_Log_Data.AggregateIn, MTL_Log_Data.AggregateOut, MTL_Log_Data.Flow, MTL_Log_Data.TransactionChoice, MTL_Log_Data.CheckNumber, MTL_Log_Data.CheckChoice, MTL_Log_Data.Area, MTL_Log_Data.Shift, MTL_Log_Data.LegalName, MTL_Log_Data.Height, MTL_Log_Data.Weight, MTL_Log_Data.Age, MTL_Log_Data.Hair, MTL_Log_Data.Race, MTL_Log_Data.Address, MTL_Log_Data.City, MTL_Log_Data.State, MTL_Log_Data.Zip, MTL_Log_Data.BankInfo, MTL_Log_Data.DetailDate, MTL_Log_Data.Comments, Format((CDate([DetailDate])),'mm/dd/yyyy hh:nn:ss') AS Detail_Date_Convert" & _
                                " FROM MTL_Log_Data;"


'qry_MTL_Update_Table_02

    DoCmd.RunSQL "INSERT INTO tlb_MTL_Data" & _
                                " SELECT tbl_Temp_MTL_Data.*" & _
                                "FROM tbl_Temp_MTL_Data;"



'qry_Clearn_MTL_Temp

    DoCmd.RunSQL "DELETE tbl_Temp_MTL_Data.*" & _
                                "FROM tbl_Temp_MTL_Data;"



















'DoCmd.SetWarnings True

End Sub
