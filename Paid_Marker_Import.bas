Attribute VB_Name = "Paid_Marker_Import"
Option Compare Database

Public Sub Move_Paid_Markers()




'Clear Temp Tables


DoCmd.RunSQL "DELETE tbl_Temp_Paid_Markers.*" & _
                            "FROM tbl_Temp_Paid_Markers;"
                            

DoCmd.RunSQL " DELETE tbl_Temp_Paid_Markers_dt.*" & _
                            "FROM tbl_Temp_Paid_Markers_dt;"



'qry_Paid_Marker_Temp_Import_Append

DoCmd.RunSQL "INSERT INTO tbl_Temp_Paid_Markers ( DocumentID, Created, Status, DocDescription, CreatedAt, Date_Created, Player, CreatedBy, Location, PlayerID, Amount, Balance )" & _
                            " SELECT Documents_Markers.DocumentID, Documents_Markers.Created, Documents_Markers.Status, Documents_Markers.DocDescription, Documents_Markers.CreatedAt, Documents_Markers.Date_Created, Documents_Markers.Player, Documents_Markers.CreatedBy, Documents_Markers.Location, Documents_Markers.PlayerID, Documents_Markers.Amount, Documents_Markers.Balance" & _
                            " FROM Documents_Markers" & _
                            " GROUP BY Documents_Markers.DocumentID, Documents_Markers.Created, Documents_Markers.Status, Documents_Markers.DocDescription, Documents_Markers.CreatedAt, Documents_Markers.Date_Created, Documents_Markers.Player, Documents_Markers.CreatedBy, Documents_Markers.Location, Documents_Markers.PlayerID, Documents_Markers.Amount, Documents_Markers.Balance" & _
                            " HAVING (((Documents_Markers.DocumentID) Is Not Null));"


'qry_Paid_Marker_Temp_Import_dt_Append

DoCmd.RunSQL "INSERT INTO tbl_Temp_Paid_Markers_dt ( DocumentID, dt_DocTransID, dt_Date, dt_Status, dt_move, dt_Location, dt_Amount, dt_User, dt_Comment, dt_Paymethod )" & _
                            " SELECT Documents_Markers.DocumentID, Documents_Markers.dt_DocTransID, Documents_Markers.dt_Date, Documents_Markers.dt_Status, Documents_Markers.dt_move, Documents_Markers.dt_Location, Documents_Markers.dt_Amount, Documents_Markers.dt_User, Documents_Markers.dt_Comment, Documents_Markers.dt_PayMethod" & _
                            " FROM Documents_Markers" & _
                            " GROUP BY Documents_Markers.DocumentID, Documents_Markers.dt_DocTransID, Documents_Markers.dt_Date, Documents_Markers.dt_Status, Documents_Markers.dt_move, Documents_Markers.dt_Location, Documents_Markers.dt_Amount, Documents_Markers.dt_User, Documents_Markers.dt_Comment, Documents_Markers.dt_PayMethod" & _
                            " HAVING (((Documents_Markers.dt_DocTransID) Is Not Null));"


'qry_Paid_Marker_Append

DoCmd.RunSQL "INSERT INTO tbl_Paid_Markers ( DocumentID, Created, Status, DocDescription, CreatedAt, Date_Created, Player, CreatedBy, Location, PlayerID, Amount, Balance )" & _
                            " SELECT tbl_Temp_Paid_Markers.DocumentID, tbl_Temp_Paid_Markers.Created, tbl_Temp_Paid_Markers.Status, tbl_Temp_Paid_Markers.DocDescription, tbl_Temp_Paid_Markers.CreatedAt, tbl_Temp_Paid_Markers.Date_Created, tbl_Temp_Paid_Markers.Player, tbl_Temp_Paid_Markers.CreatedBy, tbl_Temp_Paid_Markers.Location, tbl_Temp_Paid_Markers.PlayerID, tbl_Temp_Paid_Markers.Amount, tbl_Temp_Paid_Markers.Balance " & _
                            " FROM tbl_Temp_Paid_Markers;"


'qry_Paid_Marker_dt_Append

DoCmd.RunSQL "INSERT INTO tbl_Paid_Markers_dt ( DocumentID, dt_DocTransID, dt_Date, dt_Status, dt_move, dt_Location, dt_Amount, dt_User, dt_Paymethod, dt_Comment )" & _
                            " SELECT tbl_Temp_Paid_Markers_dt.DocumentID, tbl_Temp_Paid_Markers_dt.dt_DocTransID, tbl_Temp_Paid_Markers_dt.dt_Date, tbl_Temp_Paid_Markers_dt.dt_Status, tbl_Temp_Paid_Markers_dt.dt_move, tbl_Temp_Paid_Markers_dt.dt_Location, tbl_Temp_Paid_Markers_dt.dt_Amount, tbl_Temp_Paid_Markers_dt.dt_User, tbl_Temp_Paid_Markers_dt.dt_Paymethod, tbl_Temp_Paid_Markers_dt.dt_Comment" & _
                           "  FROM tbl_Temp_Paid_Markers_dt;"


'Clear Temp Tables


DoCmd.RunSQL "DELETE tbl_Temp_Paid_Markers.*" & _
                            "FROM tbl_Temp_Paid_Markers;"
                            

DoCmd.RunSQL " DELETE tbl_Temp_Paid_Markers_dt.*" & _
                            "FROM tbl_Temp_Paid_Markers_dt;"





End Sub
