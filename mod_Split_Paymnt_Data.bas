Attribute VB_Name = "mod_Split_Paymnt_Data"
Option Compare Database

Public Function Break_Out_Text(str$, i%) As String
  On Error Resume Next
  Break_Out_Text = Split(str, ",")(i)
End Function
