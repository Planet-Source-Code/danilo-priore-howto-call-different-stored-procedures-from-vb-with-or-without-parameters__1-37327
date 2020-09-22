<div align="center">

## Howto call different Stored Procedures from VB with or without parameters


</div>

### Description

This a example for call one or more store procedures with different input parameters or without parameters.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Danilo Priore](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/danilo-priore.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/danilo-priore-howto-call-different-stored-procedures-from-vb-with-or-without-parameters__1-37327/archive/master.zip)





### Source Code

<font size="-2">
<b>Code:</b><br>
<br>
Public Function ExecuteSP(sProcName As String,
ParamArray aParams()) As ADODB.Recordset<br>
 Dim cmd As ADODB.Command<br>
 Set cmd = New ADODB.Command<br>
 Set cmd.ActiveConnection = conn<br>
 cmd.CommandText = sProcName<br>
 cmd.CommandType = adCmdStoredProc<br>
 If aParams(0) Is Nothing Then<br>
 <font color="green">
' if NOT use parameters</font><br>
 Set ExecuteSP = cmd.Execute<br>
 Else<br>
 <font color="green">
 ' if use parameters</font><br>
 Set ExecuteSP = cmd.Execute(, aParams)<br>
 End If<br>
 Set cmd = Nothing<br>
End Function<br>
<br>
<b>Example to call:</b><br>
<br>
Dim rs = ADODB.Recordset<br>
<font color="green">
' without params</font><br>
Set rs = ExecuteSP("sp_selectall",Nothing)<br>
<font color="green">
' with params</font><br>
Set rs = ExecuteSP("sp_find","Danilo","Priore")<br>
<font color="green">
' with params without return records</font><br>
Call ExecuteSP("sp_delete",1234)<br>
<font color="green">
' when sp_selectall = "select * from users"<br>
' and sp_find = "select * from user where name=@name and surname=@surname"<br>
' and sp_delete = "delete form user where id=@id"<br>
</font>
</font>

