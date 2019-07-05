<%
'**************************************
'**		inc_connect.asp
'**
'** 文件说明：数据库连接文件
'** 修改日期：2005-04-07
'**************************************

'请将下面第二行中的“#database.asp”替换为你自己重命名后的数据库地址
On Error Resume Next

dim db,conn,connstr,sql,rs
	db="dataps/#databaseps.asp"
	set conn = server.createobject("ADODB.connection")
	connstr="provider=microsoft.JET.OLEDB.4.0;data source=" & server.mappath(db)
	conn.open connstr

If Err Then
	Err.Clear
	Set conn = Nothing
	Response.Write "连接数据库的时候出错。"
	Response.End
End If

function closedatabase
	conn.close
	set conn = nothing
end function
%>