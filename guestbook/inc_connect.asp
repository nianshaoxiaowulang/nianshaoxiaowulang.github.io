<%
'**************************************
'**		inc_connect.asp
'**
'** �ļ�˵�������ݿ������ļ�
'** �޸����ڣ�2005-04-07
'**************************************

'�뽫����ڶ����еġ�#database.asp���滻Ϊ���Լ�������������ݿ��ַ
On Error Resume Next

dim db,conn,connstr,sql,rs
	db="dataps/#databaseps.asp"
	set conn = server.createobject("ADODB.connection")
	connstr="provider=microsoft.JET.OLEDB.4.0;data source=" & server.mappath(db)
	conn.open connstr

If Err Then
	Err.Clear
	Set conn = Nothing
	Response.Write "�������ݿ��ʱ�����"
	Response.End
End If

function closedatabase
	conn.close
	set conn = nothing
end function
%>