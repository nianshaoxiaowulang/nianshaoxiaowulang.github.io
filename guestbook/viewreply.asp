<% Option Explicit %>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		viewreply.asp
'**
'** �ļ�˵�����鿴�ظ�ҳ��
'** �޸����ڣ�2005-04-07
'**************************************

dim id,replycode
	id=sql_filter(Trim(Request.QueryString("id")))
	replycode=Request.Form("replycode")

if id="" or (not isnumeric(id)) then
	errinfo="<li>�����id��ţ�"
	call error()
end if

pagename="�鿴�ظ�"
call pageinfo()

	sql="Select * from [topic] where id="&id
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3

if rs.eof and rs.bof then
	errinfo="<li>�����Բ����ڡ�"
else
	if rs("replycode")="" or (not rs("reply")=1) then
		errinfo="<li>������û�б��ظ���"
	end if
end if
call error()

if replycode="" then
	Response.Write ""
else
	if replycode=rs("replycode") then
	%>
	<table align="center" cellpadding="7" cellspacing="1" class="table1" width="95%">
		<tr>
			<td width="100%" class="tablebody3">
				<B>�������ԣ�</B><br>
				<FONT face="Verdana" SIZE="1" COLOR="#434259"><%=rs("usertime")%></FONT>
			</td>
		</tr>
		<tr>
			<td width="100%" class="tablebody1">
				���⣺<b><%=HTMLencode(rs("usertitle"))%></b>
				<p><%=UBBcode(rs("usercontent"),0)%></p><p>
			</td>
		</tr>
		<tr>
			<td width="100%" class="tablebody3"	style="color: <%=maincolor%>">
				<B>����Ա�Ļظ���</B><br>
				<FONT SIZE="1" face="Verdana"><%=rs("retime")%></FONT>
			</td>
		</tr>
		<tr>
			<td width="80%" class="tablebody1" valign="top">
				<%=UBBcode(rs("recontent"),1)%><p>
			</td>
		</tr>
		<tr>
			<td valign="middle" align="center" class="tablebody3" height="21">
				<a href='javascript:window.close()'>�رմ���</a>
			</td>
		</tr>
	</table>
	<%
	else
	%>
	<p align=center><img width="122" height="50" border="0" src="images/error.gif"></p>
	<p>
	<table cellpadding=6 cellspacing=1 align=center class=table1 width='550'>
		<tr>
			<td width='100%' class=tablebody3><B><FONT COLOR="red">������Ļظ��鿴�벻��ȷ�����������룺</FONT></B></td>
		</tr>
		<form action="viewreply.asp?id=<%=id%>" method="POST" name="replyform">
		<tr>
			<td width='100%' height='100' align=center class=tablebody1>
				�ظ��鿴�룺<input name="replycode" type="password" size="15" maxlength="100">&nbsp;
				<input type="submit" name="submit" value="�鿴�ظ�">
			</td>
		</tr>
		</form>
		<tr>
			<td width='100%' align=center class=tablebody3><a href='javascript:window.close()'>�رմ���</a></td>
		</tr>
	</table>
	</p>
<%
	end if
end if
%>