<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		reply.asp
'**
'** �ļ�˵�����ظ�����ҳ��
'** �޸����ڣ�2005-04-07
'**************************************

if not session("login")="true" then
	errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
	call error()
end if

pagename="�ظ�����"
call pageinfo()

select case Request.QueryString("act")
	case "update"
	call update()
	case else
	call main()
end select

sub main()
dim id
	id=Request.QueryString("id")

	sql="Select id,username,usertime,usertitle,usercontent,recontent,top,checked from [topic] where id="&id
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3

	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		errinfo="<li>�����Բ����ڡ�"
		call error()
	end if
%>
<script src="UBB.js"></script>
<form action="?act=update" method="POST" name="lw_form">
	<input type="hidden" name="id" size="12" maxlength="15" value="<%=id%>">
	<div align="center">
		<table align="center" cellpadding="5" cellspacing="1" class="table1" width="95%">
			<tr>
				<td width="100%" class="tablebody3">���⣺<b><%=HTMLencode(rs(3))%></b><br>
				���ߣ�<%=HTMLencode(rs(1))%><br>
				ʱ�䣺<font face="Verdana" SIZE="1" COLOR="#434259"><%=rs(2)%></font>
		<hr color="#555555" align="left" width="40%" size="1">
				<%=UBBcode(rs("usercontent"),rs("top"))%><p></td>
			</tr>
			<tr>
				<td width="100%" align="center" class="tablebody1"><b>����Ա�ظ�</b></td>
			</tr>
			<tr>
				<td width="80%" align="center" class="tablebody2">
				<!--#include file="inc_UBB.asp"-->
				<textarea class="smallarea" cols="60" name="usercontent" title="Ctrl+Enter�ύ" rows="12" onkeydown="ctlent()"><%=rs(5)%></textarea><br>
				<%
	dim ii,i
	for i=1 to 42
		if len(i)=1 then ii="0" & i else ii=i
		response.write "<img src=""images/faces/"&ii&".gif"" width=20 height=20 border=0 onclick=""insertsmilie('[face"&ii&"]')"" style=""CURSOR: hand"">&nbsp;"
		if i=17 or i=34 then response.write "<br>"
	next
%> </td>
			</tr>
			<tr>
				<td valign="middle" align="center" class="tablebody1">
				<input type="Submit" value="�� ��" name="Submit">&nbsp;&nbsp;
				<input type="reset" name="Submit2" value="�� ��"> </td>
			</tr>
			<%
if not rs(6)=1 then
%>
			<tr>
				<td valign="middle" align="center" class="tablebody1"><b>
				��ʾ��</b>������Ա��ظ��������Զ�ͨ����ˡ�</td>
			</tr>
			<%
end if
%>
		</table>
	</div>
</form>
<%
	conn.close
	set rs=nothing
end sub

sub update()

	dim recontent,id
	recontent=request.Form("usercontent")
	id=request.Form("id")

	if recontent="" then
		errinfo="<li>δ��д�ظ����ݡ�"
		call error()
	end if

	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from [topic] where id="&id
	rs.open sql,conn,3,2
	rs.update
	rs("reply")="1"
	rs("retime")=NOW()
	rs("recontent")=recontent
	rs("checked")=1
	rs.update
	rs.close

%>
<script>self.opener.location.reload();</script>
<script>self.close();</script>
<%
end sub
%>