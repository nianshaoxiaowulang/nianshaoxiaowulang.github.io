<% option explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		edit.asp
'**
'** �ļ�˵���������޸�ҳ��
'** �޸����ڣ�2005-04-07
'**************************************

if not session("login")="true" then
	errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
	call error()
end if

select case request.querystring("act")
	case "addnew"
	call addnew()
	case else
	call main()
end select

sub main()

dim id
	id=request.querystring("id")
	pagename="�༭����"
	call pageinfo()
	
	sql="select usertitle,usercontent,checked from [topic] where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3

	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		errinfo="<li>�����Բ����ڡ�"
		call error()
	end if

'---------------������ʾҳ������--------
%>
<br>
<script src="UBB.js"></script>
<form action="?act=addnew" method="post" name="lw_form">
	<table align="center" cellpadding="3" cellspacing="1" class="table1" width="95%">
		<tr>
			<td width="12%" class="tablebody3" align="right">
			<font face="����" color="red">***</font> <b>����</b></td>
			<td width="88%" class="tablebody2">
			<input name="usertitle" size="40" maxlength="100" value="<%=rs("usertitle")%>"></td>
		</tr>
		<tr>
			<td valign="top" class="tablebody3" align="right">
			<font face="����" color="red">***</font> <b>����</b></td>
			<td class="tablebody2">
			<!--#include file="inc_UBB.asp"-->
			<textarea class="smallarea" cols="60" name="usercontent" title="ctrl+enter�ύ" rows="12" onkeydown="ctlent()"><%=back_filter(rs("usercontent"))%></textarea><br>
			<br>
			���������ſ��Խ���������ģ��������ݲ��ܴ���300�ַ�����<br>
			<%
dim ii,i
for i=1 to 42
	if len(i)=1 then ii="0" & i else ii=i
	response.write "<img src=""images/faces/"&ii&".gif"" width=20 height=20 border=0 onclick=""insertsmilie('[face"&ii&"]')"" style=""cursor: hand"">&nbsp;"
	if i=17 or i=34 then response.write "<br>"
next
%> </td>
		</tr>
		<tr>
			<td valign="middle" colspan="2" align="center" class="tablebody1">
			<input type="hidden" name="id" value="<%=id%>">
			<input type="submit" value="�� ��" name="submit">&nbsp;&nbsp;
			<input type="reset" name="submit2" value="�� ��"></td>
		</tr>
<%
if not rs(2)=1 then
%>
		<tr>
			<td colspan="2" align="center" class="tablebody1"><b>
			��ʾ��</b>������Ա��༭�������Զ�ͨ����ˡ�</td>
		</tr>
<%
end if
%>
	</table>
	</div>
</form>
<%
'--------------ҳ��������ʾ����--------
end sub

sub addnew()

	dim id,username,userURL,usermail,userqq,usertitle,usercontent
	id=request.form("id")
	usertitle=trim(request.form("usertitle"))
	usercontent=request.form("usercontent")

	if len(usertitle)>50 then
		errinfo=errinfo & "<li>�����ı���"
	end if

	if usercontent="" then
		errinfo=errinfo & "<li>δ��д��������"
	end if

call error()

	set rs= server.createobject("adodb.recordset")
	sql="select * from [topic] where id="&id
	rs.open sql,conn,3,2
	rs.update

	rs("usertitle")=usertitle
	rs("usercontent")=usercontent
	rs("checked")=1
	rs.update
	rs.close
%>
<script>self.opener.location.reload();</script>
<script>self.close();</script>
<%
end sub
%>