<%
'**************************************
'**		inc_top.asp
'**
'** �ļ�˵�������Ա�������Ϣ
'** �޸����ڣ�2005-04-07
'**************************************
%>
<table width="770" border="0" align="center" cellpadding="1" cellspacing="0" style="border-collapse: collapse">
  <tr>
		<td width="150" style="border-bottom: 1px solid #969696">
		<a onfocus=this.blur() href="index.asp"><img border="0" src="images/logo.gif" width="150" height="45"></a></td>
		<td style="border-bottom: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#FFFFFF', endcolorstr='<%=maincolor%>', gradienttype='1'"></td>
	</tr>
</table>
<%
sql="select id from [topic] order by usertime desc"
set rs=server.createobject("ADODB.recordset")
rs.open sql,conn,1,1

dim totalrec
totalrec=rs.recordcount
%>

<table width="770" height="30" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <%if not session("login")="true" then%>
  <form method='post' action='admin.asp?act=admchk' name='form1' id='form1'>
	<tr>
		<td valign="middle" style="border-bottom: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#F1F1F1', endcolorstr='#FFFFFF', gradienttype='1'">&nbsp;&nbsp;��ӭ����&nbsp;<b><a href="<%=URL%>"><%=site%></a></b>&nbsp;&nbsp;����Ա��<b><a href="mailto:<%=adminmail%>"><%=name%></a></b>&nbsp;&nbsp;���룺<input type="password" name='adminpass' size="8">&nbsp;<input type="submit" value="�����½">&nbsp;&nbsp;&nbsp;����ͳ�ƣ�<b><%=stat%></b>&nbsp;�Σ�&nbsp;&nbsp;����������<b><%=totalrec%></b>&nbsp;��</td>
	</tr></form>
<%
else
%>
	<tr>
		<td valign="middle" colspan="2" style="border-bottom: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#F1F1F1', endcolorstr='#FFFFFF', gradienttype='1'">&nbsp;&nbsp;&nbsp;��ӭ&nbsp;<b><%=name%></b>��<a href="bulletin.asp">��������</a>��<a href="index.asp">��������</a>&nbsp;<a href="admin.asp?act=batch">����</a>��<a href="admin.asp?act=main">�������Ա�</a>��<a href="admin.asp?act=logout"><font color="red"><b>�˳���½</b></font></a>��&nbsp;&nbsp;����ͳ�ƣ�<b><%=stat%></b>&nbsp;�Σ�&nbsp;&nbsp;����������<b><%=totalrec%></b>&nbsp;��
	</td>
	</tr>
<%
end if
%>
</table>