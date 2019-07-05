<%
'**************************************
'**		inc_top.asp
'**
'** 文件说明：留言本顶部信息
'** 修改日期：2005-04-07
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
		<td valign="middle" style="border-bottom: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#F1F1F1', endcolorstr='#FFFFFF', gradienttype='1'">&nbsp;&nbsp;欢迎来到&nbsp;<b><a href="<%=URL%>"><%=site%></a></b>&nbsp;&nbsp;管理员：<b><a href="mailto:<%=adminmail%>"><%=name%></a></b>&nbsp;&nbsp;密码：<input type="password" name='adminpass' size="8">&nbsp;<input type="submit" value="管理登陆">&nbsp;&nbsp;&nbsp;访问统计：<b><%=stat%></b>&nbsp;次，&nbsp;&nbsp;留言总数：<b><%=totalrec%></b>&nbsp;条</td>
	</tr></form>
<%
else
%>
	<tr>
		<td valign="middle" colspan="2" style="border-bottom: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#F1F1F1', endcolorstr='#FFFFFF', gradienttype='1'">&nbsp;&nbsp;&nbsp;欢迎&nbsp;<b><%=name%></b>┃<a href="bulletin.asp">发布公告</a>┃<a href="index.asp">管理留言</a>&nbsp;<a href="admin.asp?act=batch">批量</a>┃<a href="admin.asp?act=main">设置留言本</a>┃<a href="admin.asp?act=logout"><font color="red"><b>退出登陆</b></font></a>┃&nbsp;&nbsp;访问统计：<b><%=stat%></b>&nbsp;次，&nbsp;&nbsp;留言总数：<b><%=totalrec%></b>&nbsp;条
	</td>
	</tr>
<%
end if
%>
</table>