<% Option Explicit %>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		preview.asp
'**
'** 文件说明：留言预览页面
'** 修改日期：2005-04-07
'** 作者：Howlion
'** Email：howlion@163.com
'**************************************

pagename="留言预览"
call pageinfo()

dim UBB_super,usercontent
	UBB_super=request.form("UBB_super")

dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\[(.[^\]]*)\])"

if len(request.form("usercontent"))>maxnum and UBB_super=0 then
	usercontent="<font size='3' color='red'>留言内容的字数超过了限制！</font>"
elseif re.Replace(Replace(request.form("usercontent"), CHR(13)&CHR(10), ""),"")="" then
	usercontent="<font size='3' color='red'>留言为空！</font>"
else
	usercontent=UBBCode(sql_filter(request.form("usercontent")),UBB_super)
end if
set re=nothing
%>
<table cellpadding="5" cellspacing="1" width="550" align="center" class="table1">
	<tr>
		<td width="100%" class="tablebody3">标题：<b><%=HTMLencode(request.form("usertitle"))%></b>
		</td>
	</tr>
	<tr>
		<td width="100%" class="tablebody1" valign="top">
		<%=usercontent%></td>
	</tr>
	<tr>
		<td width="100%" class="tablebody3" align="center">
		<a href="javascript:window.close()">关闭窗口</a> </td>
	</tr>
</table>