<% option explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		index.asp
'** 文件说明：留言本首页
'** 修改日期：2005-04-07
'**************************************

dim currentpage,page_count,Pcount
dim totalrec,endpage
if Request.QueryString("page")="" then
	currentPage=1
else
	if (not isnumeric(Request.QueryString("page"))) then
		errinfo="<li>非法的页面参数！"
		call error()
	end if
	currentPage=cint(Request.QueryString("page"))
end if

if session("login")="true" then
	pagename="管理留言"
	mainpic="page_admin_lw.gif"
else
	pagename="查看留言"
	mainpic="page_index.gif"
end if

call pageinfo()
call skin1()
'---------------以下显示页面主体--------
%>

<script language=javascript>
<!--
function go(src,q){
	var ret;
	ret = confirm(q);
	if(ret!=false)window.location=src;
}

function openwin(URL, width, height){
	var win = window.open(URL,"openscript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=1,menubar=0,status=1');
}

function openreply(){
	document.viewreply.replycodes.value=document.replyform.replycode.value;
	var popupwin = window.open('viewreply.asp', 'viewreply_page', 'scrollbars=yes,width=700,height=450');
	document.viewreply.submit()
}
//-->
</script>
<!--#include file="UBB.asp"-->
<%
if session("login")="true" then
	sql="select * from [topic] order by top desc,usertime desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
else
	sql="select * from [topic] where checked=1 order by top desc,usertime desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
end if

if rs.eof and rs.bof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="390">
	<tr>
		<td valign="middle" align="center"><font size="4">暂时没有留言，欢迎您留言！</font></td>
	</tr>
</table>
<%
	set rs=nothing
	call skin2()
	response.end
end if

	rs.pagesize = perpage
	rs.absolutepage=currentpage
	page_count=0
	totalrec=rs.recordcount

if not totalrec mod perpage=0 then
	if currentPage > (totalrec/perpage)+1 then response.redirect "?page=" & Int((totalrec/perpage))+1
else
	if currentPage > (totalrec/perpage) then response.redirect "?page=" & Int((totalrec/perpage))
end if

call pages()

while (not rs.eof) and (not page_count = rs.pagesize)

dim userURL,usermail
	if len(HTMLencode(rs("userURL")))>22 then
		userURL = left(HTMLencode(rs("userURL")),22)&"..."
	else
		userURL = HTMLencode(rs("userURL"))
	end if

	if len(HTMLencode(rs("usermail")))>22 then
		usermail = left(HTMLencode(rs("usermail")),22)&"..."
	else
		usermail = HTMLencode(rs("usermail"))
	end if
%>
<div align="center">
	<table border="0" cellpadding="5" cellspacing="1" width="95%" class="table1">
		<tr>
			<td width="180" rowspan="2" class="tablebody3" align="center" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td width="100%" colspan="2" align="center"><br><%if rs("top")=1 then%><img border="0" width="60" height="60" src="images/bulletin.gif"><%else%><img border="0" width="90" height="90" src="images/userfaces/<%=rs("userface")%>.gif" style="border: 1 solid #000000"><%end if%><br>
					<br>
					<%if rs("top")=1 then%><font color="<%=maincolor%>"><b>管理员公告</b></font><%else%><b><font COLOR="#000000"><%=HTMLencode(rs("username"))%></font></b><%end if%></td>
				</tr>
				<tr>
					<td width="100%" colspan="2" align="center">
					<img SRC="images/blank.gif" WIDTH="160" HEIGHT="10" BORDER="0"></td>
				</tr>
				<%if not rs("userURL")="" then%><tr>
					<td width="10%" align="right">
					<img border="0" width="18" height="18" src="images/homepage.gif"></td>
					<td width="90%" align="left"><a target="_blank" href="<%=HTMLencode(rs("userURL"))%>" title="访问“<%=HTMLencode(rs("username"))%>”的个人主页"><%=userURL%></a></td>
				</tr><%end if%>
				<%if not rs("usermail")="" then%><tr>
					<td width="10%" align="right"><img border="0" width="18" height="18" src="images/email.gif"></td>
					<td width="90%" align="left"><a href="mailto:<%=HTMLencode(rs("usermail"))%>" title="给“<%=HTMLencode(rs("username"))%>”发送电子邮件"><%=usermail%></a></td>
				</tr><%end if%>
				<%if not rs("userqq")="" then%><tr>
					<td width="10%" align="right"><img border="0" width="18" height="18" src="images/qq.gif"></td>
					<td width="90%" align="left"><a target="blank" href="http://wpa.qq.com/msgrd?V=1&Uin=<%=rs("userqq")%>&Site=<%=site%>&Menu=yes" title="向“<%=HTMLencode(rs("username"))%>”的QQ发起临时会话（QQ软件须开启）"><%=rs("userqq")%></a></td>
				</tr><%end if%>
				<tr>
					<td width="100%" height="10" colspan="2"> </td>
				</tr>
			</table>
			</td>
			<td width="100%" class="tablebody3">标题：<b><%if rs("whisper")="1" and not session("login")="true" then%><font COLOR="<%=maincolor%>">悄悄话留言</font><%else%><%=HTMLencode(rs("usertitle"))%><%end if%></b><br>
			时间：<font face="Verdana" SIZE="1"><%=rs("usertime")%></font></td>
		</tr>
		<tr>
			<td class="tablebody3" width="100%" height="100%" valign="top">
			<table border="0" cellpadding="5" cellspacing="5" width="100%">
				<tr>
					<td width="100%" height="110" valign="top">
					<%if rs("whisper")="1" and (not session("login")="true") then%><br>
					<p align="right"><img border="0" src="images/whisper.gif"></p>
					<p align="right">状态：<%if not rs("replycode")="" then%>
					<%if rs("reply")=1 then%> <b>已被回复</b></p>
					<p align="right"><%call viewreply(rs("id"))
					else%><b>未被回复</b> <%end if
					else%><b>无法被回复</b> <%end if%> </p>
					<%else
					Response.Write UBBCode(rs("usercontent"),rs("top"))
					if rs("reply")=1 then%><p></p>
					<table border="0" align="center" cellpadding="5" cellspacing="1" width="95%" class="table1">
						<tr>
							<td width="100%" class="tablebody1"><font COLOR="<%=maincolor%>">
							管理员回复：<br>
							<font SIZE="1" face="Verdana"><%=rs("retime")%></font>
							<hr color="<%=maincolor%>" align="left" width="20%" size="1">
							<%=UBBCode(rs("recontent"),1)%> </font></td>
						</tr>
					</table>
					<%end if%><%end if%></td>
				</tr>
			</table>
			</td>
		</tr>
		<%if session("login")="true" then%>
		<tr>
			<td class="tablebody3" colspan="2" width="17%" align="right">
			<%if rs("checked")=0 then%><a href="javascript:go('admin.asp?act=check&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','您确定要通过审核？')"><font COLOR="red"><b>通过审核</b></font></a>&nbsp;&nbsp;<%end if%><a href="javascript:go('admin.asp?act=del&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','您确定要删除？')">删除</a>&nbsp;&nbsp;<%if rs("whisper")=1 and rs("replycode")="" then%><font COLOR="red"><b>无法回复的悄悄话</b></font><%else%><a href="JavaScript:openwin('reply.asp?id=<%=rs("id")%>',600,500)"><%if rs("whisper")=1 then%><font COLOR="red"><b>悄悄话回复/编辑回复</b></font><%else%>回复/编辑回复<%end if%></a><%end if%>&nbsp;&nbsp;<a href="JavaScript:openwin('edit.asp?id=<%=rs("id")%>',600,500)">编辑</a>&nbsp;&nbsp;留言IP：<%=rs("ip")%>
			</td>
		</tr>
		<%end if%>
	</table>
</div>
<p>
<%
	page_count = page_count + 1
	rs.movenext
wend
call pages()
rs.close
set rs=nothing

sub viewreply(id) '------查看回复的表单------------
%>
			<script language=javascript>
			<!--
			function submitcheck_<%=id%>(){
			if (document.viewreply_<%=id%>.replycode.value.length==0){
			alert("请先输入回复查看码！");
			document.viewreply_<%=id%>.replycode.focus();
			return false;
			}
			return true
			}

			function openreply_<%=id%>()
			{
			document.viewreply_<%=id%>.replycode.value=document.replyform_<%=id%>.replycode.value;
			var popupwin = window.open('viewreply.asp?id=<%=id%>', 'viewreply_page', 'resizable=0,scrollbars=1,menubar=0,status=1,width=600,height=500');
			document.viewreply_<%=id%>.submit()
			}
			-->
			</script>

			<form action="" method="post" name="replyform_<%=id%>">
			回复查看码：<input name="replycode" type="password" size="10" maxlength="100">&nbsp;
			<input type="button" name="viewreply" value="查看回复" onclick="openreply_<%=id%>()">
			</form>
			<form name="viewreply_<%=id%>" action="viewreply.asp?id=<%=id%>" method="post" target="viewreply_page">
			<input type="hidden" name="replycode" value="">
			</form>
<%
end sub

sub pages()	'------分页代码------------
dim ii,p,n
if totalrec mod perpage=0 then
	n= totalrec \ perpage
else
	n= totalrec \ perpage+1
end if
p=(currentpage-1) \ 10
response.write "<table border=0 cellpadding=0 cellspacing=3 width='86%' align=center>"&_
"<tr>"&_
"<td valign=middle align=right>页次：<b>"& currentpage &"/"& n &"</b>页，每页<b>"& rs.pagesize &"</b>条，共<b>"& totalrec &"</b>条&nbsp;&nbsp;"

if currentpage=1 then
	response.write "<font face=webdings>9</font>	 "
else
	response.write "<a href='index.asp?page=1' title=首页><font face=webdings>9</font></a>	 "
end if
if p*10>0 then response.write "<a href='index.asp?page="&cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>	 "
response.write "<b>"
for ii=p*10+1 to p*10+10
	if ii=currentpage then
		response.write "<font size=4>"+cstr(ii)+"</font> "
	else
		response.write "<a href='index.asp?page="&cstr(ii)&"'>"+cstr(ii)+"</a>	 "
	end if
	if ii=n then exit for
	'p=p+1
next
response.write "</b>"
if ii<n then response.write "<a href='index.asp?page="&cstr(ii)&"' title=下十页><font face=webdings>8</font></a>	 "
if currentpage=n then
	response.write "<font face=webdings>:</font>	 "
else
	response.write "<a href='index.asp?page="&cstr(n)&"' title=尾页><font face=webdings>:</font></a>	 "
end if
response.write "</table>"
end sub
'--------------页面主题显示结束--------
call skin2()
%>