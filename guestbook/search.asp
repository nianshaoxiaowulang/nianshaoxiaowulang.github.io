<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		search.asp
'**
'** 文件说明：搜索页面
'** 修改日期：2005-04-07
'**************************************

dim key
	key=sql_filter(left(Trim(Request.QueryString("key")),20))
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

mainpic="page_search.gif"

Select Case Request.QueryString("act")
	case "fillform"
	call fillform()
	case else
	call main()
end Select

sub fillform()
pagename="搜索留言"
call pageinfo()

call skin1()
'---------------以下显示页面主体--------
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="460"><tr><td width='100%'><p align="center"><B>请输入要搜索的内容：</B><br><br><form action="search.asp" method="POST"><input type='text' name='key' size='20'>&nbsp;<input type='submit' value='搜索'></form>搜索范围将包括：留言者的称呼、留言标题、正文以及回复。</td></tr></table>
<%
'--------------页面主题显示结束--------
call skin2()
end sub

sub main()

if sql_filter(Trim(Request.Form("key")))="" and key="" then
	errinfo="<li>请输入搜索关键字！"
	call error()
else
	if (not sql_filter(Trim(Request.Form("key")))="") and key=""	then
		Response.Redirect "?key="&Trim(Request.Form("key"))
		Response.Flush
	end if
end if

pagename="搜索结果"
call pageinfo()

call skin1()
'---------------以下显示页面主体--------
%>
<SCRIPT language=JavaScript>
<!--
function go(src,q)
{
	var ret;
	ret = confirm(q);
	if(ret!=false)window.location=src;
}

function openwin(URL, width, height){
	var Win = window.open(URL,"openScript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=1,menubar=0,status=1');}
-->
</script>
<%
if session("login")="true" then
	sql = "Select * from [topic] where (usertitle like '%"&key&"%' or usercontent like '%"&key&"%' or username like '%"&key&"%' or recontent like '%"&key&"%') order by top desc,usertime desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
else
	sql = "Select * from [topic] where checked=1 and whisper=0 and (usertitle like '%"&key&"%' or usercontent like '%"&key&"%' or username like '%"&key&"%' or recontent like '%"&key&"%') order by top desc,usertime desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
end if

if rs.eof or rs.bof then
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="460">
	<tr>
		<td width='100%'><p align="center">没有找到包含“<B><%=back_filter(key)%></B>”的留言，请简化关键字后再搜索。<p align="center"><a href="javascript:history.back(1)"><B>&lt;&lt; 返回</B></p></td>
	</tr>
</table>
<%
	set rs=nothing
	call skin2()
	response.end
end if

	rs.PageSize = perpage
	rs.AbsolutePage=currentpage
	page_count=0
	totalrec=rs.recordcount

if not totalrec mod perpage=0 then
	if currentPage > (totalrec/perpage)+1 then response.redirect "?key="&key&"&page=" & Int((totalrec/perpage))+1
else
	if currentPage > (totalrec/perpage) then response.redirect "?key="&key&"&page=" & Int((totalrec/perpage))
end if

response.write "<table width='75%' cellpadding=15 cellspacing=1 align=center><tr><td valign='top' align=center class=tablebody1>共找到 <b>"& totalrec &"</b> 条包含字符“<b>"&back_filter(key)&"</b>”的留言：</td></tr>"
call pages()

while (not rs.eof) and (not page_count = rs.PageSize)

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
					<%if rs("top")=1 then%><font color="<%=maincolor%>"><b>管理员公告</b></font><%else%><b><font COLOR="#000000"><%=Boldkey(HTMLencode(rs("username")),key)%></font></b><%end if%></td>
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
			<td width="100%" class="tablebody3">标题：<b><%if rs("whisper")="1" and not session("login")="true" then%><font COLOR="<%=maincolor%>">悄悄话留言</font><%else%><%=Boldkey(HTMLencode(rs("usertitle")),key)%><%end if%></b><br>
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
					Response.Write Boldkey(UBBCode(rs("usercontent"),rs("top")),key)
					if rs("reply")=1 then%><p></p>
					<table border="0" align="center" cellpadding="5" cellspacing="1" width="95%" class="table1">
						<tr>
							<td width="100%" class="tablebody1"><font COLOR="<%=maincolor%>">
							管理员回复：<br>
							<font SIZE="1" face="Verdana"><%=rs("retime")%></font>
							<hr color="<%=maincolor%>" align="left" width="20%" size="1">
							<%=Boldkey(UBBCode(rs("recontent"),1),key)%> </font></td>
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
'--------------页面主题显示结束--------
call skin2()
end sub

sub pages()
dim ii,p,n
if totalrec mod perpage=0 then
	n= totalrec \ perpage
else
	n= totalrec \ perpage+1
end if
p=(currentpage-1) \ 10
response.write "<table border=0 cellpadding=0 cellspacing=3 width='86%' align=center>"&_
"<tr>"&_
"<td valign=middle align=right>页次：<b>"& currentPage &"/"& n &"</b>页，每页<b>"& rs.PageSize &"</b>条，共<b>"& totalrec &"</b>条&nbsp;&nbsp;"

if currentPage=1 then
	response.write "<font face=webdings>9</font>	 "
else
	response.write "<a href='?key="&key&"&page=1' title=首页><font face=webdings>9</font></a>	 "
end if
if p*10>0 then response.write "<a href='?key="&key&"&page="&Cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>	 "
response.write "<b>"
for ii=p*10+1 to P*10+10
	if ii=currentPage then
		response.write "<font size=4>"+Cstr(ii)+"</font> "
	else
		response.write "<a href='?key="&key&"&page="&Cstr(ii)&"'>"+Cstr(ii)+"</a>	 "
	end if
	if ii=n then exit for
	'p=p+1
next
response.write "</b>"
if ii<n then response.write "<a href='?key="&key&"&page="&Cstr(ii)&"' title=下十页><font face=webdings>8</font></a>	 "
if currentPage=n then
	response.write "<font face=webdings>:</font>	 "
else
	response.write "<a href='?key="&key&"&page="&Cstr(n)&"' title=尾页><font face=webdings>:</font></a>	 "
end if
response.write "</table>"
end sub

Function Boldkey(strContent,key)
	dim B_key
	Set B_key=new RegExp
	B_key.IgnoreCase =true
	B_key.Global=True


	B_key.Pattern="(" & key & ")"
	strContent=B_key.Replace(strContent,"<B style='color: white; background-color: #CC3333'>$1</B>" )

	Set B_key=Nothing
	Boldkey=strContent
End Function
%>