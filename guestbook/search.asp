<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		search.asp
'**
'** �ļ�˵��������ҳ��
'** �޸����ڣ�2005-04-07
'**************************************

dim key
	key=sql_filter(left(Trim(Request.QueryString("key")),20))
dim currentpage,page_count,Pcount
dim totalrec,endpage
if Request.QueryString("page")="" then
	currentPage=1
else
	if (not isnumeric(Request.QueryString("page"))) then
		errinfo="<li>�Ƿ���ҳ�������"
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
pagename="��������"
call pageinfo()

call skin1()
'---------------������ʾҳ������--------
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="460"><tr><td width='100%'><p align="center"><B>������Ҫ���������ݣ�</B><br><br><form action="search.asp" method="POST"><input type='text' name='key' size='20'>&nbsp;<input type='submit' value='����'></form>������Χ�������������ߵĳƺ������Ա��⡢�����Լ��ظ���</td></tr></table>
<%
'--------------ҳ��������ʾ����--------
call skin2()
end sub

sub main()

if sql_filter(Trim(Request.Form("key")))="" and key="" then
	errinfo="<li>�����������ؼ��֣�"
	call error()
else
	if (not sql_filter(Trim(Request.Form("key")))="") and key=""	then
		Response.Redirect "?key="&Trim(Request.Form("key"))
		Response.Flush
	end if
end if

pagename="�������"
call pageinfo()

call skin1()
'---------------������ʾҳ������--------
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
		<td width='100%'><p align="center">û���ҵ�������<B><%=back_filter(key)%></B>�������ԣ���򻯹ؼ��ֺ���������<p align="center"><a href="javascript:history.back(1)"><B>&lt;&lt; ����</B></p></td>
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

response.write "<table width='75%' cellpadding=15 cellspacing=1 align=center><tr><td valign='top' align=center class=tablebody1>���ҵ� <b>"& totalrec &"</b> �������ַ���<b>"&back_filter(key)&"</b>�������ԣ�</td></tr>"
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
					<%if rs("top")=1 then%><font color="<%=maincolor%>"><b>����Ա����</b></font><%else%><b><font COLOR="#000000"><%=Boldkey(HTMLencode(rs("username")),key)%></font></b><%end if%></td>
				</tr>
				<tr>
					<td width="100%" colspan="2" align="center">
					<img SRC="images/blank.gif" WIDTH="160" HEIGHT="10" BORDER="0"></td>
				</tr>
				<%if not rs("userURL")="" then%><tr>
					<td width="10%" align="right">
					<img border="0" width="18" height="18" src="images/homepage.gif"></td>
					<td width="90%" align="left"><a target="_blank" href="<%=HTMLencode(rs("userURL"))%>" title="���ʡ�<%=HTMLencode(rs("username"))%>���ĸ�����ҳ"><%=userURL%></a></td>
				</tr><%end if%>
				<%if not rs("usermail")="" then%><tr>
					<td width="10%" align="right"><img border="0" width="18" height="18" src="images/email.gif"></td>
					<td width="90%" align="left"><a href="mailto:<%=HTMLencode(rs("usermail"))%>" title="����<%=HTMLencode(rs("username"))%>�����͵����ʼ�"><%=usermail%></a></td>
				</tr><%end if%>
				<%if not rs("userqq")="" then%><tr>
					<td width="10%" align="right"><img border="0" width="18" height="18" src="images/qq.gif"></td>
					<td width="90%" align="left"><a target="blank" href="http://wpa.qq.com/msgrd?V=1&Uin=<%=rs("userqq")%>&Site=<%=site%>&Menu=yes" title="��<%=HTMLencode(rs("username"))%>����QQ������ʱ�Ự��QQ����뿪����"><%=rs("userqq")%></a></td>
				</tr><%end if%>
				<tr>
					<td width="100%" height="10" colspan="2"> </td>
				</tr>
			</table>
			</td>
			<td width="100%" class="tablebody3">���⣺<b><%if rs("whisper")="1" and not session("login")="true" then%><font COLOR="<%=maincolor%>">���Ļ�����</font><%else%><%=Boldkey(HTMLencode(rs("usertitle")),key)%><%end if%></b><br>
			ʱ�䣺<font face="Verdana" SIZE="1"><%=rs("usertime")%></font></td>
		</tr>
		<tr>
			<td class="tablebody3" width="100%" height="100%" valign="top">
			<table border="0" cellpadding="5" cellspacing="5" width="100%">
				<tr>
					<td width="100%" height="110" valign="top">
					<%if rs("whisper")="1" and (not session("login")="true") then%><br>
					<p align="right"><img border="0" src="images/whisper.gif"></p>
					<p align="right">״̬��<%if not rs("replycode")="" then%>
					<%if rs("reply")=1 then%> <b>�ѱ��ظ�</b></p>
					<p align="right"><%call viewreply(rs("id"))
					else%><b>δ���ظ�</b> <%end if
					else%><b>�޷����ظ�</b> <%end if%> </p>
					<%else
					Response.Write Boldkey(UBBCode(rs("usercontent"),rs("top")),key)
					if rs("reply")=1 then%><p></p>
					<table border="0" align="center" cellpadding="5" cellspacing="1" width="95%" class="table1">
						<tr>
							<td width="100%" class="tablebody1"><font COLOR="<%=maincolor%>">
							����Ա�ظ���<br>
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
			<%if rs("checked")=0 then%><a href="javascript:go('admin.asp?act=check&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','��ȷ��Ҫͨ����ˣ�')"><font COLOR="red"><b>ͨ�����</b></font></a>&nbsp;&nbsp;<%end if%><a href="javascript:go('admin.asp?act=del&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','��ȷ��Ҫɾ����')">ɾ��</a>&nbsp;&nbsp;<%if rs("whisper")=1 and rs("replycode")="" then%><font COLOR="red"><b>�޷��ظ������Ļ�</b></font><%else%><a href="JavaScript:openwin('reply.asp?id=<%=rs("id")%>',600,500)"><%if rs("whisper")=1 then%><font COLOR="red"><b>���Ļ��ظ�/�༭�ظ�</b></font><%else%>�ظ�/�༭�ظ�<%end if%></a><%end if%>&nbsp;&nbsp;<a href="JavaScript:openwin('edit.asp?id=<%=rs("id")%>',600,500)">�༭</a>&nbsp;&nbsp;����IP��<%=rs("ip")%>
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
'--------------ҳ��������ʾ����--------
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
"<td valign=middle align=right>ҳ�Σ�<b>"& currentPage &"/"& n &"</b>ҳ��ÿҳ<b>"& rs.PageSize &"</b>������<b>"& totalrec &"</b>��&nbsp;&nbsp;"

if currentPage=1 then
	response.write "<font face=webdings>9</font>	 "
else
	response.write "<a href='?key="&key&"&page=1' title=��ҳ><font face=webdings>9</font></a>	 "
end if
if p*10>0 then response.write "<a href='?key="&key&"&page="&Cstr(p*10)&"' title=��ʮҳ><font face=webdings>7</font></a>	 "
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
if ii<n then response.write "<a href='?key="&key&"&page="&Cstr(ii)&"' title=��ʮҳ><font face=webdings>8</font></a>	 "
if currentPage=n then
	response.write "<font face=webdings>:</font>	 "
else
	response.write "<a href='?key="&key&"&page="&Cstr(n)&"' title=βҳ><font face=webdings>:</font></a>	 "
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