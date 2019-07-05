<% option explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		bulletin.asp
'**
'** 文件说明：公告发布页面
'** 修改日期：2005-04-07
'** 作者：Howlion
'** email：howlion@163.com
'**************************************

if not session("login")="true" then
	errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
	call error()
end if

select case request.querystring("act")
	case "addnew"
	call addnew()
	case else
	call main()
end select

sub main()
	pagename="发布公告"
	call pageinfo()
	mainpic="page_bulletin.gif"
	call skin1()
'---------------以下显示页面主体--------
%>
<script src="UBB.js"></script>
<br>
<form action="?act=addnew" method="post" name="lw_form">
	<table align="center" cellpadding="3" cellspacing="1" class="table1" width="95%">
		<tr>
			<td width="20%" class="tablebody3" align="right">管理QQ</td>
			<td width="80%" class="tablebody2">
			<input name="userqq" size="19" maxlength="80"></td>
		</tr>
		<tr>
			<td width="20%" class="tablebody3" align="right">
			<font face="宋体" color="red">***</font> <b>公告标题</b></td>
			<td width="80%" class="tablebody2">
			<input name="usertitle" size="40" maxlength="100"></td>
		</tr>
		<tr>
			<td width="20%" valign="top" class="tablebody3" align="right">
			<font face="宋体" color="red">***</font> <b>公告正文</b></td>
			<td width="80%" class="tablebody2">
			<!--#include file="inc_UBB.asp"-->
			<textarea class="smallarea" cols="70" name="usercontent" title="ctrl+enter提交" rows="12" onkeydown="ctlent()"></textarea><br>
			点击表情符号可以将其加入正文：<br>
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
			<input type="hidden" name="UBB_super" value="1">
			<input type="submit" name="submit" value="提 交">&nbsp;&nbsp;
			<input type="reset" name="submit2" value="清 除">&nbsp;&nbsp;
			<input type="button" name="preview" value="预览" onclick="openpreview()">
			</td>
		</tr>
	</table>
	</div>
</form>
<form name="preview" action="preview.asp" method="post" target="preview_page">
	<input type="hidden" name="UBB_super" value>
	<input type="hidden" name="usertitle" value>
	<input type="hidden" name="usercontent" value>
</form>
<br>
<%
'--------------页面主体显示结束--------
	call skin2()
end sub

sub addnew()

	dim username,userURL,usermail,userqq,usertitle,usercontent

		userqq=sql_filter(trim(request.form("userqq")))
		usertitle=trim(request.form("usertitle"))
		usercontent=rtrim(request.form("usercontent"))

	if usertitle="" then
		errinfo=errinfo & "<li>未填写标题"
		elseif len(usertitle)>50 then
		errinfo=errinfo & "<li>过长的标题"
	end if

	if usercontent="" then
		errinfo=errinfo & "<li>未填写留言内容"
	end if

	if trim(userqq)<>"" then
		if not(isnumeric(userqq)) then
			errinfo=errinfo & "<li>QQ号码填写有误"
		end if
	end if

call error()

	set rs= server.createobject("adodb.recordset")
	sql="select * from [topic]"
	rs.open sql,conn,3,2
	rs.addnew
	rs("username")=name
	rs("xingbie")=0
	rs("userURL")=URL
	rs("usermail")=adminmail
	rs("userqq")=userqq
	rs("usertime")=now()
	rs("usertitle")=usertitle
	rs("usercontent")=usercontent
	rs("top")="1"
	rs("reply")="0"
	rs("ip")=ip
	rs("checked")=1
	rs("whisper")=0
	rs.update
	rs.close
	response.redirect "index.asp"
	response.flush

end sub
%>