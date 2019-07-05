<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		new.asp
'**
'** 文件说明：发表留言页面
'** 修改日期：2005-04-07
'**************************************

if lock="1" then
	errinfo="<li>留言本已经被管理员锁定，您无法留言。"
	call error()
end if

select case Request.QueryString("act")
	case "addnew"
	call addnew()
	case else
	call main()
end select

sub main()

pagename="写留言"
call pageinfo()
mainpic="page_new.gif"
call skin1()
'---------------以下显示页面主体--------
%>
<script language="JavaScript">
<!--
function Submitcheck(){
	if (document.lw_form.username.value.length==0){
	alert("请输入您的称呼，此为必填项！");
	document.lw_form.username.focus();
	return false;
}
	if (document.lw_form.usercontent.value.length==0){
	alert("请输入留言正文，此为必填项！");
	document.lw_form.usercontent.focus();
	return false;
	}
	return true
}
//-->
</script>
<script src="UBB.js"></script>
<br>
<form action="?act=addnew" method="POST" onSubmit="return Submitcheck()" name="lw_form">
	<div align="center">
		<table align="center" width="95%">
			<tr>
				<td><%=UBBcode(notice,1)%></td>
			</tr>
		</table>
		<table align="center" cellpadding="3" cellspacing="1" class="table1" width="95%">
			<tr>
				<td width="20%" class="tablebody3" align="right">
				<img border="0" src="images/perinfo.gif"></td>
				<td width="80%" class="tablebody2"><b>个人信息</b></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">
				<font FACE="宋体" COLOR="red">***</font> <b>您的称呼：</b></td>
				<td width="80%" class="tablebody2">
				<input name="username" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" valign="top" align="right">
				性别及头像：</td>
				<td width="80%" class="tablebody2">男性：<img src="images/userfaces/small_1.gif" width="30" height="30"><input type="radio" value="1" name="userface" checked>&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_2.gif" width="30" height="30"><input type="radio" value="2" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_3.gif" width="30" height="30"><input type="radio" value="3" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_4.gif" width="30" height="30"><input type="radio" value="4" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_5.gif" width="30" height="30"><input type="radio" value="5" name="userface"><br>
				女性：<img src="images/userfaces/small_6.gif" width="30" height="30"><input type="radio" value="6" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_7.gif" width="30" height="30"><input type="radio" value="7" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_8.gif" width="30" height="30"><input type="radio" value="8" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_9.gif" width="30" height="30"><input type="radio" value="9" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_10.gif" width="30" height="30"><input type="radio" value="10" name="userface"><br>
				中性：<img src="images/userfaces/small_11.gif" width="30" height="30"><input type="radio" value="11" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_12.gif" width="30" height="30"><input type="radio" value="12" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_13.gif" width="30" height="30"><input type="radio" value="13" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_14.gif" width="30" height="30"><input type="radio" value="14" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_15.gif" width="30" height="30"><input type="radio" value="15" name="userface"><p>
				<b>提示：</b>如果不想透露自己的性别，可以选一个中性的头像</b></font></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">个人主页：</td>
				<td width="80%" class="tablebody2">
				<input name="userURL" size="19" maxlength="80" value="http://"><br>
				如果要填写，应填写完整地址，如：http://www.howlion.com/</td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">Email：</td>
				<td width="80%" class="tablebody2">
				<input name="usermail" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">QQ：</td>
				<td width="80%" class="tablebody2">
				<input name="userqq" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">
				<img border="0" src="images/lwcontent.gif"></td>
				<td width="80%" class="tablebody2"><b>留言内容</b></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">标题：</td>
				<td width="80%" class="tablebody2">
				<input name="usertitle" size="40" maxlength="100"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" valign="top" align="right">
				<font FACE="宋体" COLOR="red">***</font> <b>正文：</b></td>
				<td width="80%" class="tablebody2">
				<!--#include file="inc_UBB.asp"-->
				<textarea cols="70" name="usercontent" title="Ctrl+Enter提交" rows="12" onkeydown="ctlent()"></textarea><br>正文内容不能大于<b><font size="3"><%=maxnum%></font></b>个字符<%if UBBcfg_face=1 then%><br>
				点击表情符号可以将其加入正文：<br>
				<%
dim ii,i
for i=1 to 42
	if len(i)=1 then ii="0" & i else ii=i
	response.write "<img src=""images/faces/"&ii&".gif"" width=20 height=20 border=0 onclick=""insertsmilie('[face"&ii&"]')"" style=""CURSOR: hand"">&nbsp;"
	if i=17 or i=34 then response.write "<br>"
next
end if%> </td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right"></td>
				<td width="80%" class="tablebody2"><input type="checkbox" name="whisper" value="1">悄悄话&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;回复查看码（可不填）：<input name="replycode" size="20" maxlength="100"><p>
				<b>提示：</b>回复查看码用于以后查看管理员对悄悄话的回复。可不填，但管理员将无法回复您的悄悄话。</td>
			</tr>
			<tr>
				<td valign="middle" colspan="2" align="center" class="tablebody1">
				<input type="hidden" name="UBB_super" value="0">
				<input type="Submit" name="Submit" value="提 交">&nbsp;&nbsp;
				<input type="reset" name="Submit2" value="清 除">&nbsp;&nbsp;
				<input type="button" name="Preview" value="预览" onclick="openpreview()">
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
'--------------页面主题显示结束--------
call skin2()
end sub


sub addnew()	'将留言加入数据库
	dim servername1,servername2
		servername1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		servername2=Cstr(Request.ServerVariables("SERVER_NAME"))

		if mid(servername1,8,len(servername2))<>servername2 then
			errinfo="<li>非法的提交动作！"
		end if

	dim username,xingbie,userface,userURL,usermail,userqq
	dim usertitle,usercontent,whisper,checked,replycode
	username=sql_filter(Trim(request.Form("username")))

	if request.Form("userface") < 6 then
		xingbie=1
		elseif request.Form("userface") < 11 then
		xingbie=2
		else
		xingbie=3
	end if

	userface=sql_filter(request.Form("userface"))

	if sql_filter(Trim(request.form("userURL")))="http://" then
		userURL=""
	else
		userURL=sql_filter(Trim(request.form("userURL")))
	end if

		usermail=sql_filter(Trim(request.form("usermail")))
		userqq=sql_filter(Trim(request.form("userqq")))
		usertitle=sql_filter(Trim(request.form("usertitle")))
		usercontent=sql_filter(Rtrim(request.form("usercontent")))

	if not request.form("whisper")="1" then
		whisper=0
	else
		whisper=1
	end if

		replycode=sql_filter(Trim(request.form("replycode")))

	if username="" then
		errinfo=errinfo & "<li>未填写您的称呼"
		elseif len(username)>20 then
		errinfo=errinfo & "<li>过长的称呼"
	end if

	if len(usertitle)>50 then
		errinfo=errinfo & "<li>过长的标题"
	end if

	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\[(.[^\]]*)\])"
	if	Trim(re.Replace(Replace(request.form("usercontent"), CHR(13)&CHR(10), ""),""))="" then
		errinfo=errinfo & "<li>未填写留言内容"
		elseif len(usercontent)>maxnum then
		errinfo=errinfo & "<li>过长的留言内容"
	end if

	if not adword="" then
		dim alladword,i
			alladword=split(adword,chr(13)&chr(10))
		for i = lbound(alladword) to ubound(alladword)
			if instr(UCase(usercontent & usertitle),UCase(trim(alladword(i))))>0 then
				errinfo="<li>发生诡异错误"	'注：此处为迷惑发广告的人
				call error()
				response.end
			end if
		next
	end if

	If userURL<>"" then
		dim isURL
		re.Pattern="http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
		isURL=re.test(userURL)
		if not isURL then
			errinfo=errinfo & "<li>个人主页地址填写有误"
		elseif len(userURL)>100 then
			errinfo=errinfo & "<li>过长的个人主页地址"
		end if
	end if

	If usermail<>"" then
		dim isEmail
		re.Pattern="[\w\[\]\@\(\)\.]+\.+[A-Za-z]{2,4}$"
		isEmail=re.test(usermail)
		if not isEmail then
			errinfo=errinfo & "<li>电子邮件地址填写有误"
		elseif len(usermail)>100 then
			errinfo=errinfo & "<li>过长的电子邮件地址"
		end if
	end if
	set re=Nothing

	if trim(userqq)<>"" then
		if not(isnumeric(userqq)) then
			errinfo=errinfo & "<li>QQ号码填写有误"
		elseif len(userqq)>10 then
			errinfo=errinfo & "<li>过长的QQ号码"
		end if
	end if

	if len(replycode)>45 then
		errinfo=errinfo & "<li>过长的回复查看码"
	end if

	if username=name then
		errinfo=errinfo & "<li>请勿使用管理员的名称"
	end if

	call error()

	if needcheck=0 or whisper=1 then
		checked=1
	else
		checked=0
	end if

	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from [topic]"
	rs.open sql,conn,3,2
	rs.addnew
	rs("username")=username
	rs("xingbie")=xingbie
	rs("userface")=userface
	rs("userURL")=userURL
	rs("usermail")=usermail
	rs("userqq")=userqq
	rs("usertime")=now()
	rs("usertitle")=usertitle
	rs("usercontent")=usercontent
	rs("whisper")=whisper
	rs("replycode")=replycode
	rs("top")="0"
	rs("reply")="0"
	rs("ip")=ip
	rs("checked")=checked
	rs.update
	rs.close

	if checked=0 then
		response.write"<script>alert('留言已成功提交，但需要通过审核后才会显示。');location='index.asp'</script>"
	else
		if whisper=1 then
			if replycode="" then
				response.write"<script>alert('悄悄话留言已成功提交，但您没有填写回复查看码，管理员将无法回复您的发言，您也无法查看回复。');location='index.asp'</script>"
			else
				response.write"<script>alert('悄悄话留言已成功提交，您可以在管理员回复后，通过输入回复查看码，查看回复内容。');location='index.asp'</script>"
			end if
		else
			Response.Redirect "index.asp"
			Response.Flush
		end if	
	end if
end sub
%>