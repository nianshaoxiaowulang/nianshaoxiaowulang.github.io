<% option explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		admin.asp
'**
'** 文件说明：留言本管理页面
'** 修改日期：2005-04-07
'** 作者：Howlion
'** email：howlion@163.com
'**************************************

select case request.querystring("act")
	case "admchk"
	call admchk()
	case "main"
	call main()
	case "update"
	call update()
	case "batch"
	call batch()
	case "check"
	call check()
	case "del"
	call del()
	case "logout"
	call logout()
	case else
	call main()
end select

sub admchk()
	dim adminpass
		adminpass=trim(request.form("adminpass"))

	if adminpass = password then
		session("login")="true"
		response.redirect "index.asp"
	else
		errinfo="<li>密码错误！请注意大小写。"
		call error()
	end if
end sub

sub main()		'-----------------留言本设置页面

	if not session("login")="true" then
		errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
		call error()
	else
		pagename="设置留言本"
		call pageinfo()
		mainpic="page_admin.gif"
		call skin1()
'---------------以下显示页面主体--------
%>
<br>
<form method="post" action="?act=update">
	<div align="center">
		<center>
		<table border="0" cellpadding="3" cellspacing="1" width="95%" class="table1">
			<tr>
				<td width="28%" class="tablebody3" align="right">网站名称：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="site" size="29" value="<%=site%>"><br>
				留言本所属的站点的名称（如：搜狐网）。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">网站地址：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="URL" size="29" value="<%=URL%>"><br>
				请注意：必须是完整的地址（如：http://www.skyim.com/）。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">管理员名称：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="name" size="29" value="<%=name%>"><br>
				换上你自己的称呼。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">管理员密码：</td>
				<td width="72%" class="tablebody2">
				<input type="password" name="password" size="29" value="<%=password%>"><br>
				若不希望修改，别改动就行了。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">重复密码：</td>
				<td width="72%" class="tablebody2">
				<input type="password" name="password2" size="29" value="<%=password%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">管理员email：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="adminmail" size="29" value="<%=adminmail%>"><br>
				换上你自己的email。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">每页显示留言数：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="perpage" size="10" value="<%=perpage%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">最大留言字数：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="maxnum" size="10" value="<%=maxnum%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">留言提示：</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="notice" cols="55"><%=notice%></textarea><br>
				可以是欢迎词、警告、站点说明等，将出现在提交留言页面的顶部，支持UBB代码。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">不受欢迎的IP：</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="badip" cols="34"><%=badip%></textarea><br>
				不受欢迎的IP地址将无法进入留言本。<b>每个IP地址必须占一行</b>。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">广告过滤：</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="adword" cols="34"><%=adword%></textarea><br>
				包含上述某一词语的留言将无法提交，如果以上某个词语和您的站点的主题有关，请将其从文本框中删去。<br><b>每个词语必须占一行</b>。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">访问次数：</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="stat" size="10" value="<%=stat%>">&nbsp;修改留言本自带的计数器的数值。</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">自定义UBB支持：</td>
				<td width="72%" class="tablebody2">
				<input type="checkbox" name="UBBcfg" value="font"<%if UBBcfg_font=1 then%> checked<%end if%>>字体&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="size"<%if UBBcfg_size=1 then%> checked<%end if%>>字号&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="color"<%if UBBcfg_color=1 then%> checked<%end if%>>文字颜色&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="bold"<%if UBBcfg_b=1 then%> checked<%end if%>>粗体&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="italic"<%if UBBcfg_i=1 then%> checked<%end if%>>斜体&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="underline"<%if UBBcfg_u=1 then%> checked<%end if%>>下划线<br><input type="checkbox" name="UBBcfg" value="center"<%if UBBcfg_center=1 then%> checked<%end if%>>居中&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="URL"<%if UBBcfg_URL=1 then%> checked<%end if%>>超链接&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="email"<%if UBBcfg_email=1 then%> checked<%end if%>>email链接&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="shadow"<%if UBBcfg_shadow=1 then%> checked<%end if%>>阴影字&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="glow"<%if UBBcfg_glow=1 then%> checked<%end if%>>发光字<br><input type="checkbox" name="UBBcfg" value="pic"<%if UBBcfg_pic=1 then%> checked<%end if%>>图片&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="swf"<%if UBBcfg_swf=1 then%> checked<%end if%>>Flash&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="face"<%if UBBcfg_face=1 then%> checked<%end if%>>表情图<br><B>注意：</B>某些Flash动画可能包含有害的脚本，请慎重选择是否支持Flash！</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">留言本状态：</td>
				<td width="72%" class="tablebody2">
				<input type="radio" value="0" <%if lock=0 then%>checked <%end if%>name="lock">开放&nbsp;&nbsp;<input type="radio" value="1" <%if lock=1 then%> checked<%end if%> name="lock">锁定&nbsp; 
				&nbsp;（若锁定，任何人都不能发表留言）</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">留言审核：</td>
				<td width="72%" class="tablebody2">
				<input type="radio" value="1" <%if needcheck=1 then%>checked <%end if%>name="needcheck">需要&nbsp;&nbsp;<input type="radio" value="0" <%if needcheck=0 then%> checked<%end if%> name="needcheck">不需要&nbsp;&nbsp;（未经审核的留言将不会被显示，但管理员可以看到）</td>
			</tr>
		</table>
		<p><input type="submit" value="提 交" name="submit">&nbsp;&nbsp;
		<input type="reset" value="清 除" name="submit2"> </p>
		</center>
	</div>
</form>

<%
	end if
	conn.close
	set rs=nothing
'--------------页面主题显示结束--------
	call skin2()
end sub

sub update()		'-----------------更新设置
if not session("login")="true" then
	errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
	call error()
else
	dim site,URL,name,password,password2,adminmail,perpage,maxnum,notice,stat,lock,needcheck
	site=trim(request.form("site"))
	URL=trim(request.form("URL"))
	name=trim(request.form("name"))
	password=trim(request.form("password"))
	password2=trim(request.form("password2"))
	adminmail=trim(request.form("adminmail"))
	perpage=trim(request.form("perpage"))
	maxnum=trim(request.form("maxnum"))
	notice=request.form("notice")
	badip=trim(request.form("badip"))
	adword=trim(request.form("adword"))
	stat=trim(request.form("stat"))
	UBBcfg=request.form("UBBcfg")
	lock=request.form("lock")
	needcheck=request.form("needcheck")

	if site="" or URL="" or name="" or password="" or adminmail="" or perpage="" or maxnum="" or stat="" or lock="" then
		errinfo=errinfo & "<li>内容填写不完整。除了留言提示、IP列表和广告过滤以外，其它各项都是必填的"
	end if

	if password<>password2 then
		errinfo=errinfo & "<li>两次输入的密码不一致"
	end if

	if (not perpage="") and not(isnumeric(perpage)) then
		errinfo=errinfo & "<li>每页显示留言数必须为数字"
	end if

	if (not maxnum="") and not(isnumeric(maxnum)) then
		errinfo=errinfo & "<li>最大留言字数必须为数字"
	end if

	if (not stat="") and not(isnumeric(stat)) then
		errinfo=errinfo & "<li>访问次数必须为数字"
	end if

	call error()

	set rs= server.createobject("adodb.recordset")
		sql="select * from [admin] where id=1"
		rs.open sql,conn,3,2
		rs.update
		rs("site")=site
		rs("URL")=URL
		rs("name")=name
		rs("password")=password
		rs("adminmail")=adminmail
		rs("perpage")=perpage
		rs("maxnum")=maxnum
		rs("notice")=notice
		rs("badip")=badip
		rs("adword")=adword
		rs("stat")=stat
		rs("ubbconfig")=UBBcfg
		rs("lock")=lock
		rs("needcheck")=needcheck
		rs.update
		rs.close

	response.redirect "index.asp"
	response.flush

end if
end sub

sub batch()		'-----------------批量管理留言

	dim currentpage,page_count,pcount
	dim totalrec,endpage
	if request.querystring("page")="" then
		currentpage=1
	else
		currentpage=cint(request.querystring("page"))
	end if

	if not session("login")="true" then
		errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
		call error()
	end if

	if (not isnumeric(request.querystring("page"))) or (not isnumeric(request.querystring("page_num"))) then
		errinfo="<li>非法的页面参数！"
		call error()
	end if

	pagename="批量管理留言"
	call pageinfo()
	mainpic="page_admin_lw.gif"
	call skin1()
	'---------------以下显示页面主体--------
%>
<script language="JavaScript" type="text/JavaScript">
<!--

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
}

function MM_jumpMenu(targ,selObj,restore){
	eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
	if (restore) selObj.selectedIndex=0;
}

function SetSubmitType(sub_type){
	if (confirm("确定要执行批量操作吗？")){
	SetSubmitType = sub_type;
	}
}

function Submit_all(theForm){
	var flag = false;
		if ( SetSubmitType == 'del'){
			flag = true;
			theForm.action = theForm.action + "del";
		}
		else if (SetSubmitType == 'check'){
			flag = true;
			theForm.action = theForm.action + "check";
		}
	return flag;
}

function go(src,q){
	var ret;
	ret = confirm(q);
	if(ret!=false)window.location=src;
}

function openwin(URL, width, height){
	var win = window.open(URL,"openscript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=1,menubar=0,status=1');
}
//-->
</script>
<%
	dim view,page_num
		view=request.querystring("view")
		if request.querystring("page_num")="" then
			page_num=20
		else
			page_num=request.querystring("page_num")
		end if

	select case request.querystring("view")
		case "1"
			sql="select * from [topic] where (not checked=1) order by usertime desc"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
		case "2"
			sql="select * from [topic] where (not reply=1) order by usertime desc"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
		case else
			sql="select * from [topic] order by usertime desc"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
	end select

	if rs.eof and rs.bof then
	%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" height="390">
		<tr>
			<td valign="middle" align="center"><font size="4">没有此类留言。</font></td>
		</tr>
	</table>
	<%
		set rs=nothing
		call skin2()
		response.end
	end if

		rs.pagesize = page_num
		rs.absolutepage=currentpage
		page_count=0
		totalrec=rs.recordcount

%>
<form name="batch_form" action="admin.asp?act=" method="post" onsubmit="return Submit_all(batch_form);">
<div align="center">
	<table border="0" cellpadding="3" cellspacing="1" width="95%">
		<tr>
			<td width="100%" align="right">
			显示：<select name="menu_view" onChange="MM_jumpMenu('parent',this,0)">
			<option value="?act=batch&page_num=<%=page_num%>"<%if view="" then%> selected<%end if%>>所有留言</option>
			<option value="?act=batch&view=1&page_num=<%=page_num%>"<%if view="1" then%> selected<%end if%>>所有未通过审核的留言</option>
			<option value="?act=batch&view=2&page_num=<%=page_num%>"<%if view="2" then%> selected<%end if%>>所有未被回复的留言</option>
			</select>&nbsp;每页显示：<select name="menu_page_num" onChange="MM_jumpMenu('parent',this,0)">
			<option value="?act=batch&view=<%=view%>&page_num=10"<%if page_num="10" then%> selected<%end if%>>10条</option>
			<option value="?act=batch&view=<%=view%>&page_num=20"<%if page_num="20" then%> selected<%end if%>>20条</option>
			<option value="?act=batch&view=<%=view%>&page_num=50"<%if page_num="50" then%> selected<%end if%>>50条</option>
			<option value="?act=batch&view=<%=view%>&page_num=100"<%if page_num="100" then%> selected<%end if%>>100条</option>
			</select></td>
		</tr>
	</table>
	<table border="0" cellpadding="5" cellspacing="1" width="95%">
		<tr>
			<td width="6%" align="center">
			<input type="hidden" value="<%=Request.ServerVariables("HTTP_URL")%>" name="thisURL">
			<input type="checkbox" value="on" name="chkall" onclick="CheckAll(this.form)"><br>全选
			</td>
			<td width="30%">
			<input type="submit" name="Submit_Del" value="批量删除" onclick="SetSubmitType('del');">
			<input type="submit" name="Submit_Check" value="批量审核" onclick="SetSubmitType('check');">
			</td>
			<td width="64%">
			<%call pages(currentpage,page_count,pcount,totalrec,endpage,page_num,view)%>
			</td>
		</tr>
	</table>
	<table border="0" cellpadding="5" cellspacing="1" width="95%" class="table1">
		<%while (not rs.eof) and (not page_count = rs.pagesize)%>
		<tr>
			<td width="6%" class="tablebody3" align="center" rowspan="2"><input type="checkbox" name="id" value="<%=rs("id")%>"></td>
			<td width="94%" class="tablebody3">
			<B>标题</B>：<%=batchEncode(rs("usertitle"))%><br><B>时间</B>：<font face="Verdana" SIZE="1"><%=rs("usertime")%></font>
			</td>
		</tr>
		<tr>
			<td class="tablebody3">
			<B>内容代码</B>：<%=batchEncode(rs("usercontent"))%>
			</td>
		</tr>
		<tr>
			<td class="tablebody1" colspan="2" align="right">
			<%if rs("checked")=0 then%><a href="javascript:go('admin.asp?act=check&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','您确定要通过审核？')"><font COLOR="red"><b>通过审核</b></font></a>&nbsp;&nbsp;<%end if%><a href="javascript:go('admin.asp?act=del&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','您确定要删除？')">删除</a>&nbsp;&nbsp;<%if rs("whisper")=1 and rs("replycode")="" then%><font COLOR="red"><b>无法回复的悄悄话</b></font><%else%><a href="JavaScript:openwin('reply.asp?id=<%=rs("id")%>',600,500)"><%if rs("whisper")=1 then%><font COLOR="red"><b>悄悄话回复/编辑回复</b></font><%else%>回复/编辑回复<%end if%></a><%end if%>&nbsp;&nbsp;<a href="JavaScript:openwin('edit.asp?id=<%=rs("id")%>',600,500)">编辑</a>&nbsp;&nbsp;留言IP：<%=rs("ip")%>
			</td>
		</tr>
		<%page_count = page_count + 1
		rs.movenext
		wend%>
	</table>
	<table border="0" cellpadding="5" cellspacing="1" width="95%">
		<tr>
			<td width="6%" align="center">
			<input type="checkbox" name="chkall2" onclick="javascript:chkall.click()"><br>全选
			</td>
			<td width="30%">
			<input type="submit" name="Submit_Del2" value="批量删除" onclick="SetSubmitType('del');">
			<input type="submit" name="Submit_Check2" value="批量审核" onclick="SetSubmitType('check');">
			</td>
			<td width="64%">
			<%call pages(currentpage,page_count,pcount,totalrec,endpage,page_num,view)%>
			</td>
		</tr>
	</table>
</div>
</form>
<%

rs.close
set rs=nothing
'--------------页面主题显示结束--------
call skin2()
end sub

sub pages(currentpage,page_count,pcount,totalrec,endpage,page_num,view)	'--分页代码--
dim ii,p,n
if totalrec mod page_num=0 then
	n= totalrec \ page_num
else
	n= totalrec \ page_num+1
end if
p=(currentpage-1) \ 10
response.write "<table border=0 cellpadding=0 cellspacing=3 width='97%' align=center>"&_
"<tr>"&_
"<td valign=middle align=right>页次：<b>"& currentpage &"/"& n &"</b>页，共<b>"& totalrec &"</b>条&nbsp;&nbsp;"

if currentpage=1 then
	response.write "<font face=webdings>9</font>	 "
else
	response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page=1' title=首页><font face=webdings>9</font></a>	 "
end if
if p*10>0 then response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>	 "
response.write "<b>"
for ii=p*10+1 to p*10+10
	if ii=currentpage then
		response.write "<font size=4>"+cstr(ii)+"</font> "
	else
		response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(ii)&"'>"+cstr(ii)+"</a>	 "
	end if
	if ii=n then exit for
	'p=p+1
next
response.write "</b>"
if ii<n then response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(ii)&"' title=下十页><font face=webdings>8</font></a>	 "
if currentpage=n then
	response.write "<font face=webdings>:</font>	 "
else
	response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(n)&"' title=尾页><font face=webdings>:</font></a>	 "
end if
response.write "</table>"
end sub

function batchEncode(fString)
	if not isnull(fString) then
		fString = back_filter(fString)
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, "''", "'")
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(36), "&#36;")
		batchEncode = fString
	end if
end function

sub del()		'-----------------删除留言
	dim id
	id = request("id")

	if not session("login")="true" then
		errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
		call error()
	else
		if id="" then
			errinfo="<li>您未选定任何留言。"
			call error()
		end if

		sql="select id from [topic] where id in ("&id&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3

		if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		errinfo="<li>该留言不存在。"
		call error()
		end if

		rs.close
		set rs=nothing

		sql="delete from [topic] where id in ("&id&")"
		conn.execute(sql)

		if request.querystring("page")="" then
			response.redirect request("thisURL")
		else
			response.redirect request("thisURL") & "&page=" & request.querystring("page")
		end if
		response.flush

	end if
end sub

sub check()		'-----------------审核留言
	dim id
	id = request("id")

	if not session("login")="true" then
		errinfo="<li>您未登陆或已经退出登陆，不能进入该页。"
		call error()
	else
		if id="" then
			errinfo="<li>您未选定任何留言。"
			call error()
		end if

		sql="select id from [topic] where id in ("&id&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			errinfo="<li>该留言不存在。"
			call error()
		else
			rs.close
			set rs=nothing
		end if

		sql="update topic set checked='1' where id in ("&id&")"
		conn.execute(sql)

		if request.querystring("page")="" then
			response.redirect request("thisURL")
		else
			response.redirect request("thisURL") & "&page=" & request.querystring("page")
		end if
		response.flush

	end if
end sub

sub logout()
	session("login")="false"
	response.redirect "index.asp"
end sub
%>