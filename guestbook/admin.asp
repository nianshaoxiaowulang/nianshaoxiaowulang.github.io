<% option explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		admin.asp
'**
'** �ļ�˵�������Ա�����ҳ��
'** �޸����ڣ�2005-04-07
'** ���ߣ�Howlion
'** email��howlion@163.com
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
		errinfo="<li>���������ע���Сд��"
		call error()
	end if
end sub

sub main()		'-----------------���Ա�����ҳ��

	if not session("login")="true" then
		errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
		call error()
	else
		pagename="�������Ա�"
		call pageinfo()
		mainpic="page_admin.gif"
		call skin1()
'---------------������ʾҳ������--------
%>
<br>
<form method="post" action="?act=update">
	<div align="center">
		<center>
		<table border="0" cellpadding="3" cellspacing="1" width="95%" class="table1">
			<tr>
				<td width="28%" class="tablebody3" align="right">��վ���ƣ�</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="site" size="29" value="<%=site%>"><br>
				���Ա�������վ������ƣ��磺�Ѻ�������</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">��վ��ַ��</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="URL" size="29" value="<%=URL%>"><br>
				��ע�⣺�����������ĵ�ַ���磺http://www.skyim.com/����</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">����Ա���ƣ�</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="name" size="29" value="<%=name%>"><br>
				�������Լ��ĳƺ���</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">����Ա���룺</td>
				<td width="72%" class="tablebody2">
				<input type="password" name="password" size="29" value="<%=password%>"><br>
				����ϣ���޸ģ���Ķ������ˡ�</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">�ظ����룺</td>
				<td width="72%" class="tablebody2">
				<input type="password" name="password2" size="29" value="<%=password%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">����Աemail��</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="adminmail" size="29" value="<%=adminmail%>"><br>
				�������Լ���email��</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">ÿҳ��ʾ��������</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="perpage" size="10" value="<%=perpage%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">�������������</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="maxnum" size="10" value="<%=maxnum%>"></td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">������ʾ��</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="notice" cols="55"><%=notice%></textarea><br>
				�����ǻ�ӭ�ʡ����桢վ��˵���ȣ����������ύ����ҳ��Ķ�����֧��UBB���롣</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">���ܻ�ӭ��IP��</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="badip" cols="34"><%=badip%></textarea><br>
				���ܻ�ӭ��IP��ַ���޷��������Ա���<b>ÿ��IP��ַ����ռһ��</b>��</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">�����ˣ�</td>
				<td width="72%" class="tablebody2">
				<textarea rows="5" name="adword" cols="34"><%=adword%></textarea><br>
				��������ĳһ��������Խ��޷��ύ���������ĳ�����������վ��������йأ��뽫����ı�����ɾȥ��<br><b>ÿ���������ռһ��</b>��</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">���ʴ�����</td>
				<td width="72%" class="tablebody2">
				<input type="text" name="stat" size="10" value="<%=stat%>">&nbsp;�޸����Ա��Դ��ļ���������ֵ��</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">�Զ���UBB֧�֣�</td>
				<td width="72%" class="tablebody2">
				<input type="checkbox" name="UBBcfg" value="font"<%if UBBcfg_font=1 then%> checked<%end if%>>����&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="size"<%if UBBcfg_size=1 then%> checked<%end if%>>�ֺ�&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="color"<%if UBBcfg_color=1 then%> checked<%end if%>>������ɫ&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="bold"<%if UBBcfg_b=1 then%> checked<%end if%>>����&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="italic"<%if UBBcfg_i=1 then%> checked<%end if%>>б��&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="underline"<%if UBBcfg_u=1 then%> checked<%end if%>>�»���<br><input type="checkbox" name="UBBcfg" value="center"<%if UBBcfg_center=1 then%> checked<%end if%>>����&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="URL"<%if UBBcfg_URL=1 then%> checked<%end if%>>������&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="email"<%if UBBcfg_email=1 then%> checked<%end if%>>email����&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="shadow"<%if UBBcfg_shadow=1 then%> checked<%end if%>>��Ӱ��&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="glow"<%if UBBcfg_glow=1 then%> checked<%end if%>>������<br><input type="checkbox" name="UBBcfg" value="pic"<%if UBBcfg_pic=1 then%> checked<%end if%>>ͼƬ&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="swf"<%if UBBcfg_swf=1 then%> checked<%end if%>>Flash&nbsp;&nbsp;<input type="checkbox" name="UBBcfg" value="face"<%if UBBcfg_face=1 then%> checked<%end if%>>����ͼ<br><B>ע�⣺</B>ĳЩFlash�������ܰ����к��Ľű���������ѡ���Ƿ�֧��Flash��</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">���Ա�״̬��</td>
				<td width="72%" class="tablebody2">
				<input type="radio" value="0" <%if lock=0 then%>checked <%end if%>name="lock">����&nbsp;&nbsp;<input type="radio" value="1" <%if lock=1 then%> checked<%end if%> name="lock">����&nbsp; 
				&nbsp;�����������κ��˶����ܷ������ԣ�</td>
			</tr>
			<tr>
				<td width="28%" class="tablebody3" align="right">������ˣ�</td>
				<td width="72%" class="tablebody2">
				<input type="radio" value="1" <%if needcheck=1 then%>checked <%end if%>name="needcheck">��Ҫ&nbsp;&nbsp;<input type="radio" value="0" <%if needcheck=0 then%> checked<%end if%> name="needcheck">����Ҫ&nbsp;&nbsp;��δ����˵����Խ����ᱻ��ʾ��������Ա���Կ�����</td>
			</tr>
		</table>
		<p><input type="submit" value="�� ��" name="submit">&nbsp;&nbsp;
		<input type="reset" value="�� ��" name="submit2"> </p>
		</center>
	</div>
</form>

<%
	end if
	conn.close
	set rs=nothing
'--------------ҳ��������ʾ����--------
	call skin2()
end sub

sub update()		'-----------------��������
if not session("login")="true" then
	errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
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
		errinfo=errinfo & "<li>������д������������������ʾ��IP�б�͹��������⣬��������Ǳ����"
	end if

	if password<>password2 then
		errinfo=errinfo & "<li>������������벻һ��"
	end if

	if (not perpage="") and not(isnumeric(perpage)) then
		errinfo=errinfo & "<li>ÿҳ��ʾ����������Ϊ����"
	end if

	if (not maxnum="") and not(isnumeric(maxnum)) then
		errinfo=errinfo & "<li>���������������Ϊ����"
	end if

	if (not stat="") and not(isnumeric(stat)) then
		errinfo=errinfo & "<li>���ʴ�������Ϊ����"
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

sub batch()		'-----------------������������

	dim currentpage,page_count,pcount
	dim totalrec,endpage
	if request.querystring("page")="" then
		currentpage=1
	else
		currentpage=cint(request.querystring("page"))
	end if

	if not session("login")="true" then
		errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
		call error()
	end if

	if (not isnumeric(request.querystring("page"))) or (not isnumeric(request.querystring("page_num"))) then
		errinfo="<li>�Ƿ���ҳ�������"
		call error()
	end if

	pagename="������������"
	call pageinfo()
	mainpic="page_admin_lw.gif"
	call skin1()
	'---------------������ʾҳ������--------
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
	if (confirm("ȷ��Ҫִ������������")){
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
			<td valign="middle" align="center"><font size="4">û�д������ԡ�</font></td>
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
			��ʾ��<select name="menu_view" onChange="MM_jumpMenu('parent',this,0)">
			<option value="?act=batch&page_num=<%=page_num%>"<%if view="" then%> selected<%end if%>>��������</option>
			<option value="?act=batch&view=1&page_num=<%=page_num%>"<%if view="1" then%> selected<%end if%>>����δͨ����˵�����</option>
			<option value="?act=batch&view=2&page_num=<%=page_num%>"<%if view="2" then%> selected<%end if%>>����δ���ظ�������</option>
			</select>&nbsp;ÿҳ��ʾ��<select name="menu_page_num" onChange="MM_jumpMenu('parent',this,0)">
			<option value="?act=batch&view=<%=view%>&page_num=10"<%if page_num="10" then%> selected<%end if%>>10��</option>
			<option value="?act=batch&view=<%=view%>&page_num=20"<%if page_num="20" then%> selected<%end if%>>20��</option>
			<option value="?act=batch&view=<%=view%>&page_num=50"<%if page_num="50" then%> selected<%end if%>>50��</option>
			<option value="?act=batch&view=<%=view%>&page_num=100"<%if page_num="100" then%> selected<%end if%>>100��</option>
			</select></td>
		</tr>
	</table>
	<table border="0" cellpadding="5" cellspacing="1" width="95%">
		<tr>
			<td width="6%" align="center">
			<input type="hidden" value="<%=Request.ServerVariables("HTTP_URL")%>" name="thisURL">
			<input type="checkbox" value="on" name="chkall" onclick="CheckAll(this.form)"><br>ȫѡ
			</td>
			<td width="30%">
			<input type="submit" name="Submit_Del" value="����ɾ��" onclick="SetSubmitType('del');">
			<input type="submit" name="Submit_Check" value="�������" onclick="SetSubmitType('check');">
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
			<B>����</B>��<%=batchEncode(rs("usertitle"))%><br><B>ʱ��</B>��<font face="Verdana" SIZE="1"><%=rs("usertime")%></font>
			</td>
		</tr>
		<tr>
			<td class="tablebody3">
			<B>���ݴ���</B>��<%=batchEncode(rs("usercontent"))%>
			</td>
		</tr>
		<tr>
			<td class="tablebody1" colspan="2" align="right">
			<%if rs("checked")=0 then%><a href="javascript:go('admin.asp?act=check&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','��ȷ��Ҫͨ����ˣ�')"><font COLOR="red"><b>ͨ�����</b></font></a>&nbsp;&nbsp;<%end if%><a href="javascript:go('admin.asp?act=del&id=<%=rs("id")%>&thisURL=<%=Request.ServerVariables("HTTP_URL")%>','��ȷ��Ҫɾ����')">ɾ��</a>&nbsp;&nbsp;<%if rs("whisper")=1 and rs("replycode")="" then%><font COLOR="red"><b>�޷��ظ������Ļ�</b></font><%else%><a href="JavaScript:openwin('reply.asp?id=<%=rs("id")%>',600,500)"><%if rs("whisper")=1 then%><font COLOR="red"><b>���Ļ��ظ�/�༭�ظ�</b></font><%else%>�ظ�/�༭�ظ�<%end if%></a><%end if%>&nbsp;&nbsp;<a href="JavaScript:openwin('edit.asp?id=<%=rs("id")%>',600,500)">�༭</a>&nbsp;&nbsp;����IP��<%=rs("ip")%>
			</td>
		</tr>
		<%page_count = page_count + 1
		rs.movenext
		wend%>
	</table>
	<table border="0" cellpadding="5" cellspacing="1" width="95%">
		<tr>
			<td width="6%" align="center">
			<input type="checkbox" name="chkall2" onclick="javascript:chkall.click()"><br>ȫѡ
			</td>
			<td width="30%">
			<input type="submit" name="Submit_Del2" value="����ɾ��" onclick="SetSubmitType('del');">
			<input type="submit" name="Submit_Check2" value="�������" onclick="SetSubmitType('check');">
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
'--------------ҳ��������ʾ����--------
call skin2()
end sub

sub pages(currentpage,page_count,pcount,totalrec,endpage,page_num,view)	'--��ҳ����--
dim ii,p,n
if totalrec mod page_num=0 then
	n= totalrec \ page_num
else
	n= totalrec \ page_num+1
end if
p=(currentpage-1) \ 10
response.write "<table border=0 cellpadding=0 cellspacing=3 width='97%' align=center>"&_
"<tr>"&_
"<td valign=middle align=right>ҳ�Σ�<b>"& currentpage &"/"& n &"</b>ҳ����<b>"& totalrec &"</b>��&nbsp;&nbsp;"

if currentpage=1 then
	response.write "<font face=webdings>9</font>	 "
else
	response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page=1' title=��ҳ><font face=webdings>9</font></a>	 "
end if
if p*10>0 then response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(p*10)&"' title=��ʮҳ><font face=webdings>7</font></a>	 "
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
if ii<n then response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(ii)&"' title=��ʮҳ><font face=webdings>8</font></a>	 "
if currentpage=n then
	response.write "<font face=webdings>:</font>	 "
else
	response.write "<a href='?act=batch&view="& view &"&page_num="& page_num &"&page="&cstr(n)&"' title=βҳ><font face=webdings>:</font></a>	 "
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

sub del()		'-----------------ɾ������
	dim id
	id = request("id")

	if not session("login")="true" then
		errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
		call error()
	else
		if id="" then
			errinfo="<li>��δѡ���κ����ԡ�"
			call error()
		end if

		sql="select id from [topic] where id in ("&id&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3

		if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		errinfo="<li>�����Բ����ڡ�"
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

sub check()		'-----------------�������
	dim id
	id = request("id")

	if not session("login")="true" then
		errinfo="<li>��δ��½���Ѿ��˳���½�����ܽ����ҳ��"
		call error()
	else
		if id="" then
			errinfo="<li>��δѡ���κ����ԡ�"
			call error()
		end if

		sql="select id from [topic] where id in ("&id&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			errinfo="<li>�����Բ����ڡ�"
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