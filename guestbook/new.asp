<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<%
'**************************************
'**		new.asp
'**
'** �ļ�˵������������ҳ��
'** �޸����ڣ�2005-04-07
'**************************************

if lock="1" then
	errinfo="<li>���Ա��Ѿ�������Ա���������޷����ԡ�"
	call error()
end if

select case Request.QueryString("act")
	case "addnew"
	call addnew()
	case else
	call main()
end select

sub main()

pagename="д����"
call pageinfo()
mainpic="page_new.gif"
call skin1()
'---------------������ʾҳ������--------
%>
<script language="JavaScript">
<!--
function Submitcheck(){
	if (document.lw_form.username.value.length==0){
	alert("���������ĳƺ�����Ϊ�����");
	document.lw_form.username.focus();
	return false;
}
	if (document.lw_form.usercontent.value.length==0){
	alert("�������������ģ���Ϊ�����");
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
				<td width="80%" class="tablebody2"><b>������Ϣ</b></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">
				<font FACE="����" COLOR="red">***</font> <b>���ĳƺ���</b></td>
				<td width="80%" class="tablebody2">
				<input name="username" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" valign="top" align="right">
				�Ա�ͷ��</td>
				<td width="80%" class="tablebody2">���ԣ�<img src="images/userfaces/small_1.gif" width="30" height="30"><input type="radio" value="1" name="userface" checked>&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_2.gif" width="30" height="30"><input type="radio" value="2" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_3.gif" width="30" height="30"><input type="radio" value="3" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_4.gif" width="30" height="30"><input type="radio" value="4" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_5.gif" width="30" height="30"><input type="radio" value="5" name="userface"><br>
				Ů�ԣ�<img src="images/userfaces/small_6.gif" width="30" height="30"><input type="radio" value="6" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_7.gif" width="30" height="30"><input type="radio" value="7" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_8.gif" width="30" height="30"><input type="radio" value="8" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_9.gif" width="30" height="30"><input type="radio" value="9" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_10.gif" width="30" height="30"><input type="radio" value="10" name="userface"><br>
				���ԣ�<img src="images/userfaces/small_11.gif" width="30" height="30"><input type="radio" value="11" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_12.gif" width="30" height="30"><input type="radio" value="12" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_13.gif" width="30" height="30"><input type="radio" value="13" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_14.gif" width="30" height="30"><input type="radio" value="14" name="userface">&nbsp;&nbsp;&nbsp;
				<img src="images/userfaces/small_15.gif" width="30" height="30"><input type="radio" value="15" name="userface"><p>
				<b>��ʾ��</b>�������͸¶�Լ����Ա𣬿���ѡһ�����Ե�ͷ��</b></font></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">������ҳ��</td>
				<td width="80%" class="tablebody2">
				<input name="userURL" size="19" maxlength="80" value="http://"><br>
				���Ҫ��д��Ӧ��д������ַ���磺http://www.howlion.com/</td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">Email��</td>
				<td width="80%" class="tablebody2">
				<input name="usermail" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">QQ��</td>
				<td width="80%" class="tablebody2">
				<input name="userqq" size="19" maxlength="80"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">
				<img border="0" src="images/lwcontent.gif"></td>
				<td width="80%" class="tablebody2"><b>��������</b></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" align="right">���⣺</td>
				<td width="80%" class="tablebody2">
				<input name="usertitle" size="40" maxlength="100"></td>
			</tr>
			<tr>
				<td width="20%" class="tablebody3" valign="top" align="right">
				<font FACE="����" COLOR="red">***</font> <b>���ģ�</b></td>
				<td width="80%" class="tablebody2">
				<!--#include file="inc_UBB.asp"-->
				<textarea cols="70" name="usercontent" title="Ctrl+Enter�ύ" rows="12" onkeydown="ctlent()"></textarea><br>�������ݲ��ܴ���<b><font size="3"><%=maxnum%></font></b>���ַ�<%if UBBcfg_face=1 then%><br>
				���������ſ��Խ���������ģ�<br>
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
				<td width="80%" class="tablebody2"><input type="checkbox" name="whisper" value="1">���Ļ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ظ��鿴�루�ɲ����<input name="replycode" size="20" maxlength="100"><p>
				<b>��ʾ��</b>�ظ��鿴�������Ժ�鿴����Ա�����Ļ��Ļظ����ɲ��������Ա���޷��ظ��������Ļ���</td>
			</tr>
			<tr>
				<td valign="middle" colspan="2" align="center" class="tablebody1">
				<input type="hidden" name="UBB_super" value="0">
				<input type="Submit" name="Submit" value="�� ��">&nbsp;&nbsp;
				<input type="reset" name="Submit2" value="�� ��">&nbsp;&nbsp;
				<input type="button" name="Preview" value="Ԥ��" onclick="openpreview()">
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
'--------------ҳ��������ʾ����--------
call skin2()
end sub


sub addnew()	'�����Լ������ݿ�
	dim servername1,servername2
		servername1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		servername2=Cstr(Request.ServerVariables("SERVER_NAME"))

		if mid(servername1,8,len(servername2))<>servername2 then
			errinfo="<li>�Ƿ����ύ������"
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
		errinfo=errinfo & "<li>δ��д���ĳƺ�"
		elseif len(username)>20 then
		errinfo=errinfo & "<li>�����ĳƺ�"
	end if

	if len(usertitle)>50 then
		errinfo=errinfo & "<li>�����ı���"
	end if

	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\[(.[^\]]*)\])"
	if	Trim(re.Replace(Replace(request.form("usercontent"), CHR(13)&CHR(10), ""),""))="" then
		errinfo=errinfo & "<li>δ��д��������"
		elseif len(usercontent)>maxnum then
		errinfo=errinfo & "<li>��������������"
	end if

	if not adword="" then
		dim alladword,i
			alladword=split(adword,chr(13)&chr(10))
		for i = lbound(alladword) to ubound(alladword)
			if instr(UCase(usercontent & usertitle),UCase(trim(alladword(i))))>0 then
				errinfo="<li>�����������"	'ע���˴�Ϊ�Ի󷢹�����
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
			errinfo=errinfo & "<li>������ҳ��ַ��д����"
		elseif len(userURL)>100 then
			errinfo=errinfo & "<li>�����ĸ�����ҳ��ַ"
		end if
	end if

	If usermail<>"" then
		dim isEmail
		re.Pattern="[\w\[\]\@\(\)\.]+\.+[A-Za-z]{2,4}$"
		isEmail=re.test(usermail)
		if not isEmail then
			errinfo=errinfo & "<li>�����ʼ���ַ��д����"
		elseif len(usermail)>100 then
			errinfo=errinfo & "<li>�����ĵ����ʼ���ַ"
		end if
	end if
	set re=Nothing

	if trim(userqq)<>"" then
		if not(isnumeric(userqq)) then
			errinfo=errinfo & "<li>QQ������д����"
		elseif len(userqq)>10 then
			errinfo=errinfo & "<li>������QQ����"
		end if
	end if

	if len(replycode)>45 then
		errinfo=errinfo & "<li>�����Ļظ��鿴��"
	end if

	if username=name then
		errinfo=errinfo & "<li>����ʹ�ù���Ա������"
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
		response.write"<script>alert('�����ѳɹ��ύ������Ҫͨ����˺�Ż���ʾ��');location='index.asp'</script>"
	else
		if whisper=1 then
			if replycode="" then
				response.write"<script>alert('���Ļ������ѳɹ��ύ������û����д�ظ��鿴�룬����Ա���޷��ظ����ķ��ԣ���Ҳ�޷��鿴�ظ���');location='index.asp'</script>"
			else
				response.write"<script>alert('���Ļ������ѳɹ��ύ���������ڹ���Ա�ظ���ͨ������ظ��鿴�룬�鿴�ظ����ݡ�');location='index.asp'</script>"
			end if
		else
			Response.Redirect "index.asp"
			Response.Flush
		end if	
	end if
end sub
%>