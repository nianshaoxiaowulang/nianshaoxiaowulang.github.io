<%Session.CodePage=936%>
<!--#include file="inc_connect.asp"-->
<!--#include file="inc_css.asp"-->
<%
'**************************************
'**		inc_common.asp
'**
'** �ļ�˵�������ú���
'** �޸����ڣ�2005-04-07
'**************************************

'----------------����ͳ��--------------
if not session("in_site") = "true" then
	conn.execute("update admin set stat=stat+1 where id=1")
	session("in_site") = "true"
end if

'---------------����IP����-----------
dim ip
if request.servervariables("http_x_forwarded_for")="" then
	ip=request.servervariables("remote_addr")
else
	ip=request.servervariables("http_x_forwarded_for")
end if

if not badip="" then
	dim allbadip,i
		allbadip=split(badip,chr(13)&chr(10))
	for i = lbound(allbadip) to ubound(allbadip)
		if ip=trim(allbadip(i)) then
			errinfo="<li>����ip��ַ�޷��������Ա���"
			call error()
			response.end
		end if
	next
end if

'-----------------����������ȡ----------
dim name,password,perpage,site,URL,adminmail
dim maxnum,notice,stat,lock,needcheck,badip,adword,UBBcfg
	sql="select top 1 * from [admin]"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
		name=rs("name")
		password=rs("password")
		perpage=rs("perpage")
		site=rs("site")
		URL=rs("URL")
		adminmail=rs("adminmail")
		maxnum=rs("maxnum")
		notice=rs("notice")
		stat=rs("stat")
		lock=rs("lock")
		needcheck=rs("needcheck")
		badip=rs("badip")
		adword=rs("adword")
		UBBcfg=rs("UBBconfig")
		'�Զ���UBB֧��
		dim UBBcfg_font,UBBcfg_size,UBBcfg_color,UBBcfg_b,UBBcfg_i,UBBcfg_u
		dim UBBcfg_center,UBBcfg_URL,UBBcfg_email,UBBcfg_shadow,UBBcfg_glow
		dim UBBcfg_pic,UBBcfg_swf,UBBcfg_face
			if instr(UBBcfg,"font")>0 then UBBcfg_font=1
			if instr(UBBcfg,"size")>0 then UBBcfg_size=1
			if instr(UBBcfg,"color")>0 then UBBcfg_color=1
			if instr(UBBcfg,"bold")>0 then UBBcfg_b=1
			if instr(UBBcfg,"italic")>0 then UBBcfg_i=1
			if instr(UBBcfg,"underline")>0 then UBBcfg_u=1
			if instr(UBBcfg,"center")>0 then UBBcfg_center=1
			if instr(UBBcfg,"URL")>0 then UBBcfg_URL=1
			if instr(UBBcfg,"email")>0 then UBBcfg_email=1
			if instr(UBBcfg,"shadow")>0 then UBBcfg_shadow=1
			if instr(UBBcfg,"glow")>0 then UBBcfg_glow=1
			if instr(UBBcfg,"pic")>0 then UBBcfg_pic=1
			if instr(UBBcfg,"swf")>0 then UBBcfg_swf=1
			if instr(UBBcfg,"face")>0 then UBBcfg_face=1
	rs.close
	set rs=nothing

'---------------ҳ��ͷ����Ϣ-------------
sub pageinfo()
%>
<HTML>
<head>
<meta name="description" content="PHOTOSHOP��ѧ��<%=xm_version%>">
<meta name="keywords" content="<%=site%>,PHOTOSHOP�̳�,PHOTOSHOP�̲�,ͼ��ͼ��,���Ի滭,��ͼ�̳�,��Ӱ���ڹ���,ͼƬ����">
<meta http-equiv="content-language" content="zh-cn">
<meta http-equiv="content-type" content="text/HTML; charset=gb2312">
<%if pagename="�鿴����" then%>
<title>���Ա�--<%=site%></title>
<%else%>
<title><%=pagename%>__<%=site%></title>
<%end if%>
</head>

<script language="javascript">
function submitonce(theform){
	if (document.all||document.getelementbyid){
		for (i=0;i<theform.length;i++){
		var tempobj=theform.elements[i]
		if(tempobj.type.tolowercase()=="submit"||tempobj.type.tolowercase()=="reset")
		tempobj.disabled=true
		}
	}
}
</script>
<%
end sub

'--------------ͨ�ý���ǰ�벿��------------
sub skin1()
%>
<body topmargin="0" leftmargin="0">
<!--#include file="inc_top.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <tr>
		<td valign="top" width="100" bgcolor="#F1F1F1" style="border-right: 1px solid #969696">
		<!--#include file="inc_menu.asp"-->
		</td>
		<td valign="top">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
			<tr>
				<td width="400" bgcolor="<%=maincolor%>"><img src="images/<%=mainpic%>" width="400" height="46"></td>
				<td width="100%" bgcolor="<%=maincolor%>">&nbsp;</td>
				<td width="185" bgcolor="<%=maincolor%>"><img src="images/pagetitle2.gif" width="185" height="46"></td>
			</tr>
			<tr>
				<td width="100%" colspan="3"><%
end sub

'--------------ͨ�ý����벿��------------
sub skin2()
%>
				</td>
			</tr>
		</table>	
		</td>
	</tr>
	<tr>
		
    <td height="50" colspan="2" align="center" style="border-top: 1px solid #969696; filter: progid:dximagetransform.microsoft.gradient(startcolorstr='#FFFFFF', endcolorstr='<%=maincolor%>', gradienttype='1'"><a href="<%=URL%>"><%=site%></a>��Ȩ����<br>
      Copyright 2004-2005 All Right Reserved </td>
	</tr>
</table>
</body>
</HTML>
<%
end sub

'----------------������Ϣ--------------
dim errinfo
errinfo=""

sub error()
if not errinfo="" then
%>
<title>����</title>
<p align=center><img width="122" height="50" border="0" src="images/error.gif"></p>
<p>
<table cellpadding=6 cellspacing=1 align=center class=table1 width='550'>
	<tr>
		<td width='100%' class=tablebody3><b><font color="red">����</font></b></td>
	</tr>
	<tr>
		<td width='100%' class=tablebody1><%=errinfo%></td>
	</tr>
	<tr align=center>
		<td width='100%' class=tablebody3><a href="javascript:history.back(1)"><b>&lt;&lt; ����</b></a></td>
	</tr>
</table>
</p>
<%
response.end
end if
end sub

dim pagename,maincolor
	maincolor="#5581D2"	'���Ա�����ɫ��
dim xm_version
	xm_version="2005" '�汾��
dim mainpic

function sql_filter(text)	'-------���ύ����ʱ����SQL����-------
	if isnull(text) then
		sql_filter=""
		exit function
	end if
	text = LCase(text)
	text = Replace(text,"'","''")
	text = Replace(text,">","&gt;")
	text = Replace(text,"<","&lt;")
	text = Replace(text,";","��")
	text = Replace(text,"and","����")
	text = Replace(text,"exec","������")
	text = Replace(text,"execute","�����������")
	text = Replace(text,"insert","�������")
	text = Replace(text,"select","�������")
	text = Replace(text,"delete","��������")
	text = Replace(text,"update","���������")
	text = Replace(text,"count","�������")
	text = Replace(text,"*","��")
	text = Replace(text,"%","��")
	text = Replace(text,"chr","����")
	text = Replace(text,"mid","����")
	text = Replace(text,"master","��������")
	text = Replace(text,"truncate","������������")
	text = Replace(text,"char","�����")
	text = Replace(text,"declare","��������")
	sql_filter = text
end function

function back_filter(text)	'-------����ʾ����ʱ��ԭ���滻��SQL-------
	if isnull(text) then
		back_filter=""
		exit function
	end if
	text = Replace(text,"''","'")
	text = Replace(text,"��",";")
	text = Replace(text,"����","and")
	text = Replace(text,"������","exec")
	text = Replace(text,"�����������","execute")
	text = Replace(text,"�������","insert")
	text = Replace(text,"�������","select")
	text = Replace(text,"��������","delete")
	text = Replace(text,"���������","update")
	text = Replace(text,"�������","count")
	text = Replace(text,"��","*")
	text = Replace(text,"��","%")
	text = Replace(text,"����","chr")
	text = Replace(text,"����","mid")
	text = Replace(text,"��������","master")
	text = Replace(text,"������������","truncate")
	text = Replace(text,"�����","char")
	text = Replace(text,"��������","declare")
	back_filter = text
end function
%>