<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060200") then Call ReturnError1()
Dim TempSysRootDir
if SysRootDir = "" then
	TempSysRootDir = ""
else
	TempSysRootDir = "/" & SysRootDir
end if

dim JSID,JSObj,TempManner,TempDateStr
if request("JSID")<>"" then
	JSID = Clng(request("JSID"))
Set JSObj = Conn.Execute("select * from FS_FreeJS where ID = "&JSID&"")
if JSObj.eof then
	 Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	 response.end
end if
TempManner = JSObj("Manner")
TempDateStr = JSObj("ShowTimeTF")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����JS�޸�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body leftmargin="2" topmargin="2" >
<form action="" method="post" name="JSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.JSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="action" type="hidden" id="action3" value="mod"> 
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E7E7E7">
    <tr bgcolor="#FFFFFF" Title="JS���������ƣ����ں�̨���ĺ͹����벻Ҫ����25���ַ���"> 
      <td width="10%"> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="3"> 
        <input name="CName" type="text" id="CName" value="<%=JSObj("CName")%>" style="width:100%"> 
        <div align="center"></div></td>
      <td rowspan="12" align="center" id="PreviewArea"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">Ӣ������</div></td>
      <td colspan="3"> 
        <input name="EName" type="text" id="EName" value="<%=JSObj("EName")%>" disabled style="width:100%" Title="JS��Ӣ�����ƣ�����ǰ̨���ã��벻Ҫ����50���ַ��Ҳ������Ѿ����ڵ�JS������" > 
        <div align="center"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td width="20%"> 
        <input id="TypeW" name="Type" type="radio" value="0" <%if JSObj("Type")="0" then response.write("checked") end if%> onclick="TypeChoose();ChoosePic();" Title="JS���ͣ����֣�ѡ��" >
        ���� 
        <input id="TypeP" type="radio" name="Type" value="1" <%if JSObj("Type")="1" then response.write("checked") end if%> onclick="TypeChoose();ChoosePic();" Title="JS���ͣ�ͼƬ��ѡ��" >
        ͼƬ</td>
      <td width="10%"> 
        <div align="center">��������</div></td>
      <td width="20%"> 
        <input name="NewsNum" type="text" id="NewsNum2" value="<%=JSObj("NewsNum")%>" Title="��JS������õ�����������"   style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <select name="Manner" id="MannerW" style="width:100% " Title="����JS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic();">
          <option value="1" <%if JSObj("Manner")="1" then response.write("selected") end if%>>��ʽA</option>
          <option value="2" <%if JSObj("Manner")="2" then response.write("selected") end if%>>��ʽB</option>
          <option value="3" <%if JSObj("Manner")="3" then response.write("selected") end if%>>��ʽC</option>
          <option value="4" <%if JSObj("Manner")="4" then response.write("selected") end if%>>��ʽD</option>
          <option value="5" <%if JSObj("Manner")="5" then response.write("selected") end if%>>��ʽE</option>
        </select> </td>
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="RowNum" type="text" id="RowNum3" Title="��������JS��ÿ������ʾ����������������ز�Ҫ��Ϊ��0����"  value="<%=JSObj("RowNum")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ��ʽ</div></td>
      <td> 
        <select name="MannerP" id="MannerP" style="width:100% " disabled Title="ͼƬJS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic();">
          <option value="6" <%if JSObj("Manner")="6" then response.write("selected") end if%>>��ʽA</option>
          <option value="7" <%if JSObj("Manner")="7" then response.write("selected") end if%>>��ʽB</option>
          <option value="8" <%if JSObj("Manner")="8" then response.write("selected") end if%>>��ʽC</option>
          <option value="9" <%if JSObj("Manner")="9" then response.write("selected") end if%>>��ʽD</option>
          <option value="10" <%if JSObj("Manner")="10" then response.write("selected") end if%>>��ʽE</option>
          <option value="11" <%if JSObj("Manner")="11" then response.write("selected") end if%>>��ʽF</option>
          <option value="12" <%if JSObj("Manner")="12" then response.write("selected") end if%>>��ʽG</option>
          <option value="13" <%if JSObj("Manner")="13" then response.write("selected") end if%>>��ʽH</option>
          <option value="14" <%if JSObj("Manner")="14" then response.write("selected") end if%>>��ʽI</option>
          <option value="15" <%if JSObj("Manner")="15" then response.write("selected") end if%>>��ʽJ</option>
          <option value="16" <%if JSObj("Manner")="16" then response.write("selected") end if%>>��ʽK</option>
        </select></td>
      <td> 
        <div align="center">�����о�</div></td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace3" value="<%=JSObj("RowSpace")%>"  style="width:100%;" Title="��������������������֮����о࣬��ע��������ֵ��" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" Title="���ű����CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�"  value="<%=JSObj("TitleCSS")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">�¿�����</div></td>
      <td> 
        <select name="OpenMode" id="select5" style="width:100%">
          <option value="1" <%If JSObj("OpenMode")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If JSObj("OpenMode")=0 then Response.Write("selected")%>>��</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="NewsTitleNum" type="text" id="NewsTitleNum2" value="<%=JSObj("NewsTitleNum")%>" Title="ÿ�����ŵı�����ʾ������"   style="width:100%;"></td>
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <select name="ShowTimeTF" id="select6" style="width:100%" onChange="ChooseDate(this.value);" Title="�������������ű�������Ƿ���ʾ�������ŵĸ���ʱ�䣡" >
          <option value="1" <%If JSObj("ShowTimeTF")=1 then Response.Write("selected")%>>����</option>
          <option value="0" <%If JSObj("ShowTimeTF")=0 then Response.Write("selected")%>>������</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="ContentCSS" type="text" id="ContentCSS" Title="�������ݵ�CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�"  value="<%=JSObj("ContentCSS")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" value="<%=JSObj("DateCSS")%>"  style="width:100%;" Title="���������CSS��ʽ��ֱ��������ʽ���Ƽ��ɣ�" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="ContentNum" type="text" id="ContentNum2" Title="Ϊ��Ҫ��ʾ�������ݵ���ʽ����ÿ�����ŵ�������ʾ������"  value="<%=JSObj("ContentNum")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="BackCSS" type="text" id="BackCSS2" value="<%=JSObj("BackCSS")%>"  style="width:100%;" Title="����JS�ı�����ʽ�������ʽ������ֱ��������ʽ���Ƽ��ɡ������ѡ�ô������ã����ÿգ�" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <select name="MoreContent" id="select" style="width:100% " Title="����Ϊ���������ݵ���ʽ�������½Ǽ�һ���ӵ�������ҳ�����ӣ��������ʾ�����ӣ���ѡ�񡰲���ʾ����" >
          <option value="1" <%If JSObj("MoreContent")=1 then Response.Write("selected") %>>��ʾ</option>
          <option value="0" <%If JSObj("MoreContent")=0 then Response.Write("selected") %>>����ʾ</option>
        </select></td>
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <select name="DateType" id="select7" style="width:100%" Title="���ڵ�����ʽ,Ĭ��ΪX��X�գ�" >
          <option value="1" <%if JSObj("DateType") = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
          <option value="2" <%if JSObj("DateType") = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
          <option value="3" <%if JSObj("DateType") = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
          <option value="4" <%if JSObj("DateType") = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
          <option value="5" <%if JSObj("DateType") = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
          <option value="6" <%if JSObj("DateType") = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
          <option value="7" <%if JSObj("DateType") = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
          <option value="8" <%if JSObj("DateType") = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
          <option value="9" <%if JSObj("DateType") = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
          <option value="10" <%if JSObj("DateType") = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
          <option value="11" <%if JSObj("DateType") = "11" then Response.Write("selected") end if%>><%=Month(Now)&"��"&Day(Now)&"��"%></option>
          <option value="12" <%if JSObj("DateType") = "12" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"ʱ"%></option>
          <option value="13" <%if JSObj("DateType") = "13" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"��"%></option>
          <option value="14" <%if JSObj("DateType") = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"ʱ"&Minute(Now)&"��"%></option>
          <option value="15" <%if JSObj("DateType") = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
          <option value="16" <%if JSObj("DateType") = "16" then Response.Write("selected") end if%>><%=Year(Now)&"��"&Month(Now)&"��"&Day(Now)&"��"%></option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" Title="Ϊ��Ҫ��ʾ�������ӵ���ʽ��������������������ͼƬ��ַ�������ͼƬ��ַ������<br>����img src=../img/1.gif border=0������ʽ�����С�src=����ΪͼƬ·������border=0��ΪͼƬ�ޱ߿�"  value="<%=JSObj("LinkWord")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" Title="Ϊ��������ѡ��CSS��ʽ��ֱ������CSS��ʽ���Ƽ��ɣ�"  value="<%=JSObj("LinkCSS")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ���</div></td>
      <td> 
        <input name="PicWidth" type="text" disabled id="PicWidth3" Title="������ΪͼƬ���͵�JS����ͼƬ�Ŀ�Ȳ�����"  value="<%=JSObj("PicWidth")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">ͼƬ�߶�</div></td>
      <td> 
        <input name="PicHeight" type="text" disabled id="PicHeight3" title="������ΪͼƬ���͵�JS����ͼƬ�ĸ߶Ȳ�����"  value="<%=JSObj("PicHeight")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">����ͼƬ</div></td>
      <td colspan="4"> 
        <input name="NaviPic" type="text" id="NaviPic" Title="���ű���ǰ��ĵ���ͼ�꣬�����ǡ��������ַ���Ҳ������ͼƬ��ַ�������ͼƬ��ַ������<br>����img src=../img/1.gif border=0������ʽ�����С�src=����ΪͼƬ·������border=0��ΪͼƬ�ޱ߿�"  value="<%=JSObj("NaviPic")%>" style="width:100%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">�м�ͼƬ</div></td>
      <td colspan="4"> 
        <input name="RowBettween" type="text" id="RowBettween" style="width:80%;" size="26" value="<%=JSObj("RowBettween")%>" Title="��������������������֮��ļ��ͼƬ��������ѡ��ͼƬ����ť�������ã����Ϊ�գ�" > 
        <input id="RowBettweenButton" type="button" name="Submit34" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.RowBettween);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ��ַ</div></td>
      <td colspan="4"> 
        <input name="PicPath" type="text" id="PicPath" value="<%=JSObj("PicPath")%>" style="width:80%;" size="26" disabled Title="Ϊ����һ��ͼƬ����ʽ����ͼƬ��������ѡ��ͼƬ����ťѡ��ͼƬ��" > 
        <input id="PicChooseButton" type="button" name="Submit34" value="ѡ��ͼƬ" disabled onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.PicPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;ע</div></td>
      <td colspan="4"> 
        <textarea name="Info" rows="6" id="Info" style="width:100%" Title="��ע�����ڴ������ʱ����鿴���ԣ�" ><%=JSObj("Info")%></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<script> 
var TempDateScr = '<% = TempDateStr%>';
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	DocumentReadyTF=true;
	TypeChoose();
	ChoosePic();
	ChooseDate(TempDateScr);
}
function TypeChoose()
{
	if (document.JSForm.TypeW.checked==true)
	{ 
		document.JSForm.MannerW.disabled=false;
		document.JSForm.MannerP.disabled=true;
		document.JSForm.PicPath.disabled=true;
		document.JSForm.PicChooseButton.disabled=true;
		document.JSForm.PicWidth.disabled=true;
		document.JSForm.PicHeight.disabled=true;
	}
	else
	{
		document.JSForm.MannerW.disabled=true;
		document.JSForm.MannerP.disabled=false;
		document.JSForm.PicPath.disabled=false;
		document.JSForm.PicChooseButton.disabled=false;
		document.JSForm.PicWidth.disabled=false;
		document.JSForm.PicHeight.disabled=false;
	}
}
  
function ShowTitle(TempStr)
{
	document.all.TempTip.innerHTML='<font color=red>��ʾ��</font><br><br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>'+TempStr+'</font>';
}
   
function ChooseDate(DateStr)
{ 
	if (DateStr==1)
	{
		document.JSForm.DateType.disabled=false;
		document.JSForm.DateCSS.disabled=false;
	}
	else
	{
		document.JSForm.DateType.disabled=true;
		document.JSForm.DateCSS.disabled=true;
	}
}
 
function ChoosePic()
{
	if (document.JSForm.MannerW.disabled==false) 
		 document.all.PreviewArea.innerHTML='<img src="Img/Css'+document.JSForm.MannerW.value+'.gif" border="0">';
	else 
		 document.all.PreviewArea.innerHTML='<img src="Img/Css'+document.JSForm.MannerP.value+'.gif" border="0">';
}
</script>
<%
  if Request.Form("action")="mod" then
     dim CNameWordNum,CNameStr,ENameWordNum,ENameStr,JSAddObj,JSNewsNum,JSNewsTitleNum,JSRowNum,JSContentNum,RsJSObj,RsJSSql
	 if NoCSSHackAdmin(request.form("CName"),"����")<>"" then
	    CNameStr = Replace(Replace(request.form("CName"),"""",""),"'","")
		CNameWordNum = Cint(Len(CNameStr))
		if CNameWordNum>25 then
			 response.Write("<script>alert(""�������Ʋ��ܶ���25���ַ�"");history.back();</script>")
			 response.end
		end if
	  else
		 response.Write("<script>alert(""��������������"");history.back();</script>")
		 response.end
	 end if
  	 if isnumeric(request.form("NewsNum")) = false then
		 response.Write("<script>alert(""���ŵ���������Ϊ������"");history.back();</script>")
		 response.end
	 else
		 JSNewsNum = Cint(request.form("NewsNum"))
	 end if
  	 if isnumeric(request.form("NewsTitleNum")) = false then
		 response.Write("<script>alert(""���ű�����������Ϊ������"");history.back();</script>")
		 response.end
	 else
		 JSNewsTitleNum = Cint(request.form("NewsTitleNum"))
	 end if
  	 if isnumeric(request.form("RowNum")) = false or request.form("RowNum")="0" then
		 response.Write("<script>alert(""���Ų�����������Ϊ�������Ҳ���Ϊ0"");history.back();</script>")
		 response.end
	 else
		 JSRowNum = Cint(request.form("RowNum"))
	 end if
  	 if isnumeric(request.form("ContentNum")) = false then
		 response.Write("<script>alert(""����������������Ϊ������"");history.back();</script>")
		 response.end
	 else
		 JSContentNum = Cint(request.form("ContentNum"))
	 end if
	  Set RsJSObj=server.createobject(G_FS_RS)
	  RsJSSql="select * from FS_FreeJS where ID = "&JSID&""
	  RsJSObj.open RsJSSql,Conn,1,3
	  Dim TempEName
	  TempEName = Cstr(RsJSObj("EName"))
	  RsJSObj("CName") = Cstr(CNameStr)
	  RsJSObj("Type") = Cint(Replace(Replace(Request.Form("Type"),"""",""),"'",""))
	  if Request.Form("Type") = "0" then
		  RsJSObj("Manner") = Cint(Replace(Replace(Request.Form("Manner"),"""",""),"'",""))
	  else
		  RsJSObj("Manner") = Cint(Replace(Replace(Request.Form("MannerP"),"""",""),"'",""))
	  end if
	  if Request.Form("PicWidth")<>"" and isnull(Request.Form("PicWidth"))=false then
	     if isnumeric(Request.Form("PicWidth"))=true then
			  RsJSObj("PicWidth") = Cint(Request.Form("PicWidth"))
	      else
			 response.Write("<script>alert(""ͼƬ��ȱ���Ϊ������"");history.back();</script>")
			 response.end
		  end if
	  end if
	  if Request.Form("PicHeight")<>"" and isnull(Request.Form("PicHeight"))=false then
	     if isnumeric(Request.Form("PicHeight"))=true then
			  RsJSObj("PicHeight") = Cint(Request.Form("PicHeight"))
	      else
			 response.Write("<script>alert(""ͼƬ�߶ȱ���Ϊ������"");history.back();</script>")
			 response.end
		  end if
	  end if
	  RsJSObj("NewsNum") = Cint(JSNewsNum)
	  RsJSObj("NewsTitleNum") = Cint(JSNewsTitleNum)
	  RsJSObj("TitleCSS") = Cstr(Request.Form("TitleCSS"))
	  RsJSObj("ContentCSS") = Cstr(Request.Form("ContentCSS"))
	  RsJSObj("BackCSS") = Cstr(Request.Form("BackCSS"))
	  RsJSObj("RowNum") = Cint(JSRowNum)
	  if Request.Form("MannerP")="12" or Request.Form("MannerP")="16" then
		  if Replace(Replace(Request.Form("PicPath"),"""",""),"'","")<>"" then
			  RsJSObj("PicPath") = Cstr(Request.Form("PicPath"))
		  else
			 response.Write("<script>alert(""��ѡ��ͼƬ��ַ"");history.back();</script>")
			 response.end
		  end if
	  end if
	  if Request.Form("ShowTimeTF")="1" then
		  RsJSObj("ShowTimeTF") = Cint(Request.Form("ShowTimeTF"))
	   else
		  RsJSObj("ShowTimeTF") = "0"
	  end if
	  RsJSObj("ContentNum") = Cint(JSContentNum)
	  RsJSObj("NaviPic") = Cstr(Request.Form("NaviPic"))
	  if Request.Form("DateType")="" or isnull(Request.Form("DateType")) or isnumeric(Request.Form("DateType"))=false then
		  RsJSObj("DateType") = "11"
	  else
		  RsJSObj("DateType") = Cint(Request.Form("DateType"))
	  end if
	  RsJSObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	  RsJSObj("Info") = Request.Form("Info")
	  RsJSObj("MoreContent") = Request.Form("MoreContent")
	  if Request.Form("MoreContent")=1 then
		  If Request.Form("LinkWord")<>"" and isnull(Request.Form("LinkWord"))=false then
			  RsJSObj("LinkWord") = Request.Form("LinkWord")
		  Else
		  	Response.Write("<script>alert(""�������������ֻ�ͼƬ"");</script>")
			Response.End
		  End If
		  If Request.Form("LinkCSS")<>"" or isnull(Request.Form("LinkCSS"))=false then
			  RsJSObj("LinkCSS") = Request.Form("LinkCSS")
		  End If
	  End If
	  If isnumeric(Request.Form("RowSpace")) then
		  RsJSObj("RowSpace") = Cint(Request.Form("RowSpace"))
	  Else
		  RsJSObj("RowSpace") = 2
	  End If
	  RsJSObj("RowBettween") = Request.Form("RowBettween")
	  RsJSObj("OpenMode") = Request.Form("OpenMode")
	  RsJSObj.update
	  RsJSObj.close
	  Set RsJSObj = Nothing
  '--------------------��Ҫ��������JS�ļ�---------------------------------
	Dim JSClassObj,ReturnValue
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = TempSysRootDir
	Dim RefreshManner
	if Request.Form("Type") = "0" then
	  RefreshManner = Cint(Replace(Replace(Request.Form("Manner"),"""",""),"'",""))
	else
	  RefreshManner = Cint(Replace(Replace(Request.Form("MannerP"),"""",""),"'",""))
	end if
	Select case RefreshManner
		case "1"   ReturnValue = JSClassObj.WCssA(TempEName,True)
		case "2"   ReturnValue = JSClassObj.WCssB(TempEName,True)
		case "3"   ReturnValue = JSClassObj.WCssC(TempEName,True)
		case "4"   ReturnValue = JSClassObj.WCssD(TempEName,True)
		case "5"   ReturnValue = JSClassObj.WCssE(TempEName,True)
		case "6"   ReturnValue = JSClassObj.PCssA(TempEName,True)
		case "7"   ReturnValue = JSClassObj.PCssB(TempEName,True)
		case "8"   ReturnValue = JSClassObj.PCssC(TempEName,True)
		case "9"   ReturnValue = JSClassObj.PCssD(TempEName,True)
		case "10"   ReturnValue = JSClassObj.PCssE(TempEName,True)
		case "11"   ReturnValue = JSClassObj.PCssF(TempEName,True)
		case "12"  ReturnValue = JSClassObj.PCssG(TempEName,True)
		case "13"   ReturnValue = JSClassObj.PCssH(TempEName,True)
		case "14"   ReturnValue = JSClassObj.PCssI(TempEName,True)
		case "15"   ReturnValue = JSClassObj.PCssJ(TempEName,True)
		case "16"   ReturnValue = JSClassObj.PCssK(TempEName,True)
		case "17"   ReturnValue = JSClassObj.PCssL(TempEName,True)
	End Select
	Set JSClassObj = Nothing
  '--------------------��Ҫ��������JS�ļ�---------------------------------
	if ReturnValue <> "" then
		Response.Write("<script>alert('" & ReturnValue & "');location='FreeJSList.asp'</script>")
	else
		Response.Redirect("FreeJSList.asp")
	end if
  end if

end if
Set Conn = Nothing
%>
