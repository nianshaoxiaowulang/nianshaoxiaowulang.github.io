<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<!--#include file="../../../Inc/ThumbnailFunction.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P060100") then Call ReturnError1()
Dim TempSysRootDir
if SysRootDir = "" then
	TempSysRootDir = ""
else
	TempSysRootDir = "/" & SysRootDir
end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����JS���</title>
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
          <td>&nbsp; <input name="action" type="hidden" id="action2" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E7E7E7">
    <tr bgcolor="#FFFFFF"> 
      <td width="10%"> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td colspan="3"> 
        <input name="CName" type="text" id="CName" style="width:100%" title="JS���������ƣ����ں�̨���ĺ͹����벻Ҫ����25���ַ���" value="<%=request("CName")%>" maxlength="25"> 
        <div align="center"></div></td>
      <td rowspan="11" align="center" valign="middle" id="PreviewArea"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">Ӣ������</div></td>
      <td colspan="3"> 
        <input name="EName" type="text" id="EName" value="<%=request("EName")%>" style="width:100%" title="JS��Ӣ�����ƣ�����ǰ̨���ã��벻Ҫ����50���ַ��Ҳ������Ѿ����ڵ�JS������"> 
        <div align="center"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td width="20%"> 
        <input id="TypeW" name="Type" type="radio" value="0" <%if request("Type")="" or request("Type")="0" then response.write("checked") end if%> onClick="TypeChoose();ChoosePic();" title="JS���ͣ����֣�ѡ��">
        ���� 
        <input id="TypeP" type="radio" name="Type" value="1" <%if request("Type")="1" then response.write("checked") end if%> onClick="TypeChoose();ChoosePic();" title="JS���ͣ�ͼƬ��ѡ��">
        ͼƬ</td>
      <td width="10%" valign="middle"> 
        <div align="center">��������</div></td>
      <td width="20%" valign="middle"> 
        <input name="NewsNum" type="text" id="NewsNum3" value="<%if request("NewsNum")<>"" then response.write(request("NewsNum")) else response.write("5") end if%>" title="��JS������õ�����������" style="width:100% "></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <select name="Manner" id="MannerW" style="width:100% " title="����JS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic();">
          <option value="1" <%if request("Manner")="1" then response.write("selected") end if%>>��ʽA</option>
          <option value="2" <%if request("Manner")="2" then response.write("selected") end if%>>��ʽB</option>
          <option value="3" <%if request("Manner")="3" then response.write("selected") end if%>>��ʽC</option>
          <option value="4" <%if request("Manner")="4" then response.write("selected") end if%>>��ʽD</option>
          <option value="5" <%if request("Manner")="5" then response.write("selected") end if%>>��ʽE</option>
        </select> </td>
      <td valign="middle"> 
        <div align="center">��������</div></td>
      <td valign="middle"> 
        <input name="RowNum" type="text" id="RowNum3" title="ShowTitle('��������JS��ÿ������ʾ����������������ز�Ҫ��Ϊ��0����" value="<%if request("RowNum")<>"" then response.write(request("RowNum")) else response.write("2") end if%>" style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ��ʽ</div></td>
      <td> 
        <select name="MannerP" id="MannerP" style="width:100% " disabled title="ͼƬJS��ʽѡ�������д���ʽ��Ԥ����" onChange="ChoosePic();">
          <option value="6" <%if request("Manner")="6" then response.write("selected") end if%>>��ʽA</option>
          <option value="7" <%if request("Manner")="7" then response.write("selected") end if%>>��ʽB</option>
          <option value="8" <%if request("Manner")="8" then response.write("selected") end if%>>��ʽC</option>
          <option value="9" <%if request("Manner")="9" then response.write("selected") end if%>>��ʽD</option>
          <option value="10" <%if request("Manner")="10" then response.write("selected") end if%>>��ʽE</option>
          <option value="11" <%if request("Manner")="11" then response.write("selected") end if%>>��ʽF</option>
          <option value="12" <%if request("Manner")="12" then response.write("selected") end if%>>��ʽG</option>
          <option value="13" <%if request("Manner")="13" then response.write("selected") end if%>>��ʽH</option>
          <option value="14" <%if request("Manner")="14" then response.write("selected") end if%>>��ʽI</option>
          <option value="15" <%if request("Manner")="15" then response.write("selected") end if%>>��ʽJ</option>
          <option value="16" <%if request("Manner")="16" then response.write("selected") end if%>>��ʽK</option>
        </select></td>
      <td valign="middle"> 
        <div align="center">�����о�</div></td>
      <td valign="middle"> 
        <input name="RowSpace" type="text" id="RowSpace3" value="<%=Request("RowSpace")%>" style="width:100%;" title="��������������������֮����о࣬��ע��������ֵ��"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" title="���ű����CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�" value="<%=request("TitleCSS")%>" style="width:100%;"></td>
      <td valign="middle"> 
        <div align="center">�¿�����</div></td>
      <td valign="middle"> 
        <select name="OpenMode" id="select" style="width:100%;">
          <option value="1" <%If Request("OpenMode")=1 then Response.Write("selected")%>>��</option>
          <option value="0" <%If Request("OpenMode")=0 then Response.Write("selected")%>>��</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="NewsTitleNum" type="text" id="NewsTitleNum2" value="<%if request("NewsTitleNum")<>"" then response.write(request("NewsTitleNum")) else response.write("5") end if%>" title="ÿ�����ŵı�����ʾ������;" style="width:100%;"></td>
      <td valign="middle"> 
        <div align="center">��������</div></td>
      <td valign="middle"> 
        <select name="ShowTimeTF" id="select4" style="width:100%;" onChange="ChooseDate(this.value);" title="�������������ű�������Ƿ���ʾ�������ŵĸ���ʱ�䣡">
          <option value="1" <%If Request("ShowTimeTF")=1 then Response.Write("selected")%>>����</option>
          <option value="0" <%If Request("ShowTimeTF")=0 then Response.Write("selected")%>>������</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">������ʽ</div></td>
      <td> 
        <input name="ContentCSS" type="text" id="ContentCSS" title="�������ݵ�CSS��ʽ����ֱ��������ʽ���ơ������ѡ�ô������ã����ÿգ�" value="<%=request("ContentCSS")%>" style="width:100% "></td>
      <td valign="middle"> 
        <div align="center">������ʽ</div></td>
      <td valign="middle"> 
        <input name="DateCSS" type="text" id="DateCSS" value="<%=request("DateCSS")%>" style="width:100%;" title="���������CSS��ʽ��ֱ��������ʽ���Ƽ��ɣ�"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="ContentNum" type="text" id="ContentNum2" title="Ϊ��Ҫ��ʾ�������ݵ���ʽ����ÿ�����ŵ�������ʾ������" value="<%if request("ContentNum")<>"" then response.write(request("ContentNum")) else response.write("30") end if%>" style="width:100% "></td>
      <td valign="middle"> 
        <div align="center">������ʽ</div></td>
      <td valign="middle"> 
        <input style="width:100%;" name="BackCSS" type="text" id="BackCSS2" value="<%=request("BackCSS")%>" size="14" title="����JS�ı�����ʽ�������ʽ������ֱ��������ʽ���Ƽ��ɡ������ѡ�ô������ã����ÿգ�"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <select name="MoreContent" id="select2" style="width:100%;" title="����Ϊ���������ݵ���ʽ�������½Ǽ�һ���ӵ�������ҳ�����ӣ��������ʾ�����ӣ���ѡ�񡰲���ʾ����">
          <option value="1" <%If Request("MoreContent")=1 then Response.Write("selected") %>>��ʾ</option>
          <option value="0" <%If Request("MoreContent")=0 then Response.Write("selected") %>>����ʾ</option>
        </select></td>
      <td valign="middle"> 
        <div align="center">������ʽ</div></td>
      <td valign="middle"> 
        <select name="DateType" id="select5" style="width:100%;" title="���ڵ�����ʽ,Ĭ��ΪX��X�գ�">
          <option value="1" <%if Request("DateType") = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
          <option value="2" <%if Request("DateType") = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
          <option value="3" <%if Request("DateType") = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
          <option value="4" <%if Request("DateType") = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
          <option value="5" <%if Request("DateType") = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
          <option value="6" <%if Request("DateType") = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
          <option value="7" <%if Request("DateType") = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
          <option value="8" <%if Request("DateType") = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
          <option value="9" <%if Request("DateType") = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
          <option value="10" <%if Request("DateType") = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
          <option value="11" <%if Request("DateType") = "11" then Response.Write("selected") end if%>><%=Month(Now)&"��"&Day(Now)&"��"%></option>
          <option value="12" <%if Request("DateType") = "12" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"ʱ"%></option>
          <option value="13" <%if Request("DateType") = "13" then Response.Write("selected") end if%>><%=day(Now)&"��"&Hour(Now)&"��"%></option>
          <option value="14" <%if Request("DateType") = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"ʱ"&Minute(Now)&"��"%></option>
          <option value="15" <%if Request("DateType") = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
          <option value="16" <%if Request("DateType") = "16" then Response.Write("selected") end if%>><%=Year(Now)&"��"&Month(Now)&"��"&Day(Now)&"��"%></option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" title="Ϊ��Ҫ��ʾ�������ӵ���ʽ��������������������ͼƬ��ַ�������ͼƬ��ַ������<br>����img src=../img/1.gif border=0������ʽ�����С�src=����ΪͼƬ·������border=0��ΪͼƬ�ޱ߿�" value="<%=Request("LinkWord")%>" style="width:100%;"></td>
      <td valign="middle"> 
        <div align="center">������ʽ</div></td>
      <td valign="middle"> 
        <input name="LinkCSS" type="text" id="LinkCSS" title="Ϊ��������ѡ��CSS��ʽ��ֱ������CSS��ʽ���Ƽ��ɣ�" value="<%=Request("LinkCSS")%>" style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ���</div></td>
      <td> 
        <input name="PicWidth" type="text" disabled id="PicWidth4" onFocus="ShowTitle('������Ϊ����ͼƬ���͵�JS����ͼƬ�Ŀ�Ȳ�����');" onBlur="ShowTitle('');" value="<%if request("PicWidth")<>"" then response.write(request("PicWidth")) else response.write("60") end if%>" size="14" style="width:100%;"></td>
      <td> 
        <div align="center">ͼƬ�߶�</div></td>
      <td> 
        <input name="PicHeight" type="text" disabled id="PicHeight3" onFocus="ShowTitle('������Ϊ����ͼƬ���͵�JS����ͼƬ�ĸ߶Ȳ�����');" onBlur="ShowTitle('');" value="<%if request("PicHeight")<>"" then response.write(request("PicHeight")) else response.write("60") end if%>" size="14" style="width:100%;"></td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">����ͼƬ</div></td>
      <td colspan="4"> 
        <input name="NaviPic" type="text" id="NaviPic2" title="���ű���ǰ��ĵ���ͼ�꣬�����ǡ��������ַ���Ҳ������ͼƬ��ַ�������ͼƬ��ַ������<br>����img src=../img/1.gif border=0������ʽ�����С�src=����ΪͼƬ·������border=0��ΪͼƬ�ޱ߿�" value="<%=request("NaviPic")%>" style="width:100%; "></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">�м�ͼƬ</div></td>
      <td colspan="4"> 
        <input name="RowBettween" type="text" id="RowBettween" size="26" value="<%=request("RowBettween")%>" title="��������������������֮��ļ��ͼƬ��������ѡ��ͼƬ����ť�������ã����Ϊ�գ�" style="width:80%;"> 
        <input id="RowBettweenButton" type="button" name="Submit34" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.RowBettween);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">ͼƬ��ַ</div></td>
      <td colspan="4"> 
        <input name="PicPath" type="text" id="PicPath" value="<%=request("PicPath")%>" style="width:80%;" disabled title="Ϊ����һ��ͼƬ����ʽ����ͼƬ��������ѡ��ͼƬ����ťѡ��ͼƬ��"> 
        <input id="PicChooseButton" type="button" name="Submit34" value="ѡ��ͼƬ" disabled onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.PicPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;ע</div></td>
      <td colspan="4"> 
        <textarea name="Info" rows="6" id="Info" style="width:100%" Title="��ע�����ڴ������ʱ����鿴���ԣ�"><%=request("Info")%></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<script> 
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	DocumentReadyTF=true;
	ChoosePic();
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
//	document.all.TempTip.innerHTML='<font color=red>��ʾ��</font><br><br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>'+TempStr+'</font>';
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
  if Request.Form("action")="add" then
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
	 if request.form("EName")<>"" then
	    ENameStr = Replace(Replace(request.form("EName"),"""",""),"'","")
		ENameWordNum = Cint(Len(ENameStr))
		if ENameWordNum>=50 then
			 response.Write("<script>alert(""Ӣ�����Ʋ��ܶ���50���ַ�"");history.back();</script>")
			 response.end
		end if
		Set JSAddObj = Conn.Execute("select EName from FS_FreeJS where EName='"&ENameStr&"'")
		if Not JSAddObj.eof then
			 response.Write("<script>alert(""Ӣ�������ظ�,��������д"");history.back();</script>")
			 response.end
		end if
		JSAddObj.close
		Set JSAddObj = Nothing
	  else
		 response.Write("<script>alert(""������Ӣ������"");history.back();</script>")
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
	  RsJSSql="select * from FS_FreeJS"
	  RsJSObj.open RsJSSql,Conn,3,3
	  RsJSObj.addnew 
	  RsJSObj("EName") = Cstr(ENameStr)
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
	  if Replace(Replace(Request.Form("TitleCSS"),"""",""),"'","")<>"" then
		  RsJSObj("TitleCSS") = Cstr(Request.Form("TitleCSS"))
	  end if
	  if Replace(Replace(Request.Form("ContentCSS"),"""",""),"'","")<>"" then
		  RsJSObj("ContentCSS") = Cstr(Request.Form("ContentCSS"))
	  end if
	  if Replace(Replace(Request.Form("BackCSS"),"""",""),"'","")<>"" then
		  RsJSObj("BackCSS") = Cstr(Request.Form("BackCSS"))
	  end if
	  RsJSObj("RowNum") = Cint(JSRowNum)
	  Dim OpenCreateThumbnail,CreateSmallPicOK
	  CreateSmallPicOK=False
 	  OpenCreateThumbnail=Conn.Execute("Select ThumbnailComponent from FS_Config")(0)
	  if Request.Form("MannerP")="12" or Request.Form("MannerP")="16" then
		  if Replace(Replace(Request.Form("PicPath"),"""",""),"'","")<>"" then
				'======================================
				'���ϵͳ��������������ͼ���� ����������ͼ
				If OpenCreateThumbnail=1 then 
					Dim sRootDir,PicFileName
					PicFileName=mid(Request.Form("PicPath"),InStrRev(Request.Form("PicPath"),"/")+1)
					sRootDir=TempSysRootDir& left(Request.Form("PicPath"),instrrev(Request.Form("PicPath"),"/"))
					CreateSmallPicOK=CreateThumbnail(sRootDir&PicFileName,Request.Form("PicWidth"),Request.Form("PicHeight"),"0",sRootDir&"s_"&PicFileName)'��ԭͼƬ����ָ����Ⱥ͸߶ȵ�����ͼ,����ɹ�����True,ʧ�ܷ���False
					'=======================================
					If CreateSmallPicOK=True then
						RsJSObj("PicPath") =left(Request.Form("PicPath"),InStrRev(Request.Form("PicPath"),"/"))&"s_"&PicFileName
					Else
						RsJSObj("PicPath") =Replace(Replace(Request.Form("PicPath"),"""",""),"'","")
					End If
				Else
					RsJSObj("PicPath") =Replace(Replace(Request.Form("PicPath"),"""",""),"'","")
				End If
		  else
				response.Write("<script>alert(""��ѡ��ͼƬ��ַ"");history.back();</script>")
				response.end
		  end if
	  end if
	  RsJSObj("ShowTimeTF") = Cint(Request.Form("ShowTimeTF"))
	  RsJSObj("AddTime") = Now()
	  RsJSObj("ContentNum") = Cint(JSContentNum)
	  if Replace(Replace(Request.Form("NaviPic"),"""",""),"'","")<>"" then
		  RsJSObj("NaviPic") = Cstr(Request.Form("NaviPic"))
	  end if
	  if Request.Form("DateCSS")<>"" and isnull(Request.Form("DateColor"))=false then
		  RsJSObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	  end if
	  if Request.Form("DateType")<>"" or isnull(Request.Form("DateType"))=false then
		  RsJSObj("DateType") = Cint(Request.Form("DateType"))
	  end if
	  RsJSObj("Info") = Request.Form("Info")
	  RsJSObj("MoreContent") = Request.Form("MoreContent")
	  if Request.Form("MoreContent")=1 then
		  If Request.Form("LinkWord")<>"" and  isnull(Request.Form("LinkWord"))=false then
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
	  RsJSObj.Update
	  RsJSObj.Close
	  Set RsJSObj = Nothing
  '--------------------��Ҫ����JS�ļ�---------------------------------
	Dim JSClassObj,ReturnValue
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = TempSysRootDir
	Select case Request.Form("Manner")
		case "1"   ReturnValue = JSClassObj.WCssA(ENameStr,True)
		case "2"   ReturnValue = JSClassObj.WCssB(ENameStr,True)
		case "3"   ReturnValue = JSClassObj.WCssC(ENameStr,True)
		case "4"   ReturnValue = JSClassObj.WCssD(ENameStr,True)
		case "5"   ReturnValue = JSClassObj.WCssE(ENameStr,True)
		case "6"   ReturnValue = JSClassObj.PCssA(ENameStr,True)
		case "7"   ReturnValue = JSClassObj.PCssB(ENameStr,True)
		case "8"   ReturnValue = JSClassObj.PCssC(ENameStr,True)
		case "9"   ReturnValue = JSClassObj.PCssD(ENameStr,True)
		case "10"   ReturnValue = JSClassObj.PCssE(ENameStr,True)
		case "11"   ReturnValue = JSClassObj.PCssF(ENameStr,True)
		case "12"  ReturnValue = JSClassObj.PCssG(ENameStr,True)
		case "13"   ReturnValue = JSClassObj.PCssH(ENameStr,True)
		case "14"   ReturnValue = JSClassObj.PCssI(ENameStr,True)
		case "15"   ReturnValue = JSClassObj.PCssJ(ENameStr,True)
		case "16"   ReturnValue = JSClassObj.PCssK(ENameStr,True)
		case "17"   ReturnValue = JSClassObj.PCssL(ENameStr,True)
	End Select
	Set JSClassObj = Nothing
	Dim TempFreeJSID
	TempFreeJSID = Conn.Execute("Select ID from FS_FreeJS where EName='" & ENameStr & "'")("ID")
	if ReturnValue <> "" then
		Response.Write("<script>alert('" & ReturnValue & "');location='FreeJSList.asp'</script>")
	else
		Response.Redirect("FreeJSList.asp")
	end if
end if
%>
