<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->

<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================

%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070501") then Call ReturnError1()
%>
<html>
<head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������ӹ���</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" method = "post" name ="FriendLinkForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="����" onClick="document.FriendLinkForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input type="hidden" name="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="14%"> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input name="Name" type="text" id="Name" style="width:92%" value="<%=Request("Name")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
      <td> 
        <input name="Type" type="radio" id="TypeFL" onclick="ChoosePic();" value="0" <% if Request("Type")="0" then response.Write("checked") end if%>>
        ���� 
        <input name="Type" type="radio" id="TypeFLP" onclick="ChoosePic();" value="1" <% if Request("Type")="1" or Request("Type")="" then Response.Write("checked") end if%>>
        ͼƬ</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">��ʾ����</div></td>
      <td> 
        <input name="Content" type="text" id="Content" size="35" value="<%=Request("Content")%>"> 
        <input id="PicChoose" type="button" name="PicChoose" value="ѡ��ͼƬ"  onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,290,window,document.FriendLinkForm.Content);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">Ӧ��ҳ��</div></td>
      <td> 
        <input name="AddressIndex" type="checkbox" id="AddressIndex" value="1" <%if Request("AddressIndex") = "1" then Response.Write("checked") end if%>>
        ��ҳ 
        <input name="AddressClass" type="checkbox" id="AddressClass" value="2" <%if Request("AddressClass") = "1" then Response.Write("checked") end if%>>
        ��Ŀ 
        <input name="AddressNews" type="checkbox" id="AddressNews" value="3" <%if Request("AddressNews") = "1" then Response.Write("checked") end if%>>
        ���� 
        <input name="AddressSpecial" type="checkbox" id="AddressSpecial" value="4" <%if Request("AddressSpecial") = "1" then Response.Write("checked") end if%>>
        ר��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">���ӵ�ַ</div></td>
      <td> 
        <input name="UrlT" type="text" id="Url2"  style="width:92%" value="<%if Request("UrlT")<>"" then Response.Write(Request("UrlT")) else Response.Write("http://") end if%>"></td>
    </tr>
  </table>
</form>
</body>
</html>
<script>
function ChoosePic()
{
    if (document.FriendLinkForm.TypeFL.checked==true)
	   {
		document.FriendLinkForm.PicChoose.disabled=true;
		}
	else
	   {
		document.FriendLinkForm.PicChoose.disabled=false;
		}
 }
</script>
<%
  if request.form("action") = "add" then
     Dim FLAddObj,FLAddSql,FLAddress,FLName,FLContent,FLUrl,FLAddressIndex,FLAddressClass,FLAddressNews,FLAddressSpecial
	     if NoCSSHackAdmin(request.Form("Name"),"����")<>"" then
		    FLName = Replace(replace(Request.form("Name"),"'",""),"""","")
		  else
			  response.write("<script>alert(""����д������������"");location=""javascript:history.back(-1)"";</script>")
			  response.end
		 end if
		 if request.form("Content")<>"" then
			 FLContent = Replace(replace(Request.form("Content"),"'",""),"""","")
		 else
			  response.write("<script>alert(""����д������������"");location=""javascript:history.back(-1)"";</script>")
			  response.end
		 end if
		 if request.form("UrlT")<> "" then
			 FLUrl = Replace(replace(Request.form("UrlT"),"'",""),"""","")
		 else
			  response.write("<script>alert(""����д�������ӵ�ַ"");location=""javascript:history.back(-1)"";</script>")
			  response.end
		 end if
		 if Request.Form("AddressIndex")="" and Request.Form("AddressClass")="" and Request.Form("AddressNews")="" and Request.Form("AddressSpecial")="" then
			 FLAddress = 0
		  else
			 FLAddress = Cint(Request.Form("AddressIndex")&Request.Form("AddressClass")& Request.Form("AddressNews")&Request.Form("AddressSpecial"))
		  end if
		  Set FLAddObj=server.createobject(G_FS_RS)
		  FLAddSql="select * from FS_FriendLink"
		  FLAddObj.open FLAddSql,Conn,3,3
		  FLAddObj.addnew
		  FLAddObj("Name") = Cstr(FLName)
		  FLAddObj("Content") = Cstr(FLContent)
		  FLAddObj("Url") = Cstr(FLUrl)
		  FLAddObj("Type") = Replace(replace(Request.form("Type"),"'",""),"""","")
		  FLAddObj("Address") = FLAddress
		  FLAddObj("AddTime") = Now()
		  FLAddObj.update
		  FLAddObj.Close
		  Set FLAddObj = Nothing
		Response.Redirect("OrdinaryFriendLink.asp")
		response.end
  end if
  Set Conn = Nothing
%>
