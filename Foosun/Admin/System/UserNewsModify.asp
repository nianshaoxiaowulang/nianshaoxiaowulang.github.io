<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Md5.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040401") then Call ReturnError1()
If  Request.Form("action") = "add" then
    Dim UserAddObj,UserAddSql,ChooseMemNameObj,MemNameStr
	If NoCSSHackAdmin(Request.Form("Title"),"����")="" or isnull(Request.Form("Title")) then
		Response.Write("<script>alert(""����д����"");location=""javascript:history.back()"";</script>")
		Response.End
	Else
	End If
	If len(Request.Form("Title"))>100 then
		Response.Write("<script>alert(""���ⲻ���Գ���100���ַ�"");location=""javascript:history.back()"";</script>")
		Response.End
	End If 
	If Request.Form("Content")="" then
		Response.Write("<script>alert(""����д����"");location=""javascript:history.back()"";</script>")
		Response.End
	End If
	Set UserAddObj = Server.CreateObject(G_FS_RS)
		UserAddSql = "Select * from FS_MemberNews where ID="&clng(Request.Form("id"))
		UserAddObj.Open UserAddSql,Conn,3,3
		UserAddObj("Title") = Replace(Replace(Request.Form("Title"),"""",""),"'","")
		UserAddObj("Content") = Request.Form("Content")
		UserAddObj("Popid") = Cint(Request.Form("Popid"))
		UserAddObj("Author") = Replace(Replace(Request.Form("Author"),"""",""),"'","")
		If Request.Form("isLock") = "0" then
			UserAddObj("isLock") = 0
		Else
			UserAddObj("isLock") = 1
		End If
		UserAddObj("Addtime") = Request.Form("addtime")
		UserAddObj.Update
		UserAddObj.Close
		Set UserAddObj = Nothing
		Response.Redirect("SysUserNews.asp")
		Response.End
End If
Dim RsUserModifyObj,UserModifySQL
Set RsUserModifyObj = Server.CreateObject(G_FS_RS)
UserModifySQL = "Select * From Fs_MemberNews where id="&Clng(Request("ID"))
RsUserModifyObj.Open UserModifySQL,Conn,1,3
Dim NewsContent
NewsContent = Replace(Replace(RsUserModifyObj("Content"),"""","%22"),"'","%27")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӻ�Ա</title>
</head>
<body leftmargin="2" topmargin="2">
<form action="" method="post" name="NewsForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����"  onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input name="action" type="hidden" id="action" value="add">
              <input type="hidden" name="Content" value="<% = NewsContent %>">
              <input name="ID" type="hidden" id="ID" value="<% = RsUserModifyObj("ID")%>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="3"></td>
    </tr>
  </table>
  <table width="100%" height="168"  border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
    <tr> 
      <td width="81" height="30" bgcolor="#F2F2F2"> 
        <div align="center">���⡡</div></td>
      <td width="909" colspan="3" bgcolor="#F2F2F2"> 
        <input name="Title" type="text"  id="Title" style="width:100%" value="<% = RsUserModifyObj("Title")%>"></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="30"> 
        <div align="center">�����ˡ�</div></td>
      <td colspan="3"> 
        <input name="Author" type="text" id="Author" style="width:100%" value="<% = RsUserModifyObj("Author")%>"></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="27"> 
        <div align="center">���Ȩ�ޡ�</div></td>
      <td colspan="3"> 
        <select name="PoPid" id="PoPid">
          <option value="0" <%If RsUserModifyObj("PoPid") = 0 then Response.Write("selected")%> >������</option>
          <option value="1" <%If RsUserModifyObj("PoPid") = 1 then Response.Write("selected")%>>һ���Ա</option>
          <option value="2" <%If RsUserModifyObj("PoPid") = 2 then Response.Write("selected")%>>�м���Ա</option>
          <option value="3" <%If RsUserModifyObj("PoPid") = 3 then Response.Write("selected")%>>�߼���Ա</option>
          <option value="4" <%If RsUserModifyObj("PoPid") = 4 then Response.Write("selected")%>>VIP��Ա</option>
        </select></td>
    </tr>
    <tr> 
      <td height="19" colspan="4" bgcolor="#EBEBEB"> 
        <iframe id='NewsContent' src='../../Editer/NewsEditer.asp' frameborder=0 scrolling=no width='100%' height='350'></iframe></td>
    </tr>
    <tr bgcolor="#F2F2F2"> 
      <td height="31"> 
        <div align="center">����ʱ�䡡</div></td>
      <td colspan="3"> 
        <input name="Addtime" type="text" id="Addtime" value="<% = RsUserModifyObj("Addtime")%>">
        ,����ȷ��дʱ���ʽ��</td>
    </tr>
    <tr valign="middle" bgcolor="#F2F2F2"> 
      <td> 
        <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����</div></td>
      <td> 
        <input type="radio" name="isLock" value="1" <%If RsUserModifyObj("isLock") = 1 then Response.Write("checked")%>>
        �� 
        <input name="isLock" type="radio" value="0" <%If RsUserModifyObj("isLock") = 0 or Request("Lock") = "" then Response.Write("checked")%>>
        ��</td>
    </tr>
  </table>
</form>
</body>
</html>
<%
RsUserModifyObj.Close
Set RsUserModifyObj = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.UserAddSForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.UserAddSForm.Content.value=='') document.UserAddSForm.Content.value=TempContentArray[i];
			else document.UserAddSForm.Content.value=document.UserAddSForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.UserAddSForm.submit();
}
</script>
<script>
function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.UserAddSForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.UserAddSForm.Content.value=='') document.UserAddSForm.Content.value=TempContentArray[i];
			else document.UserAddSForm.Content.value=document.UserAddSForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.UserAddSForm.submit();
}
</script>
