<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'�ж�Ȩ��
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080202") then Call ReturnError1()
'�ж�Ȩ�޽���
Dim RuleID
RuleID = Request("RuleID")
if Request.Form("Result")="Edit" then
    Dim Sql,RsEditObj
	if RuleID <> "" then
		Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
		Sql = "Select * from FS_Rule where id=" & RuleID
		RsEditObj.Open Sql,CollectConn,1,3
		if RsEditObj.Eof then
			Response.Write"<script>alert(""û���޸Ĺ���"");location.href=""javascript:history.back()"";</script>"
			Response.End
		end if
		RsEditObj("RuleName") = NoCSSHackAdmin(Request.Form("RuleName"),"��������")
		RsEditObj("SiteId") = Request.Form("SiteId")
		Dim KeywordSetting
		If InStr(Request.Form("KeywordSetting"),"[�����ַ���]")<>0 then
			KeywordSetting = Split(Request.Form("KeywordSetting"),"[�����ַ���]",-1,1)
			RsEditObj("HeadSeting") = KeywordSetting(0)
			RsEditObj("FootSeting") = KeywordSetting(1)
		End If
		RsEditObj("ReContent") = Request.Form("ReContent")
		RsEditObj.UpDate
		RsEditObj.Close
		Set RsEditObj = Nothing
	else
		Response.Write"<script>alert(""�������ݴ���"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
	Response.Redirect("Rule.asp")
	Response.End
end if

Dim RsRuleObj
if RuleID <> "" then
	Set RsRuleObj = Server.CreateObject ("ADODB.RecordSet")
	Sql = "Select * from FS_Rule where id=" & RuleID
	RsRuleObj.Open Sql,CollectConn,1,3
	if RsRuleObj.Eof then
		Response.Write"<script>alert(""û���޸Ĺ���"");location.href=""javascript:history.back()"";</script>"
		Response.End
	end if
else
	Response.Write"<script>alert(""�������ݴ���"");location.href=""javascript:history.back()"";</script>"
	Response.End
end if
	
Dim SiteList,RsSiteObj
Set RsSiteObj = Server.CreateObject("Adodb.RecordSet")
RsSiteObj.Source = "Select ID,SiteName from FS_Site order by id desc"
RsSiteObj.open RsSiteObj.Source,CollectConn,1,3
do while Not RsSiteObj.Eof
	if Clng(RsRuleObj("SiteID")) = Clng(RsSiteObj("ID")) then
		SiteList = SiteList & "<option selected value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	else
		SiteList = SiteList & "<option value=" & RsSiteObj("ID") & "" & ">" & RsSiteObj("SiteName") & "</option><br>"
	end if
	RsSiteObj.MoveNext	
loop
RsSiteObj.Close
Set RsSiteObj = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�Զ����Ųɼ���վ������</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="form1" id="form1" method="post" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="document.form1.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="����" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result4" value="Edit">
          <input name="id" type="hidden" id="id2" value="<% = RuleID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr> 
      <td width="100"> <div align="center">��������</div></td>
      <td> <input name="RuleName" style="width:100%;" type="text" id="RuleName" value="<% = RsRuleObj("RuleName") %>"> 
        <div align="right"></div></td>
    </tr>
    <tr> 
      <td><div align="center">Ӧ�õ�</div></td>
      <td><select name="SiteId" style="width:100%;" id="SiteId">
          <% =SiteList %>
        </select></td>
    </tr>
    <tr> 
      <td> <div align="center">�����ַ���</div></td>
      <td> &nbsp;&nbsp;�������� <span onClick="if(document.Form1.KeywordSetting.rows>2)document.Form1.KeywordSetting.rows-=1" style='cursor:hand'><b>��С</b></span> 
        <span onClick="document.Form1.KeywordSetting.rows+=1" style='cursor:hand'><b>����</b></span> 
        &nbsp;&nbsp;���ñ�ǩ:<font onClick="addTag('[�����ַ���]')" style="CURSOR: hand"><b>[�����ַ���]</b></font> 
        &nbsp;&nbsp;&nbsp;<font onClick="addTag('[����]')" style="CURSOR: hand"><b>[����]</b></font><br>
        <br>
	  <textarea name="KeywordSetting"  onfocus="getActiveText(this)" onclick="getActiveText(this)"  onchange="getActiveText(this)" rows="5" id="textarea2" style="width:100%;"><% = RsRuleObj("HeadSeting") %>[�����ַ���]<% = RsRuleObj("FootSeting") %></textarea> 
	  </td>
    </tr>
    <tr> 
      <td> <div align="center"> 
          �滻Ϊ</div></td>
      <td colspan="3"><textarea style="width:100%;" name="ReContent" cols="30" rows="5" id="ReContent"><% = RsRuleObj("ReContent") %></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set CollectConn = Nothing
Set RsRuleObj = Nothing
%>

<script language="javaScript">

currObj = "uuuu";
function getActiveText(obj)
{
	currObj = obj;
}

function addTag(code)
{
	addText(code);
}

function addText(ibTag)
{
	var isClose = false;
	var obj_ta = currObj;
//alert("ok");
	if (obj_ta.isTextEdit)
	{
	//alert("nooooo");
		obj_ta.focus();
		var sel = document.selection;
		var rng = sel.createRange();
		rng.colapse;

		if((sel.type == "Text" || sel.type == "None") && rng != null)
		{
			rng.text = ibTag;
		}

		obj_ta.focus();

		return isClose;
	}
	else
		return false;
}	

</script>
