<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P080502") then Call ReturnError1()
	Dim RsOCObj,TempFlag
	Set RsOCObj = Conn.Execute("Select * from FS_WebInfo")
	If RsOCObj.eof then
		TempFlag = false
	Else
		TempFlag = true
	End If
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��վά��</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<form action="" method="post" name="VOForm">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="28" class="ButtonListLeft"> <div align="center"><strong>��վ��Ϣά��</strong></div></td>
    </tr>
  </table>
  <br>
  <table width="75%"  border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="e6e6e6" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="24%">&nbsp;&nbsp;&nbsp;&nbsp;��վ����</td>
      <td width="76%"> 
        <input name="WebName" type="text" id="WebName" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebName")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��վ��ַ</td>
      <td> 
        <input name="WebUrl" type="text" id="WebUrl" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebUrl")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;����Ա</td>
      <td> 
        <input name="WebAdmin" type="text" id="WebAdmin" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebAdmin")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��վ����</td>
      <td> 
        <input name="WebEmail" type="text" id="WebEmail" style="width:90%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebEmail")) end if%>"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��ʼͳ��ʱ��</td>
      <td> 
        <input name="WebCountTime" type="text" readonly id="WebCountTime" style="width:71%" value="<%If TempFlag = true then Response.Write(RsOCObj("WebCountTime")) end if%>">
      <input type="button" name="dfgdf" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.VOForm.WebCountTime);"></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;��վ����</td>
      <td> 
        <textarea name="WebIntro" id="WebIntro" style="width:90%"><%If TempFlag = true then Response.Write(RsOCObj("WebIntro")) end if%></textarea></td>
  </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> 
        <div align="center">
      <input type="submit" name="Submit" value=" ȷ �� ">&nbsp;&nbsp;
      <input name="action" type="hidden" id="action" value="trues">
      <input type="reset" name="Submit" value=" �� ԭ ">&nbsp;&nbsp;
      <input type="button" name="Submit" value=" ȡ �� " onclick="history.back();">
    </div></td>
    </tr>
</table>
</form>
</body>
</html>
<%
	If Request.Form("action") = "trues" then
		Dim VOModObj,VoModSql
		Set VOModObj = Server.CreateObject(G_FS_RS)
		VoModSql = "Select * from FS_WebInfo order by ID asc"
		VOModObj.Open VoModSql,Conn,3,3
		If TempFlag = false then
		VOModObj.AddNew
		End If
		VOModObj("WebName") = Replace(Replace(Request.Form("WebName"),"""",""),"'","")
		VOModObj("WebUrl") = Request.Form("WebUrl")
		VOModObj("WebIntro") = Request.Form("WebIntro")
		VOModObj("WebEmail") = Request.Form("WebEmail")
		VOModObj("WebAdmin") = Request.Form("WebAdmin")
		VOModObj("WebCountTime") = Request.Form("WebCountTime")
		VOModObj.Update
		VOModObj.Close
		Set VOModObj = Nothing
		Response.Write("<script>alert(""��վ��Ϣά���ɹ�"");history.back();</script>")
	End If
Conn.Close
Set Conn = Nothing
%>