<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError()
Dim NewsID,DownLoadID,Action,ProductID
NewsID = Request("NewsID")
DownLoadID = Request("DownLoadID")
ProductID=Request("ProductID")
Action = Request("Action")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="8" cellpadding="0">
  <form name="DelForm" method="post" action="">
    <tr> 
      <td width="21%"> <div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="79%" colspan="2"> ȷ��Ҫ������? 
        <input name="Action" type="hidden" id="Action" value="Submit"> 
        <input name="NewsID" type="hidden" id="NewsID" value="<% = NewsID %>"> 
        <input name="DownLoadID" type="hidden" id="DownLoadID" value="<% = DownLoadID %>"></td>
    </tr>
    <tr> 
      <td colspan="3"> <div align="center"> 
          <input name="Submitsadf" type="submit" id="Submitsadf" value=" �� �� ">
          
          <input type="button" onClick="window.close();" name="Submit3" value=" ȡ �� ">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
if Action = "Submit" then
	Dim RsRefreshObj,RefreshNewsNum,RefreshDownLoadNum,RefreshProductNum,AlertStr
	RefreshNewsNum = 0
	RefreshDownLoadNum = 0
	if NewsID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010510") then Call ReturnError()
		NewsID = Replace(NewsID,"***","','")
		Set RsRefreshObj = Conn.Execute("Select * from FS_News where AuditTF=1 and NewsID in ('" & NewsID & "')")
		do while Not RsRefreshObj.Eof
			RefreshNews RsRefreshObj
			RefreshNewsNum = RefreshNewsNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
			AlertStr = "����" & RefreshNewsNum & "������"
	end if
	if DownLoadID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010706") then Call ReturnError()
		DownLoadID = Replace(DownLoadID,"***","','")
		Set RsRefreshObj = Conn.Execute("Select * from FS_DownLoad where AuditTF=1 and DownLoadID in ('" & DownLoadID & "')")
		do while Not RsRefreshObj.Eof
			RefreshDownLoad RsRefreshObj
			RefreshDownLoadNum = RefreshDownLoadNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
		AlertStr = "����" & RefreshDownLoadNum & "������"
	end if
	if ProductID <> "" then
		if Not JudgePopedomTF(Session("Name"),"P010806") then Call ReturnError()
		ProductID = Replace(ProductID,"***",",")
		Set RsRefreshObj = Conn.Execute("Select * from FS_Shop_Products where id in (" & ProductID & ")")

		do while Not RsRefreshObj.Eof
			RefreshOneMallProduct RsRefreshObj
			RefreshProductNum = RefreshProductNum + 1
			RsRefreshObj.MoveNext
		Loop
		Set RsRefreshObj = Nothing
		AlertStr = "����" & RefreshProductNum & "����Ʒ"
	end if

	Response.Write("<script language=""JavaScript"">alert('" & AlertStr & "');window.close();</script>")
end if
%>