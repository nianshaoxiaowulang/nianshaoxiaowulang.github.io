<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="Cls_Ads.asp" -->
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
Dim AdsTempLocation,TempTypes,TipWord,FileObj,ADialog
AdsTempLocation = Request("Location")
TempTypes = Request("Types")
if AdsTempLocation = "" or TempTypes = "" then
	%>
		<script>alert('�������ݴ���');history.back();</script>
	<%
else
	Select Case TempTypes
		Case "Dell"
			if Not JudgePopedomTF(Session("Name"),"P070203") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "ɾ��"
		Case "Stop"
			if Not JudgePopedomTF(Session("Name"),"P070204") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "��ͣ"
		Case "Star"
			if Not JudgePopedomTF(Session("Name"),"P070205") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "����"
		Case Else
			Response.Write("<script>alert('�������ݴ���');history.back();</script>")
			Response.End
	End select
	Dim ALeftPicPath,TempStateFlagState,ACycleLocation,AType,ACycleTF,APicHeight,AExplain,ARightPicPath,APicWidth
	Dim SelectLocation
	if AdsTempLocation <> "" and TempTypes="Stop" and request.form("action")="trues" then
	   Conn.Execute("update FS_Ads set State=2 where Location in (" & Replace(AdsTempLocation,"***",",") & ")")
	   ADialog = "�����ͣ�ɹ�"
	end if
	if AdsTempLocation <> "" and TempTypes="Star" and Request.form("action")="trues" then
	   Conn.Execute("update FS_Ads set State=1 where Location in (" & Replace(AdsTempLocation,"***",",") & ")")
	   ADialog = "��漤��ɹ�"
	end if
	Set TempStateFlag = Conn.Execute("Select * from FS_Ads where Location in (" & Replace(AdsTempLocation,"***",",") & ")")
	do while Not TempStateFlag.Eof
		SelectLocation = TempStateFlag("Location")
		ALeftPicPath = TempStateFlag("LeftPicPath")
		APicWidth = TempStateFlag("PicWidth")
		APicHeight = TempStateFlag("PicHeight")
		ARightPicPath = TempStateFlag("RightPicPath")
		AExplain = TempStateFlag("Explain")
		AType = TempStateFlag("Type")
		ACycleTF = TempStateFlag("CycleTF")
		ACycleLocation = TempStateFlag("CycleLocation")
		TempStateFlagState = TempStateFlag("State")
		if TempTypes <> "Dell" then	  
			Select Case AType
				Case "1" call ShowAds(SelectLocation)
				Case "2" call NewWindow(SelectLocation)
				Case "3" call OpenWindow(SelectLocation)
				Case "4" call FilterAway(SelectLocation)
				Case "5" call DialogBox(SelectLocation)
				Case "6" call ClarityBox(SelectLocation)
				Case "7" call RightBottom(SelectLocation)
				Case "8" call DriftBox(SelectLocation)
				Case "9" call LeftBottom(SelectLocation)
				Case "10" call Couplet(SelectLocation)
			  End Select
		 end if
		if SelectLocation <> "" and TempTypes="Dell" and Request.form("action") = "trues" then
			Conn.Execute("update FS_Ads Set State=0 where Location=" & SelectLocation & "")
		end if
		if ACycleTF = "1" or ACycleLocation<>"0" then
			if TempTypes<>"Dell" and AType <> "11" then	call Cycle(SelectLocation,ACycleLocation)
		end if
		if AdsTempLocation <> "" and TempTypes = "Dell" and Request.Form("action") = "trues" then
			Conn.Execute("delete from FS_Ads where Location=" & SelectLocation & "")
			Conn.Execute("delete from FS_AdsVisitList where AdsLocation=" & SelectLocation & "")
			Set FileObj = Server.CreateObject(G_FS_FSO)
			if FileObj.FileExists(Server.MapPath("\") & "\JS\AdsJs\" & SelectLocation & ".js") = True then
				FileObj.DeleteFile (Server.MapPath("\") & "\JS\AdsJs\" & SelectLocation & ".js")
			end if
			ADialog = "���ɾ���ɹ�"
		end if
		TempStateFlag.MoveNext
	Loop
	Set TempStateFlag = Nothing
end if
if ADialog <> "" then
	Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
  <form method="post">
	<tr> 
      <td width="5%">&nbsp;</td>
      <td width="26%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="64%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>��ȷ��Ҫ<%=TipWord%>��</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input type="hidden" name="action" value="trues">
          <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� " >
        </div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </form>
</table>
</body>
</html>
