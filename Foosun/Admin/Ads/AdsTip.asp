<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="Cls_Ads.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
		<script>alert('参数传递错误');history.back();</script>
	<%
else
	Select Case TempTypes
		Case "Dell"
			if Not JudgePopedomTF(Session("Name"),"P070203") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "删除"
		Case "Stop"
			if Not JudgePopedomTF(Session("Name"),"P070204") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "暂停"
		Case "Star"
			if Not JudgePopedomTF(Session("Name"),"P070205") then Call ReturnError()
			Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
			TipWord = "激活"
		Case Else
			Response.Write("<script>alert('参数传递错误');history.back();</script>")
			Response.End
	End select
	Dim ALeftPicPath,TempStateFlagState,ACycleLocation,AType,ACycleTF,APicHeight,AExplain,ARightPicPath,APicWidth
	Dim SelectLocation
	if AdsTempLocation <> "" and TempTypes="Stop" and request.form("action")="trues" then
	   Conn.Execute("update FS_Ads set State=2 where Location in (" & Replace(AdsTempLocation,"***",",") & ")")
	   ADialog = "广告暂停成功"
	end if
	if AdsTempLocation <> "" and TempTypes="Star" and Request.form("action")="trues" then
	   Conn.Execute("update FS_Ads set State=1 where Location in (" & Replace(AdsTempLocation,"***",",") & ")")
	   ADialog = "广告激活成功"
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
			ADialog = "广告删除成功"
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
<title>广告管理</title>
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
      <td>您确定要<%=TipWord%>吗？</td>
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
          <input type="submit" name="Submit" value=" 确 定 ">
          <input type="hidden" name="action" value="trues">
          <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 " >
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
