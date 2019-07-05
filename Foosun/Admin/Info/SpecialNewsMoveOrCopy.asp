<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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

%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P020402") then Call ReturnError()
Dim MoveOrCopyClassPara
If Request("MoveOrCopyClassPara") = "" then
	Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
	Response.End
Else
	MoveOrCopyClassPara = Request("MoveOrCopyClassPara")
	Dim MoveTF,MoveTFText,ResultString,TempSpecialIDStr,IntentEName,TempNewsObj,IntentCName,SourceCName,SourceEName,DoNews,IntentObj,SourceObj,IntentFileSql,IntentFileObj,SourceFileObj
	ResultString = ""
	MoveTF = GetParaValue(MoveOrCopyClassPara,"MoveTF")
	If MoveTF="true" then
		MoveTFText = "移动"
	Else
		MoveTFText = "复制"
	End If
	IntentEName = GetParaValue(MoveOrCopyClassPara,"ObjectClass") '目的JS
	SourceEName = GetParaValue(MoveOrCopyClassPara,"SourceClass") '源JS
	DoNews = GetParaValue(MoveOrCopyClassPara,"SourceNews") '操作新闻 
	Set IntentObj = Conn.Execute("Select CName from FS_Special where SpecialID='"&IntentEName&"'")
	If IntentObj.eof then
		Set IntentObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""目的专题已经不存在,请刷新后再进行操作"");dialogArguments.location.reload();window.close();</script>")
		Response.End
	Else
		IntentCName = IntentObj("CName")
	End if
	Set SourceObj = Conn.Execute("Select CName from FS_Special where SpecialID='"&SourceEName&"'")
	If SourceObj.eof then
		Set SourceObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""源专题已经不存在,请刷新后再进行操作"");dialogArguments.location.reload();window.close();</script>")
		Response.End
	Else
		SourceCName = SourceObj("CName")
	End if
	If Cstr(IntentEName) = Cstr(SourceEName) then
		Set IntentObj = Nothing
		Set SourceObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""源专题和目的专题不能相同"");window.close();</script>")
		Response.End
	End if

	If Request.Form("action") = "trues" then
		Dim DoNewsArray,SpMove_i
		DoNewsArray = Array("")
		DoNewsArray = Split(DoNews,"***")
		For SpMove_i = 0 to UBound(DoNewsArray)
			Set TempNewsObj = Conn.Execute("Select SpecialID from FS_News where NewsID='"&DoNewsArray(SpMove_i)&"'")
			If Not TempNewsObj.eof then
				If MoveTF = "true" then '移动新闻
					TempSpecialIDStr = Replace(TempNewsObj("SpecialID"),SourceEName,IntentEName) 
					Conn.Execute("Update FS_News set SpecialID='"&TempSpecialIDStr&"' where NewsID='"&Cstr(DoNewsArray(SpMove_i))&"'")
					ResultString = "新闻移动成功"
				Else '复制新闻
					TempSpecialIDStr = TempNewsObj("SpecialID")&","&IntentEName
					Conn.Execute("Update FS_News set SpecialID='"&TempSpecialIDStr&"' where NewsID='"&Cstr(DoNewsArray(SpMove_i))&"'")
					ResultString = "新闻拷贝成功"
				End if
			End If
			TempNewsObj.Close
			Set TempNewsObj = Nothing
		Next  
	End If
	If ResultString <> "" then
		Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		Response.End
	End If
End If
  
Function GetParaValue(ParaStr,ParaName)
	Dim BeginIndex,EndIndex
	BeginIndex = InStr(ParaStr,ParaName)+Len(ParaName)+1
	EndIndex = InStr(BeginIndex,ParaStr,",")
	GetParaValue = Mid(ParaStr,BeginIndex,EndIndex-BeginIndex)
End Function
%>

<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>移动或复制专题新闻</title>
</head>
<body topmargin="0" leftmargin="0" >
<form action="" method="post" name="MoveForm">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="12%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="75%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>您确定要将此新闻从<font color="#FF0000"><%=SourceCName%></font><font color="#0000FF"><%=MoveTFText%></font>到<font color=red><%=IntentCName%></font>?</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" 确 定 ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
    <td height="10" colspan="2">&nbsp;</td>
    <td height="10">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
