<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
Dim SpecialID,Types,NewsID,NewsObj,SpecialIDStr
if Request("Types")<>"" and Request("NewsID")<>"" and Request("SpecialID")<>"" then
	if Not JudgePopedomTF(Session("Name"),"P020401") then Call ReturnError()
	Types = Request("Types")
	NewsID = Request("NewsID")
	SpecialID = Request("SpecialID")
elseif Request("SpecialID") <> "" and Request("Types") = "" then
	if Not JudgePopedomTF(Session("Name"),"P020300") then Call ReturnError()
	SpecialID = Cstr(Request("SpecialID"))
	Types = ""
else
	Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
	Response.End
end if 
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>频道/专题删除</title>
</head>
<body leftmargin="0" topmargin="0">
<%If Types = "" then %>
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="26%"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="74%">您确定要删除所选频道/专题?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input type="hidden" name="action" value="trues">
          <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    </tr>
</table>
</form>
<%else%>
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td colspan="2"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="74%" height="2">您确定要从频道/专题内删除所选新闻?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input name="newsaction" type="hidden" id="newsaction" value="trues">
          <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    </tr>
</table>
</form>
<%end if%>
</body>
</html>
<%
if request.Form("action")="trues" then
	if Not JudgePopedomTF(Session("Name"),"P020300") then Call ReturnError()
	Dim DspArray,Dsp_i
	DspArray = Array("")
	DspArray = Split(SpecialID,",")
	For Dsp_i = 0 to UBound(DspArray)
		Conn.Execute("delete from FS_Special where SpecialID='"&DspArray(Dsp_i)&"'")
		'----------修改新闻表的SpcialID字段--------
		Dim ModifyNewsObj,ModSpecialIDStr
		Set ModifyNewsObj = Conn.Execute("Select SpecialID,NewsID from FS_News where SpecialID like '%"&DspArray(Dsp_i)&"%'")
		while not ModifyNewsObj.eof
			ModSpecialIDStr = ModifyNewsObj("SpecialID")
			ModSpecialIDStr = Replace(ModSpecialIDStr,DspArray(Dsp_i),"")
			ModSpecialIDStr = Replace(ModSpecialIDStr,",,",",")
			If left(ModSpecialIDStr,1)="," then
				ModSpecialIDStr = right(ModSpecialIDStr,len(ModSpecialIDStr)-1)
			end if
			If right(ModSpecialIDStr,1)="," then
				ModSpecialIDStr = left(ModSpecialIDStr,len(ModSpecialIDStr)-1)
			end if
			Conn.Execute("Update FS_News set SpecialID='"&ModSpecialIDStr&"' where NewsID='"&ModifyNewsObj("NewsID")&"'")
			ModifyNewsObj.MoveNext
		wend
		ModifyNewsObj.Close
		Set ModifyNewsObj = Nothing
	Next
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
	
if Request.Form("newsaction")="trues" then
	if Not JudgePopedomTF(Session("Name"),"P020401") then Call ReturnError()
	Dim NewsArray,DSN_i
	NewsArray = Array("")
	NewsArray = Split(NewsID,"***")
	For DSN_i = 0 to UBound(NewsArray)
		Set NewsObj = Conn.Execute("Select SpecialID from FS_News where NewsID='" & NewsArray(DSN_i) & "'")
		SpecialIDStr = NewsObj("SpecialID")
		SpecialIDStr = Replace(SpecialIDStr,SpecialID,"")
		SpecialIDStr = Replace(SpecialIDStr,",,",",")
		If left(SpecialIDStr,1)="," then
			SpecialIDStr = right(SpecialIDStr,len(SpecialIDStr)-1)
		end if
		If right(SpecialIDStr,1)="," then
			SpecialIDStr = left(SpecialIDStr,len(SpecialIDStr)-1)
		end if
		Conn.Execute("Update FS_News set SpecialID='"&SpecialIDStr&"' where NewsID='"&NewsArray(DSN_i)&"'")
		NewsObj.close
		Set NewsObj = Nothing
	Next
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
Set Conn = Nothing
%>