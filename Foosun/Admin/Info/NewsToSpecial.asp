<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010507") then Call ReturnError()
Dim NewsIDStr,RsNewsObj
If Request("NewsID")<>"" then
	NewsIDStr = Cstr(Request("NewsID"))
else
	Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	Response.End
end if
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加新闻到专题</title>
</head>
<body>
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" name="ToWordJsForm" method="post">
    <tr>  
      <td width="7%" height="5">&nbsp;</td>
      <td width="16%" height="5">&nbsp;</td>
      <td width="77%" height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>专题名称</td>
      <td><select name="SpecialID" id="JSName" style="width:90%">
          <option value=""> </option>
          <%
	    Dim SpecialObj
		Set SpecialObj = Conn.Execute("Select SpecialID,CName from FS_Special order by ID desc")
	    While Not SpecialObj.eof 
	  %>
          <option value="<%=SpecialObj("SpecialID")%>"><%=SpecialObj("CName")%></option>
          <%
		SpecialObj.MoveNext
		Wend
	    SpecialObj.Close
		Set SpecialObj = Nothing
	  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit22" value=" 确 定 ">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input name="action" type="hidden" id="action2" value="trues">
          <input type="button" name="Submit32" value=" 取 消 " onClick="window.close();">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
If Request.Form("action") = "trues" then
	Dim ToSpecialID,SpecialIDStr,NewsIDArray,R_i,ErrorStr
	ToSpecialID = Request.Form("SpecialID")
	If ToSpecialID = "" or isnull(ToSpecialID) then
		Response.Write("<script>alert(""请选择专题"");</script>")
		Response.End
	End If
	NewsIDArray = Array("")
	NewsIDArray = Split(NewsIDStr,"***")
	For R_i = 0 to UBound(NewsIDArray)
		Set RsNewsObj = Server.CreateObject(G_FS_RS)
		RsNewsObj.Open "Select SpecialID,NewsID,Title from FS_News where HeadNewsTF=0 and NewsID='" & NewsIDArray(R_i) & "'",Conn,1,1
		If Not RsNewsObj.eof then
			if IsNull(RsNewsObj("SpecialID")) then
				Conn.Execute("Update FS_News set SpecialID='" & ToSpecialID & "' where NewsID='" & NewsIDArray(R_i) & "' and DelTF=0 and AuditTF=1")
			else
				if RsNewsObj("SpecialID") = "" then
					Conn.Execute("Update FS_News set SpecialID='" & ToSpecialID & "' where NewsID='" & NewsIDArray(R_i) & "'")
				elseIf Instr(1,RsNewsObj("SpecialID"),ToSpecialID) = 0 then
					SpecialIDStr = RsNewsObj("SpecialID") & "," & ToSpecialID
					Conn.Execute("Update FS_News set SpecialID='" & SpecialIDStr & "' where NewsID='" & NewsIDArray(R_i) & "'")
				else
					if ErrorStr = "" then
						ErrorStr = RsNewsObj("Title")
					else
						ErrorStr = ErrorStr & "|" & RsNewsObj("Title")
					end if
				End If
			End If
		end if
		RsNewsObj.Close
		Set RsNewsObj = Nothing
	Next
	if ErrorStr = "" then
		Response.Write("<script>alert(""所选新闻已经成功加入专题"");dialogArguments.location.reload();window.close();</script>")
	else
		Response.Write("<script>alert('专题中已经存在:" & ErrorStr & "');dialogArguments.location.reload();window.close();</script>")
	end if
End If
%>