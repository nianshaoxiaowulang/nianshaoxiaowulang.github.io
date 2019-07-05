<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System v3.1 
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071,66252421
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060300") then Call ReturnError()
dim JSID,JSDellObj,JSEName,FileObj
if Request("JSID")<>"" then
	JSID = Request("JSID")
else
	Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
response.end
end if 
Dim DelJSSysRootDir
if SysRootDir = "" then
	DelJSSysRootDir = ""
else
	DelJSSysRootDir = "/" & SysRootDir
end if
		
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自由JS删除</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<form action="" name="JSDellForm" method="post">
  <tr> 
    <td><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td colspan="2">您确定要删除此JS?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" 确 定 ">
          <input type="hidden" name="action" value="trues">
          <input type="button" name="Submit2" value=" 取 消 " onClick="window.close();">
      </div></td>
    </tr>
</form>
</table>
</body>
</html>
<%
if request.Form("action")="trues" then
  Dim DjArray,Dj_i,TemmpObjj
  DjArray = Array("")
  DjArray = Split(JSID,"***")
  For Dj_i = 0 to UBound(DjArray)
  Set TemmpObjj = Conn.Execute("Select ID,EName from FS_FreeJS where ID="&DjArray(Dj_i)&"")
  If Not TemmpObjj.eof then
		 Conn.Execute("delete from FS_FreeJsFile where JSName='"&TemmpObjj("EName")&"'")
		Set FileObj = Server.CreateObject(G_FS_FSO)
		if FileObj.FileExists(Server.MapPath(DelJSSysRootDir&"\JS\FreeJs")&"\"& TemmpObjj("EName") &".js") then
			FileObj.DeleteFile (Server.MapPath(DelJSSysRootDir&"\JS\FreeJs")&"\"& TemmpObjj("EName") &".js")
		end if 
		 Conn.Execute("delete from FS_FreeJS where ID="&TemmpObjj("ID")&"")
   end if
   TemmpObjj.Close
   Set TemmpObjj = Nothing
   Next
	response.write("<script>dialogArguments.location.reload();window.close();</script>")
	response.end
end if
%>