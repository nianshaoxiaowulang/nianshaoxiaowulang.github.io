<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
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
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
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
<title>����JSɾ��</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
<form action="" name="JSDellForm" method="post">
  <tr> 
    <td><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td colspan="2">��ȷ��Ҫɾ����JS?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input type="hidden" name="action" value="trues">
          <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
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