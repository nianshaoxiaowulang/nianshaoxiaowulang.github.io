<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P990200") then Call ReturnError()

Dim Path,AddPath,Action,sRootDir
if SysRootDir<>"" then sRootDir="/" & SysRootDir else sRootDir=""
Path = Request("Path")
AddPath = Request("AddPath")
Action = Request("Action")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<link href="../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <form action="" method="post" name="Form">
  <tr> 
    <td height="36" colspan="2"> 
      <div align="center">��Ŀ����
        <input type="text" onFocus="CheckErrorChar();" onBlur="CheckErrorChar();" name="AddPath" value="<% = AddPath %>">
      </div></td>
  </tr>
  <tr> 
    <td height="26">
<div align="center">
        <input type="submit" name="Submit" value=" ȷ �� ">
          <input name="Action" type="hidden" id="Action" value="Submit">
          <input name="ParentPath" type="hidden" value="<% = Path %>" id="ParentPath">
        </div></td>
    <td height="26">
<div align="center">
        <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
      </div></td>
  </tr>
 </form>
</table>
</body>
</html>
<script language="JavaScript">
document.Form.AddPath.focus();
function CheckErrorChar()
{
	var TempStr=document.Form.AddPath.value,AlertStr='';
	var ErrorCharArray=new Array('\'','"','*');
	var re=null;
	for (var i=0;i<ErrorCharArray.length;i++)
	{
		if (TempStr.indexOf(ErrorCharArray[i])!=-1)
		{
			AlertStr+=ErrorCharArray[i];
			re=new RegExp('['+ErrorCharArray[i]+']?','ig');
			document.Form.AddPath.value=document.Form.AddPath.value.replace(re,'');
		}
	}
	if (AlertStr!='') alert('���ַǷ��ַ�'+AlertStr);
}
</script>
<%
if Action = "Submit" then
	Dim FsoObj,PhysicalPath,ResponseStr
	if Path <> "" and AddPath <> "" then
		if Path = sRootDir &"/" then
			PhysicalPath = Server.MapPath(Path & AddPath)
		else
			PhysicalPath = Server.MapPath(Path & "/" & AddPath)
		end if
		Set FsoObj = Server.CreateObject(G_FS_FSO)
		if FsoObj.FolderExists(PhysicalPath) = True then
			ResponseStr = "Ŀ¼�Ѿ�����"
		else
			FsoObj.CreateFolder PhysicalPath
		end if
		Set FsoObj = Nothing
	end if
	if ResponseStr <> "" then
		%>
			<script language="JavaScript">alert('<% = ResponseStr %>');document.Form.AddPath.select();</script>
		<%
	else
		%>
			<script language="JavaScript">dialogArguments.location.reload();window.close();</script>
		<%
	end if
end if
%>