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
if Not JudgePopedomTF(Session("Name"),"P990100") then Call ReturnError()
Dim Path,FileName,Action
Path = Request("Path")
FileName = Request("FileName")
if FileName = "" then FileName = "Untitled.htm"
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
      <div align="center">文件名称 
          <input type="text" name="FileName" onFocus="CheckErrorChar();" onBlur="CheckErrorChar();" value="<% = FileName %>">
      </div></td>
  </tr>
  <tr> 
    <td height="26">
<div align="center">
        <input type="submit" name="Submit" value=" 确 定 ">
          <input name="Action" type="hidden" id="Action" value="Submit">
          <input name="Path" type="hidden" value="<% = Path %>" id="ParentPath">
        </div></td>
    <td height="26">
<div align="center">
        <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
      </div></td>
  </tr>
 </form>
</table>
</body>
</html>
<script language="JavaScript">
document.Form.FileName.focus();
document.Form.FileName.select();
function CheckErrorChar()
{
	var TempStr=document.Form.FileName.value,AlertStr='';
	var ErrorCharArray=new Array('\'','"','*');
	var re=null;
	for (var i=0;i<ErrorCharArray.length;i++)
	{
		if (TempStr.indexOf(ErrorCharArray[i])!=-1)
		{
			AlertStr+=ErrorCharArray[i];
			re=new RegExp('['+ErrorCharArray[i]+']?','ig');
			document.Form.FileName.value=document.Form.FileName.value.replace(re,'');
		}
	}
	if (AlertStr!='') alert('发现非法字符'+AlertStr);
}
</script>
<%
if Action = "Submit" then
	Dim FsoObj,PhysicalPath,FileObj,WriteStr,ResponseStr
	WriteStr = WriteStr & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & Chr(13) & Chr(10) & "<html>" & Chr(13) & Chr(10)
	WriteStr = WriteStr & "<head>" & Chr(13) & Chr(10) & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & Chr(13) & Chr(10)
	WriteStr = WriteStr & "<title>无标题文档</title>" & Chr(13) & Chr(10) & "</head>" & Chr(13) & Chr(10) & "<body>" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "</body>" & Chr(13) & Chr(10) & "</html>"
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	if Path <> "" And FileName <> "" then
		PhysicalPath = Server.MapPath(Path)
		if FsoObj.FolderExists(PhysicalPath) = True then
			PhysicalPath = Server.MapPath(Path) & "\" & FileName
			if FsoObj.FileExists(PhysicalPath) = False then
				Set FileObj = FsoObj.CreateTextFile(PhysicalPath)
				FileObj.WriteLine(WriteStr)
				Set FileObj = Nothing
			else
				ResponseStr = "文件已经存在"
			end if
		else
			ResponseStr = Path & "目录不存在"
		end if
	else
		ResponseStr = "参数传递错误"
	end if
	Set FsoObj = Nothing
	if ResponseStr <> "" then
		%>
			<script language="JavaScript">alert('<% = ResponseStr %>');dialogArguments.location.reload();window.close();</script>
		<%
	else
		%>
			<script language="JavaScript">dialogArguments.location.reload();window.close();</script>
		<%
	end if
end if
Set Conn = Nothing
%>