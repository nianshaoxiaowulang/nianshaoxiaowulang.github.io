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
<%
Dim Path,Result,OldPathName,NewPathName
OldPathName = Request("OldPathName")
NewPathName = Request.Form("NewPathName")
Path = Request("Path")
Result = Request.Form("Result")
if NewPathName = "" then
	NewPathName = OldPathName
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件夹重命名</title>
</head>
<link href="../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="10">
<div align="center">
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <form name="ReNameForm" action="" method="post">
      <tr>
        <td height="20">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td width="20%"> <div align="center">文件名</div></td>
        <td> <div align="left"> 
            <input style="width:100%;" value="<% = NewPathName %>" type="text" name="NewPathName">
          </div></td>
      </tr>
      <tr> 
        <td height="50" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td><div align="center"> 
                  <input type="submit" name="Submit" value=" 确 定 ">
                  <input name="Result" type="hidden" id="Result" value="Submit">
                  <input type="hidden" name="OldPathName" value="<% = OldPathName %>">
                  <input type="hidden" name="Path" value="<% = Path %>">
                  <input type="hidden" name="hiddenField2">
                </div></td>
              <td><div align="center"> 
                  <input type="button" onClick="dialogArguments.location.reload();window.close();" name="Submit2" value=" 取 消 ">
                </div></td>
            </tr>
          </table></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
document.all.NewPathName.select();
</script>
<%
if Result = "Submit" then
	Dim FsoObj,PhysicalPath,FileObj
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	if Path <> "" then
		PhysicalPath = Server.MapPath(Path)
		if (NewPathName <> "") or (OldPathName <> "") then
			PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
			if FsoObj.FolderExists(PhysicalPath) = True then
				PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
				if FsoObj.FolderExists(PhysicalPath) = False then
					Set FileObj = FsoObj.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
					FileObj.Name = NewPathName
				else
%>
<script language="JavaScript">
	alert('文件夹已经存在，重命名失败');
	dialogArguments.location.reload();
	//close();
</script>
<%
				end if
			else
%>
<script language="JavaScript">
	alert('重命名文件夹不存在，重命名失败');
	dialogArguments.location.reload();
	//close();
</script>
<%
			end if
		else
%>
<script language="JavaScript">
	alert('参数传递错误，请重试');
	dialogArguments.location.reload();
	close();
</script>
<%
		end if
	else
%>
<script language="JavaScript">
	alert('参数传递错误，请重试');
	dialogArguments.location.reload();
	close();
</script>
<%
	end if
	if Err.Number = 0 then
%>
<script language="JavaScript">
	dialogArguments.location.reload();
	close();
</script>
<%
	else
%>
<script language="JavaScript">
	alert('重命名失败');
	dialogArguments.location.reload();
	//close();
</script>
<%
	end if
	Set FsoObj = Nothing
end if
%>