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
Dim LimitUpFileFlag,CurrPath,ShowVirtualPath
LimitUpFileFlag = Request("LimitUpFileFlag")
CurrPath = Request("CurrPath")
ShowVirtualPath = Request("ShowVirtualPath")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选择图片</title>
<style type="text/css">
<!--
.PreviewStyle {
	border: 2px outset #CCCCCC;
}
 BODY   {border: 0; margin: 0; background: buttonface; cursor: default; font-family:宋体; font-size:9pt;}
 BUTTON {width:5em}
 TABLE  {font-family:宋体; font-size:9pt}
 P      {text-align:center}
-->
</style>
</head>
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body leftmargin="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><select onChange="ChangeFolder(this.value);" id="FolderSelectList" style="width:100%;" name="select">
		<option selected value="<% = CurrPath %>"><% = CurrPath %></option>
      </select></td>
    <td rowspan="2" align="center"><iframe id="PreviewArea" style="width:100%;height:380" frameborder="1" scrolling="auto" src="PreviewImage.asp"></iframe></td>
  </tr>
  <tr> 
    <td width="50%" align="center"> <iframe id="FolderList" width="100%" height="350" frameborder="1" src="FolderImageList.asp?CurrPath=<% = CurrPath %>&ShowVirtualPath=<% = ShowVirtualPath %>"></iframe></td>
  </tr>
  <tr> 
    <td height="35" colspan="2"> 
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="80" height="30"> <div align="center">Url地址</div></td>
          <td><input style="width:100%" type="text" name="UserUrl"></td>
          <td width="100" align="center">
			<input type="button" onClick="SetUserUrl();" name="Submit" value=" 确 定 ">
          </td>
          <td width="100" align="center"><input type="button" <% if LimitUpFileFlag = "yes" then Response.Write("disabled") %> onClick="UpFile();" name="Submit2" value=" 上 传 ">
          </td>
          <td width="100" align="center"><input onClick="window.close();" type="button" name="Submit3" value=" 取 消 "> 
          </td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
function ChangeFolder(FolderName)
{
	frames["FolderList"].location='FolderImageList.asp?CurrPath='+FolderName;
}
function UpFile()
{
	OpenWindow('Frame.asp?FileName=UpFileForm.asp&Path='+frames["FolderList"].CurrPath,350,150,window);
	frames["FolderList"].location='FolderImageList.asp?CurrPath='+frames["FolderList"].CurrPath;
}
function SetUserUrl()
{
	if (document.all.UserUrl.value=='') alert('请填写Url地址');
	else
	{
		window.returnValue=document.all.UserUrl.value;
		window.close();
	}
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>