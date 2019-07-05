<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
	Dim DBC,Conn
	On Error Resume Next
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030705") then Call ReturnError1()
Dim Path,FileName,EditFile
Path = Request("Path")
FileName = Request("FileName")
if Path = "/" then
	EditFile = Path & FileName
else
	EditFile = Path & "/" & FileName
end if
Dim FsoObj,FileObj,FileStreamObj,FileContent
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FileObj = FsoObj.GetFile(Server.MapPath(EditFile))
Set FileStreamObj = FileObj.OpenAsTextStream(1)
if Not FileStreamObj.AtEndOfStream then
	FileContent = FileStreamObj.ReadAll
else
	FileContent = ""
end if
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
FileContent =Replace(Replace(FileContent,"""","%22"),"'","%27")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>模板编辑----<% = EditFile %></title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body leftmargin="2" topmargin="2" onclick="return true;" onselectstart="return false;" oncontextmenu="//showMenu(MouseRightMenu);return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="SaveFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td><div align="center">
              <input name="FileContent" type="hidden" value="<% = FileContent %>">
              编辑文件: 
              <% = EditFile %>
            </div></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr id="LableToolBar" height="32"> 
    <td><iframe id="LableListWindow" src="LableInsert.asp" scrolling="no" width="100%" height="100%" frameborder="0"></iframe></td>
  </tr>
  <tr> 
    <td><iframe id="Editer" src="../../Editer/Editer.asp?Path=<% = Path %>&FileName=<% = FileName %>" scrolling="no" width="100%" frameborder="0"></iframe></td>
  </tr>
</table>
</body>
</html>
<iframe id="SaveFrame" src="SaveTempletFile.asp" width="0" height="0"></iframe>
<script language="JavaScript">
var Path=escape('<% = Path %>');
var FileName=escape('<% = FileName %>');
SetEditAreaHeight();
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-document.all.LableToolBar.height-32;
	document.all.Editer.height=EditAreaHeight;
}
function SaveFile()
{
	if (frames["Editer"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到编辑模式');return;}
	frames["Editer"].EditArea.document.body.contentEditable="false";
	var SaveForm=frames["SaveFrame"].document.SaveFileForm;
	SaveForm.Path.value=Path;
	SaveForm.FileName.value=FileName;
	SaveForm.FileContent.value=frames["Editer"].EditArea.document.documentElement.innerHTML;
	SaveForm.Result.value='Submit';
	SaveForm.submit();
	SaveForm.Result.value='';
	AlreadyEdit=false;
	frames["Editer"].EditArea.document.body.contentEditable="true";
}
</script>
