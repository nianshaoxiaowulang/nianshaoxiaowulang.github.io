<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim Path,FileName,EditFile,FileContent
Path = Request.QueryString("Path")
FileName = Request.QueryString("FileName")
EditFile = Server.MapPath(Path) & "\" & FileName
Dim FsoObj,FileObj,FileStreamObj
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FileObj = FsoObj.GetFile(EditFile)
Set FileStreamObj = FileObj.OpenAsTextStream(1)
if Not FileStreamObj.AtEndOfStream then
	FileContent = FileStreamObj.ReadAll
else
	FileContent = ""
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ı��༭��---<% = FileName %></title>
</head>
<link rel="stylesheet" href="Editer.css">
<style>
.BtnMouseOver {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #000000;
	border-right-color: #000000;
	border-bottom-color: #000000;
	border-left-color: #000000;
	cursor: default;
}
</style>
<script language="JavaScript" src="Editer.js"></script>
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body bgcolor="#EFEFEF" topmargin="2" leftmargin="2" oncontextmenu="return false;" scroll=no>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="SaveFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td><div align="center">
              �༭�ļ�: 
              <% = EditFile %>
            </div></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form id="TempForm" name="TempForm">
  <tr id="ToolBar" height="32">
    <td height="30"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="30" height="30"> <div align="center"><img alt="������" onClick="InsertTable();" src="../Images/Editer/table.gif" width="23" height="22" class="ToolBtn"></div></td>
          <td width="30" height="30"> <div align="center"><img alt="����ͼƬ" onClick="InsertImg();" src="../Images/Editer/image.gif" width="23" height="22" class="ToolBtn"></div></td>
          <td width="30" height="30"> <div align="center"><img alt="����JS" src="../Images/Editer/InsertJs.gif" width="23" height="22" class="ToolBtn" onclick="InserJS();"></div></td>
          <td width="30" height="30"> <div align="center"><img alt="�����" onClick="InsertForm();" src="../Images/Editer/form.gif" width="23" height="22" class="ToolBtn"></div></td>
		  <td width="160"> 
            <select onChange="InsertFormCode(this)" title="���������" style="width:100%;" name="select">
				<option value="" selected>���������</option>
				<option value="input" InsertType="text">�ı��ֶ�</option>
				<option value="input" InsertType="hidden">������</option>
				<option value="textarea" InsertType="">�ı�����</option>
				<option value="input" InsertType="checkbox">��ѡ��</option> 
				<option value="input" InsertType="radio">��ѡ��ť</option> 
				<option value="select" InsertType="">�б�/�˵�</option> 
				<option value="input" InsertType="submit">�ύ��ť</option> 
				<option value="input" InsertType="reset">���ð�ť</option>  
				<option value="input" InsertType="button">��ť</option> 
            </select></td>
          <td height="30">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td id="EditAreaTR"><textarea onChange="SetEditFlag();" id="EditArea" style="width:100%;" name="textarea"><% =FileContent %></textarea></td>
  </tr>
 </form>
</table>
</body>
</html>
<iframe id="SaveFrame" src="SaveFileFrame.asp" width="0" height="0"></iframe>
<%
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
%>
<script language="JavaScript">
var AlreadyEdit=false;
var bInitialized = false;
var Path='<% = Path %>';
var FileName='<% = FileName %>';
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-document.all.ToolBar.height-32;
	document.all.EditArea.style.height=EditAreaHeight;
}
SetEditAreaHeight();
window.onresize=SetEditAreaHeight;
window.onunload=PromptSave;
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (bInitialized) return;
	bInitialized = true;
	var i,j,s,curr;
	for (i=0; i<document.body.all.length;i++)
	{
		curr=document.body.all[i];
		if (curr.className == "ToolBtn") InitBtn(curr);
	}
}
function SetEditFlag()
{
	AlreadyEdit=true;
}
function PromptSave()
{
	if (AlreadyEdit==true)
	{
		if (confirm('�ļ��Ѿ��޸ģ�Ҫ������')==true) SaveFile();
	}
}
function InserJS()
{
	var Str='<'+'script language="JavaScript" type="text/JavaScript">'+'\n\n</script'+'>';
	InsertStr(Str);
}
function InsertTable()
{
	document.all.EditArea.focus();
	var Str='<table>\n <tr>\n <td>   </td>\n </tr>\n</table>';
	InsertStr(Str);
}
function InsertImg()
{
	var ReturnValue=OpenWindow('../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,290,window)
	if (ReturnValue!='')
	{
		var Str='<img src="'+ReturnValue+'">';
		InsertStr(Str);
	}
}
function InsertFormCode(Obj)
{
	var CreateObj;
	var str;
	if (Obj.value!='')
	{
		CreateObj=document.createElement(Obj.value);
		switch (Obj.value)
		{
			case 'input':
				CreateObj.type=Obj.options(Obj.selectedIndex).InsertType;
				str='<'+Obj.value +' type="'+CreateObj.type+'" name="" value="">\n'
				break;
			case 'select':
				str='<select name="">\n<option value=""> </option>\n</select>\n'
				break;
			case 'textarea':
				str='<textarea name=""> </textarea>\n'
				//CreateObj.cols='1';
				//CreateObj.rows='1';
				break;
			default:
				return;
		}
		//CreateObj.name='EditArea';
		//document.all.TempForm.appendChild(CreateObj);
		//InsertStr(document.all.TempForm.innerHTML);
		//document.all.TempForm.innerHTML='';
		InsertStr(str)
		Obj.options(0).selected=true;
	}
}
function InsertStr(Str)
{
	document.all.EditArea.focus();
	var RangeObj=document.selection.createRange();
	RangeObj.text=Str;
}
function InsertForm()
{
	var Str='<form name="" action="" method="post">\n\n</form>';
	InsertStr(Str);
}
function SaveFile()
{
	var SaveForm=frames["SaveFrame"].document.SaveFileForm;
	SaveForm.Path.value=Path;
	SaveForm.FileName.value=FileName;
	SaveForm.FileContent.value=document.all.EditArea.value;
	SaveForm.Result.value='Submit';
	SaveForm.submit();
	SaveForm.Result.value='';
	AlreadyEdit=false;
}
</script>