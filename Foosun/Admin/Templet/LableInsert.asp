<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030705") then Call ReturnError()
Dim LableSql,RsLableObj
LableSql = "Select * from FS_Lable"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<style type="text/css">
<!--
.ToolBarGegin {
	border-right-width: thin;
	border-right-style: ridge;
	height: 100%;
	width: 3px;
	border-right-color: #FFFFFF;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ǩ�б�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<link href="../../Editer/Editer.css" rel="stylesheet">
<body scroll="no" ondragstart="return false;" onselectstart="return false;" oncontextmenu="//showMenu(MouseRightMenu);return false;" topmargin="0" leftmargin="0">
<table height="32" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
  <tr> 
    <td width="1"> <div align="center" class="ToolBarGegin"></div></td>
    <td width="1"> <div align="center" class="ToolBarGegin"></div></td>
    <td width="200"> <div align="center"> 
        <select style="width:96%;" name="LableList">
          <option selected LableName="">ѡ��Ҫ����ı�ǩ</option>
          <%
Set RsLableObj = Conn.Execute(LableSql)
do while Not RsLableObj.Eof
%>
          <option LableID="<% = RsLableObj("ID") %>" LableName="<% = RsLableObj("LableName") %>"> 
          <% = RsLableObj("LableName") %>
          </option>
          <%
	RsLableObj.MoveNext
loop
Set RsLableObj = Nothing
%>
        </select>
      </div></td>
    <td width="1"> <div align="center" class="ToolSeparator"></div></td>
    <td width="30"><div align="center"><img onClick="BrowerLableAttribute();" onmouseout="this.className='';" onmouseup="this.className='ToolBtnMouseOver';"; onmousedown="this.className='ToolBtnMouseDown';" onmouseover="this.className='ToolBtnMouseOver';" alt="�鿴��ǩ" src="../../Images/Lable/ReviewLabe.gif" width="24" height="24"></div></td>
    <td width="30"><div align="center"><img onClick="InsertLable();" onmouseout="this.className='';" onmouseup="this.className='ToolBtnMouseOver';"; onmousedown="this.className='ToolBtnMouseDown';" onmouseover="this.className='ToolBtnMouseOver';" alt="�����ǩ" src="../../Images/Lable/InsertLable.gif" width="24" height="24"></div></td>
    <td width="30"><div align="center"><img onClick="OpenWindowInsertLable();" onmouseout="this.className='';" onmouseup="this.className='ToolBtnMouseOver';"; onmousedown="this.className='ToolBtnMouseDown';" onmouseover="this.className='ToolBtnMouseOver';" alt="ѡ���ǩ" src="../../Images/Lable/selectLable.gif" width="24" height="24"></div></td>
    <td style="display:none;" width="30"><div align="center"><img onClick="window.location.reload();" onmouseout="this.className='';" onmouseup="this.className='ToolBtnMouseOver';"; onmousedown="this.className='ToolBtnMouseDown';" onmouseover="this.className='ToolBtnMouseOver';" alt="ˢ�±�ǩ�б�" src="../../Images/Lable/RefreshLable.gif" width="24" height="24"></div></td>
  </tr>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
function InsertLable()
{
	var LableName=document.all.LableList.options(document.all.LableList.selectedIndex).LableName;
	if (LableName!='')
	{
		parent.frames["Editer"].InsertHTMLStr(LableName);
	}
	//document.all.LableList(0).selected=true;
}
function BrowerLableAttribute()
{
	var LableID=document.all.LableList.options(document.all.LableList.selectedIndex).LableID;
	if (LableID!=null)
	{
		OpenWindow('LableAttribute.asp?ID='+LableID,360,190);
	}
	parent.frames["Editer"].EditArea.focus();
}
function OpenWindowInsertLable()
{
	var ReturnValue=OpenWindow('LableOpenWindowInsert.asp',420,300);
	if (ReturnValue!='')
	{
		parent.frames["Editer"].InsertHTMLStr(ReturnValue);
	}
	//parent.location='LableOpenWindowInsert.asp';
}
</script>