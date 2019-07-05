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
<!--#include file="../../Inc/Function.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
Dim TempClassListStr
	TempClassListStr = ClassList("ClassEName")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>信息统计标签属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="96%" border="0" cellspacing="3" cellpadding="0">
    <tr>
      <td width="42%" height="30"><div align="center">栏目名称 
          <select name="ClassList" id="ClassList" style="width:70%;">
            <option value="" selected>栏目选择</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
      <td width="50%"><div align="center">排列方式 
          <select name="ShowMode" id="ShowMode" style="width:70%;">
            <option value="1" selected>横向</option>
            <option value="0">纵向</option>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> CSS 样式 
          <input type="text" style="width:70%;" name="CssFile" id="CssFile">
        </div></td>
      <td height="30"><div align="center" style="display:none;">分隔符号 
          <input name="CutFlag" type="text" id="CutFlag2" style="width:70%;" value="[]">
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="center">
          <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 ">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
function InsertScript()
{
	var ClassListStr=document.all.ClassList.value;
	var ShowModeStr=document.all.ShowMode.value;
	var CutFlagStr='';
	var CssFileStr=document.all.CssFile.value;
	
	window.returnValue='{%=InfoStat("'+ClassListStr+'","'+ShowModeStr+'","'+CssFileStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>