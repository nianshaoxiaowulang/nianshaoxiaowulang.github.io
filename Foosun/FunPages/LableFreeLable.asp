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
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
Dim TempFreeLableStr,SqlStr,Rs
SqlStr = "select name,freelableid from FS_freelable"
Set Rs = conn.Execute(SqlStr)
While not Rs.eof
	TempFreeLableStr = TempFreeLableStr&"<option value='"&Rs("freelableid")&"'>"&Rs("name")&"</option>"
	Rs.movenext
Wend
%>
<html>
<head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<title>设置自由标签</title>
</head>
<body>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="30%">已有标签 
      <select id="FreeLableList" name="FreeLableList" style="width:60%;" onchange="FreeLableListChange()">
        <option value="" selected>选择自由标签</option>
        <% =TempFreeLableStr %>
      </select> </td>
    <td width="70%" height=320 rowspan="6"><iframe style="width:100%;height:100%" id="PreviewStyle" name="PreviewStyle" src="PreviewStyle.asp"></iframe></td>
  </tr>
  <tr> 
    <td><div align="left">查询数量 
        <input name="QueryNumber" type="text" id="QueryNumber" style="width:60%;" value="10">
      </div></td>
  </tr>
  <tr> 
    <td><div align="left">水平间距 
        <input name="ColSpace" type="text" id="ColSpace" style="width:60%;" value="">px
      </div></td>
  </tr>
  <tr> 
    <td>垂直间距 
      <input name="RowSpace" type="text" id="RowSpace" style="width:60%;" value="">px
    </td>
  </tr>
  <tr> 
    <td align="left">行　　数 <input name="RowNum" type="text" id="RowNum" style="width:60%;" value="1"></td>
  </tr>
  <tr> 
    <td align="left">列　　数 <input name="ColNum" type="text" id="ColNum" style="width:60%;" value="1"></td>
  </tr>
  <tr> 
    <td height="47" colspan="2" align="center"> 
      <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 "> 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 "> 
    </td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
function FreeLableListChange()
{
	PreviewStyle.location = "PreviewStyle.asp?Freelableid="+document.all.FreeLableList.value;
}
function InsertScript()
{	
	var FreeLableID = document.all.FreeLableList.value;
	if(FreeLableID == "")
	{
		alert("还没有选择自由标签");
		document.all.FreeLableList.focus();
		return;
	}
	var QueryNumberStr= document.all.QueryNumber.value;
	if (IsNumeric(QueryNumberStr) == false)
	{
		alert("查询数量包含非法字符");
		document.all.QueryNumber.focus();
		return;
	}
	var ColSpaceStr=document.all.ColSpace.value;
	if (IsNumeric(ColSpaceStr) == false)
	{
		alert("水平间距包含非法字符");
		document.all.ColSpace.focus();
		return;
	}
	var RowSpaceStr=document.all.RowSpace.value;
	if (IsNumeric(RowSpaceStr) == false)
	{
		alert("垂直间距包含非法字符");
		document.all.RowSpace.focus();
		return;
	}
	var ColNumStr='';
	if (document.all.ColNum.value=='') ColNumStr='1';
	else ColNumStr=document.all.ColNum.value;
	if (IsNumeric(ColNumStr) == false)
	{
		alert("水平重复包含非法字符");
		document.all.ColNum.focus();
		return;
	}
	var RowNumStr='';
	if (document.all.RowNum.value=='') RowNumStr='1';
	else RowNumStr=document.all.RowNum.value;
	if (IsNumeric(RowNumStr) == false)
	{
		alert("水平重复包含非法字符");
		document.all.RowNum.focus();
		return;
	}
	window.returnValue='{%=FreeLable("'+FreeLableID+'","'+QueryNumberStr+'","'+ColSpaceStr+'","'+RowSpaceStr+'","'+ColNumStr+'","'+RowNumStr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function IsNumeric(Str)
{
	var i,NumericStr="0123456789";
	if(Str=="") return true;
	if(Str.substr(0,1) == "0" && Str.length > 1) return false;
	for(i=0;i<Str.length;i++)
		if(NumericStr.indexOf(Str.substr(i,1)) == -1)
			return false;
	return true;
}
</script>