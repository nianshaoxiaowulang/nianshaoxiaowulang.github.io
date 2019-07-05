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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>专题导航</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0">
<div align="center">
  <table width="90%" border="0" cellspacing="0" cellpadding="0">
    <form action="" method="post" name="CNListForm">
      <tr> 
        <td height="30">显示图片 <select name="PicTF" id="PicTF">
            <option value="1" selected>是</option>
            <option value="0">否</option>
          </select></td>
        <td>专题栏目
          <select   style="width:70%;" name="SpecialClassID" id="SpecialClassID">
            <%
			Dim RsNewsClassObj
			Set RsNewsClassObj = Conn.Execute("Select CName,EName from FS_special order by id desc")
			do while Not RsNewsClassObj.Eof
			%>
            <option  value="<% = RsNewsClassObj("EName") %>"> 
            <% = RsNewsClassObj("CName") %>
            </option>
            <%
				RsNewsClassObj.MoveNext
			loop
			Set RsNewsClassObj = Nothing
		  %>
          </select></td>
      </tr>
      <tr> 
        <td width="48%" height="30"><div align="left">图片高度 
            <input name="PicHeight" type="text" id="PicHeight"  value="80" size="15">
          </div></td>
        <td width="52%"><div align="left">图片宽度 
            <input name="PicWidth" type="text" id="RowNumber2"  value="60" size="15">
          </div></td>
      </tr>
      <tr> 
        <td height="30">导航数量 
          <input name="Dhang" type="text" id="Dhang" value="100" size="15"> 
        </td>
        <td height="30">导航 css 
          <input name="SpecialCss" type="text" id="SpecialCss" size="15"> </td>
      </tr>
      <tr> 
        <td height="30"><div align="left">更多文字 
            <input name="SpecialMore" type="text" id="TitleNumber3" value="更多.." size="15">
          </div></td>
        <td height="30"><div align="left"></div></td>
      </tr>
      <tr> 
        <td height="30" colspan="2"><div align="center"> 
            <input type="button" onClick="InsertScript();" name="Submit" value=" 确 定 ">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
          </div></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
function InsertScript()
{
	var NewsNumberStr='';
	NewsNumberStr=document.all.PicTF.value;
	var RowNumberStr='';
	RowNumberStr=document.all.SpecialClassID.value;
	var PicHeightstr='';
	if (document.all.PicHeight.value=='') PicHeightstr='80';
	else PicHeightstr=document.all.PicHeight.value;
	var PicWidthstr='';
	if (document.all.PicWidth.value=='') PicWidthstr='60';
	else PicWidthstr=document.all.PicWidth.value;
	var Dhangstr='';
	if (document.all.Dhang.value=='') Dhangstr='100';
	else Dhangstr=document.all.Dhang.value;
	var SpecialCssstr='';
	if (document.all.SpecialCss.value=='') SpecialCssstr='';
	else SpecialCssstr=document.all.SpecialCss.value;
	var SpecialMorestr='';
	if (document.all.SpecialMore.value=='') SpecialMore='';
	else SpecialMorestr=document.all.SpecialMore.value;
	window.returnValue='{%=SpecialNavi("'+NewsNumberStr+'","'+RowNumberStr+'","'+PicHeightstr+'","'+PicWidthstr+'","'+Dhangstr+'","'+SpecialCssstr+'","'+SpecialMorestr+'")%}';
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>