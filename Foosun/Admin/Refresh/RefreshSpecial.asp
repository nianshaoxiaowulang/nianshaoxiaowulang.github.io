<% Option Explicit %>
<!--#include file="Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P030500") then Call ReturnError1()
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>专题首页生成</title>
</head>

<body topmargin="2" leftmargin="2" oncontextmenu="return false;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>专题首页生成管理</strong></div></td>
</tr>
</table>
<table width="100%"  border="0" cellspacing="8" cellpadding="0">
  <tr>
    <td width="9%">&nbsp;</td>
    <td width="11%">&nbsp;</td>
    <td width="80%">&nbsp;</td>
  </tr>
  <form action="RefreshSpecialSave.asp?Types=SpecialOne" method="post" name="ReSpecialOneForm">
  <tr>
    <td>&nbsp;</td>
    <td>分类生成</td>
    <td><select name="SpecialID" style="width:20%">
	<%
		Dim RsTempSpObj
		Set RsTempSpObj = Conn.Execute("Select SpecialID,CName from FS_Special order by AddTime desc")
		do while not RsTempSpObj.eof 
	%>
		<option value="<%=RsTempSpObj("SpecialID")%>"><%=RsTempSpObj("CName")%></option>
	<%
		RsTempSpObj.MoveNext
		Loop
		RsTempSpObj.Close
		Set RsTempSpObj = Nothing
	%>
    </select>
        <input name="imageField" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0"> 
      </td>
  </tr>
  </form>
  <form action="RefreshSpecialSave.asp?Types=SpecialAll" method="post" name="ReSpecialAllForm">
  <tr>
    <td>&nbsp;</td>
    <td>全部生成</td>
      <td><input name="imageField2" type="image" src="../../Images/Publish.gif" width="75" height="21" border="0"></td>
  </tr>
  </form>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2"><font color=red>注意:如果需要生成的专题较多，建议采用分类生成,并且在生成过程中，请勿刷新此页面</font></td>
  </tr>
</table>
</body>
</html>
