<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P031204") then Call ReturnError()
Dim StyleID,Operation
StyleID = Request("ID")
Operation = Request("Operation")
if Operation = "Del" then
	if StyleID <> "" then
		Dim DelSql
		On Error Resume Next
		DelSql = "Delete from FS_MallListStyle where ID in (" & Replace(StyleID,"***",",") & ")"
		Conn.Execute(DelSql)
		if Err.Number = 0  then
			%>
			<script language="javascript">
				dialogArguments.location.reload();
				window.close();
			</script>
			<%
		else
			%>
			<script language="javascript">
				alert('删除失败');
				dialogArguments.location.reload();
				window.close();
			</script>
			<%
		end if
	else
		%>
		<script language="javascript">
			alert('没有要删除的样式');
			dialogArguments.location.reload();
			window.close();
		</script>
		<%
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除列表样式</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body scroll=no bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <form name="DelForm" method="post" action="">
    <tr> 
      <td width="42%" height="60"> <div align="right"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
      <td width="58%" height="80" colspan="2" valign="middle"> 确定要删除此样式吗? 
        <input name="hiddenField" type="hidden" value="<% = StyleID %>">
        <input type="hidden" name="Operation" value="Del"></td>
    </tr>
    <tr> 
      <td colspan="3"> <div align="center"> 
          <input name="Submitsadf" type="submit" id="Submitsadf" value=" 确 定 ">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="button" onClick="window.close();" name="Submit3" value=" 取 消 ">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>