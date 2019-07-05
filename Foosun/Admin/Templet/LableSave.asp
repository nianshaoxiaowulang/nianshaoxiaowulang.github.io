<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P0302") then Call ReturnError()
	Dim Result,LableID
	Result = Request("Result")
	if Result = "Submit" then
	Dim RsLableObj
	Set RsLableObj = Server.CreateObject(G_FS_RS)
	LableID = Request("LableID")
	if LableID = "" then
		RsLableObj.Open "Select * from FS_Lable where 1=0",Conn,3,3
		RsLableObj.AddNew()
	else
		RsLableObj.Open "Select * from FS_Lable where ID="&LableID,Conn,3,3
	end if
	RsLableObj("LableName") = Request.Form("LableName")
	RsLableObj("LableContent") = Request.Form("LableContent")
	RsLableObj("LableType") = Request.Form("LableType")
	RsLableObj.UpDate()
	RsLableObj.Close()
	Set RsLableObj = Nothing
	Set Conn = Nothing
	if Err.Number <> 0 then
%>
<script language="JavaScript">
	alert('´íÎó£º\n<% = Err.Description %>');
</script>
<%
	else
%>
<script language="JavaScript">
	alert('±£´æ³É¹¦');
</script>
<%
	end if
end if
%>
<html>
<body>
<form action="" name="SaveLableForm" method="post">
<input name="Result" type="hidden" value="">
<input name="LableID" type="hidden" value="">
<input name="LableName" type="hidden" value="">
<input name="LableContent" type="hidden" value="">
<input name="LableType" type="hidden" value="">
</form>
</body>
</html>