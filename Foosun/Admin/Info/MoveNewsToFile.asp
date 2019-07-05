<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->

<%
Dim DBC,Conn,RecordConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + server.mappath(RecordDataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set RecordConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010500") then Call ReturnError1()
Dim ClassID,NewsID,SearchSql,RsClassObj,FileDateNum
ClassID = Request("ClassID")
NewsID = Request("NewsID")
if ClassID <> "" then
	SearchSql = "Select * from FS_NewsClass where ClassID='" & ClassID & "'"
	Set RsClassObj = Conn.Execute(SearchSql)
	if Not RsClassObj.Eof then
		FileDateNum = RsClassObj("FileTime")
	else
		FileDateNum = 0
	end if
	if IsNull(FileDateNum) then FileDateNum = 0
	Set RsClassObj = Nothing
	If IsSqlDataBase=0 then
		SearchSql = "Select * from FS_News where DateAdd('d'," & FileDateNum & ",AddDate)<=Now() and ClassID='" & ClassID & "'"
	Else
		SearchSql = "Select * from FS_News where DateAdd(d," & FileDateNum & ",AddDate)<=getdate() and ClassID='" & ClassID & "'"
	End If
else
	if NewsID <> "" then
		NewsID = Replace(NewsID,"***","','")
		SearchSql = "Select * from FS_News where NewsID in ('" & NewsID & "')"
	else
		Response.Write("<script>alert('参数错误');</script>")
		Response.End
	end if
end if
Dim RsNewsObj,RsFileObj,FiledObj,MoveNumber
MoveNumber = 0
Set RsNewsObj = Conn.Execute(SearchSql)
Set RsFileObj = Server.CreateObject(G_FS_RS)
Call CreatColumn
do while Not RsNewsObj.Eof
	RsFileObj.Open "Select * from FS_News where NewsID='" & RsNewsObj("NewsID") & "'",RecordConn,3,3
	if RsFileObj.Eof then
		RsFileObj.AddNew
		for Each FiledObj in RsNewsObj.Fields
			if LCase(FiledObj.Name) <> "id" then
				RsFileObj(FiledObj.Name) = RsNewsObj(FiledObj.Name)
			end if
		Next
		RsFileObj("FileTime") = Now()
	end if
	RsFileObj.Update
	RsFileObj.Close
	Conn.Execute("Delete * from FS_News where NewsID='" & RsNewsObj("NewsID") & "'")
	MoveNumber = MoveNumber + 1
	RsNewsObj.MoveNext
Loop
Dim PromptStr
PromptStr = "归档" & MoveNumber & "条新闻"
Sub CreatColumn()
	On error resume next
	Dim RecordRs
	Set RecordRs=RecordConn.Execute("Select TitleShowReview from FS_News where 1=2")
	If err.number<>0 then 
		RecordConn.Execute("Alter Table FS_News Add TitleShowReview Int")
	End If
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="60"><div align="center"> 
        <% = PromptStr %>
      </div></td>
  </tr>
  <tr> 
    <td height="30"><div align="center"> 
        <input type="button" onClick="window.close();" name="Submit111" value=" 关 闭 ">
      </div></td>
  </tr>
</table>
</body>
</html>
<%
Set RsNewsObj = Nothing
Set RsFileObj = Nothing
Set Conn = Nothing
Set RecordConn = Nothing
%>