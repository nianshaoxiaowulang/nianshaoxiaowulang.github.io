<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="Function.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="RefreshFunction.asp" -->
<!--#include file="SelectFunction.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030400") then Call ReturnError1()
Dim Types,SearchSql,RsSearchObj,PromptInfo,NewsNo,NewsTotalNum,NumClass,NewNum,FromDate,TentDate,ClassID
Dim AlreadyRefreshID
NumClass = Request("NumClass")
NewNum = Request("NewNum")
FromDate = Request("FromDate")
TentDate = Request("TentDate")
ClassID = Request("ClassID")
AlreadyRefreshID = Request("AlreadyRefreshID")
NewsNo = Request("NewsNo")
Types = Request("Types")
if NewNum = "" then
	NewNum = 10
else
	NewNum = CInt(NewNum)
end if
if NumClass = "" then
	NumClass = 10
else
	NumClass = CInt(NumClass)
end if
if NewsNo = "" then
	NewsNo = 0
else
	NewsNo = CInt(NewsNo)
end if
if Types = "DatesType" then
	if AlreadyRefreshID = "" then
		If IsSqlDataBase=0 then
			SearchSql = "Select top 1 * from FS_download where AuditTF=1 and AddTime>=#" & FromDate & "# And AddTime<=#" & TentDate & "# order by ID" 
		Else
			SearchSql = "Select top 1 * from FS_download where AuditTF=1 and AddTime>='" & FromDate & "' And AddTime<='" & TentDate & "' order by ID" 
		End If
	else
		If IsSqlDataBase=0 then
			SearchSql = "Select top 1 * from FS_download where ID>" & AlreadyRefreshID & " and AuditTF=1 and AddTime>=#" & FromDate & "# And AddTime<=#" & TentDate & "# order by ID" 
		Else
			SearchSql = "Select top 1 * from FS_download where ID>" & AlreadyRefreshID & " and AuditTF=1 and AddTime>='" & FromDate & "' And AddTime<='" & TentDate & "' order by ID" 
		End If
	end if
	If IsSqlDataBase=0 then
		NewsTotalNum = Conn.Execute("Select count(*) from FS_download where AuditTF=1 and AddTime>=#" & FromDate & "# And AddTime<=#" & TentDate & "#")(0)
	Else
		NewsTotalNum = Conn.Execute("Select count(*) from FS_download where AuditTF=1 and AddTime>='" & FromDate & "' And AddTime<='" & TentDate & "'")(0)
	End If
elseif Types = "NewType" then
	if CInt(NewsNo) < CInt(NewNum) then
		if AlreadyRefreshID = "" then
			SearchSql = "Select Top 1 * from FS_download where AuditTF=1 order by ID"
		else
			SearchSql = "Select Top 1 * from FS_download where ID>" & AlreadyRefreshID & " and AuditTF=1 order by ID"
		end if
	else
		SearchSql = "Select * from FS_download where 1=0"
	end if
	NewsTotalNum = NewNum
elseif Types = "AllType" then
	if AlreadyRefreshID = "" then
		SearchSql = "Select top 1 * from FS_download where AuditTF=1 Order by ID"
	else
		SearchSql = "Select top 1 * from FS_download where ID>" & AlreadyRefreshID & " and AuditTF=1 Order by ID"
	end if
	NewsTotalNum = Conn.Execute("Select count(*) from FS_download where AuditTF=1")(0)
elseif Types = "ClassType" then
	if NumClass <> "0" then
		if (NumClass <> "") And (ClassID <> "") then
			if CInt(NewsNo) < CInt(NumClass) then
				if AlreadyRefreshID = "" then
					SearchSql = "Select Top 1 * from FS_download where AuditTF=1 and ClassID='" & ClassID & "' order by id"
				else
					SearchSql = "Select Top 1 * from FS_download where ID>" & AlreadyRefreshID & " and AuditTF=1 and ClassID='" & ClassID & "' order by id"
				end if
			else
				SearchSql = "Select * from FS_download where 1=0"
			end if
			NewsTotalNum = Conn.Execute("Select count(*) from FS_download where AuditTF=1 and ClassID='" & ClassID & "'")(0)
		else
			SearchSql = ""
			NewsTotalNum = 0
		end if
	else
		PromptInfo = "没有刷新下载&nbsp;&nbsp;<a href=""RefreshDownload.asp"">返回</a>"
		NewsTotalNum = 0
		Call PromptFunction
	end if
else
	SearchSql = ""
	NewsTotalNum = 0
end if
if SearchSql <> "" then
	Set RsSearchObj = Server.CreateObject(G_FS_RS)
	RsSearchObj.Open SearchSql,Conn,1,1
	if RsSearchObj.Eof then
		PromptInfo = "刷新下载成功<font color=red><b>" & NewsTotalNum & "</b></font>条下载<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Call PromptFunction
	else
		Refreshdownload RsSearchObj
		NewsNo = NewsNo + 1
		AlreadyRefreshID = RsSearchObj("ID")
		Response.Write("<meta http-equiv=""refresh"" content=""0;url=RefreshdownloadSave.asp?NumClass=" & NumClass & "&NewsNo=" & NewsNo & "&NewNum=" & NewNum & "&FromDate=" & FromDate & "&TentDate=" & TentDate & "&ClassID=" & ClassID & "&AlreadyRefreshID=" & AlreadyRefreshID & "&Types=" & Types & """>")
		PromptInfo = "共有<font color=red><b>" & NewsTotalNum & "</b></font>条下载需要刷新<br><br>正在刷新第<font color=red><b>" & NewsNo & "</b></font>条下载"
		PromptInfo = PromptInfo & "按确定键返回！<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Call PromptFunction
	end if
	Set RsSearchObj = Nothing
else
	PromptInfo = "刷新下载成功<font color=red><b>" & NewsTotalNum & "</b></font>条下载<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
	Call PromptFunction
end if

Sub PromptFunction()
	Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link rel="stylesheet" href="../../../CSS/FS_css.css">
<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <form method=post action=RefreshDownload.asp>
    <tr> 
      <td height="150"> <div align="center"> 
          <% = PromptInfo %>
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
End Sub
%>