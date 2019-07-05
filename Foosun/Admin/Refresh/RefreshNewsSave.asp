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
Dim Types,SearchSql,RsSearchObj,PromptInfo,NewsNo,NumClass,NewNum,FromDate,TentDate,ClassID
Dim AlreadyRefreshID,AllClassID
'NumClass = Request("NumClass")
NewNum = Request("NewNum")
ClassID = Request("ClassID")
AlreadyRefreshID = Request("AlreadyRefreshID")
NewsNo = Request("NewsNo")
Types = Request("Types")
if NewNum = "" then
	NewNum = 10
else
	NewNum = CLng(NewNum)
end if
'if NumClass = "" then
'	NumClass = 10
'else
'	NumClass = CInt(NumClass)
'end if
if NewsNo = "" then
	NewsNo = 0
else
	NewsNo = CInt(NewsNo)
end if
if Types = "DatesType" then
	'---------------/l
	FromDate = Request("FromDate") & " 00:00:00"
	TentDate = Request("TentDate") & " 23:59:59"
	'----------------
	if AlreadyRefreshID = "" then
		If IsSqlDataBase=0 then
			SearchSql = "Select top 1 * from FS_News where AuditTF=1 and AddDate>=#" & FromDate & "# And AddDate<=#" & TentDate & "# and DelTF=0 and HeadNewsTF=0 order by ID" 
		Else
			SearchSql = "Select top 1 * from FS_News where AuditTF=1 and AddDate>='" & FromDate & "' And AddDate<='" & TentDate & "' and DelTF=0 and HeadNewsTF=0 order by ID" 
		End If
	else
		If IsSqlDataBase=0 then
			SearchSql = "Select top 1 * from FS_News where ID>" & AlreadyRefreshID & " and AuditTF=1 and AddDate>=#" & FromDate & "# And AddDate<=#" & TentDate & "# and DelTF=0 and HeadNewsTF=0 order by ID" 
		Else
			SearchSql = "Select top 1 * from FS_News where ID>" & AlreadyRefreshID & " and AuditTF=1 and AddDate>='" & FromDate & "' And AddDate<='" & TentDate & "' and DelTF=0 and HeadNewsTF=0 order by ID" 
		End If
	end if
	If IsSqlDataBase=0 then
		If Session("NewsTotalNum")="" then
			Session("NewsTotalNum") = Conn.Execute("Select count(*) from FS_News where AuditTF=1 and AddDate>=#" & FromDate & "# And AddDate<=#" & TentDate & "# and DelTF=0 and HeadNewsTF=0")(0)
		End If
	Else
		If Session("NewsTotalNum")="" then
			Session("NewsTotalNum") = Conn.Execute("Select count(*) from FS_News where AuditTF=1 and AddDate>='" & FromDate & "' And AddDate<='" & TentDate & "' and DelTF=0 and HeadNewsTF=0")(0)
		End if
	End If
	FromDate=left(FromDate,instr(FromDate," ")-1)
	TentDate=left(TentDate,instr(TentDate," ")-1)
elseif Types = "NewType" then
	if CInt(NewsNo) < CInt(NewNum) then
		if AlreadyRefreshID = "" then
			SearchSql = "Select Top 1 * from FS_News where AuditTF=1 and DelTF=0 and HeadNewsTF=0 order by ID desc"
		else
			SearchSql = "Select Top 1 * from FS_News where ID<" & AlreadyRefreshID & " and AuditTF=1 and DelTF=0 and HeadNewsTF=0 order by ID desc"
		end if
	else
		SearchSql = "Select * from FS_News where 1=0"
	end if
		If Session("NewsTotalNum")="" then
			Session("NewsTotalNum") = Conn.Execute("Select count(id) from FS_News where AuditTF=1 and DelTF=0 and HeadNewsTF=0")(0)
			If Session("NewsTotalNum")>NewNum then Session("NewsTotalNum")=NewNum
		End if
elseif Types = "AllType" then
	if AlreadyRefreshID = "" then
		SearchSql = "Select top 1 * from FS_News where AuditTF=1 and DelTF=0  and HeadNewsTF=0 Order by ID"
	else
		SearchSql = "Select top 1 * from FS_News where ID>" & AlreadyRefreshID & " and AuditTF=1 and DelTF=0  and HeadNewsTF=0 Order by ID"
	end if
	If Session("NewsTotalNum")="" then
		Session("NewsTotalNum") = Conn.Execute("Select count(*) from FS_News where AuditTF=1 and DelTF=0  and HeadNewsTF=0")(0)
	End If
elseif Types = "ClassType" then
		if ClassID <> "" then
			If Instr(1,ClassID,",")=0 then
				If Request("IssueSubClass")="IssueSubClass" then
					If Session("AllClassID")="" then
						Session("AllClassID")="'" & ClassID & "'" & AllChildClassIDList(ClassID)
					End If
				else
					If Session("AllClassID")="" then
						Session("AllClassID")="'"&ClassID&"'"
					End If
				End If
			Else
				If Session("AllClassID")="" then
					Session("AllClassID")="'" & replace(replace(ClassID,",","','")," ","") & "'"
				End If
			End If
				if AlreadyRefreshID = "" then
					SearchSql = "Select Top 1 * from FS_News where AuditTF=1 and ClassID in(" & Session("AllClassID") & ") and DelTF=0 and HeadNewsTF=0 order by id"
				else
					SearchSql = "Select Top 1 * from FS_News where ID>" & AlreadyRefreshID & " and AuditTF=1 and ClassID in(" & Session("AllClassID") & ") and DelTF=0 and HeadNewsTF=0 order by id"
				end if
			If Session("NewsTotalNum")="" then
				Session("NewsTotalNum") = Conn.Execute("Select count(*) from FS_News where AuditTF=1 and ClassID in(" & Session("AllClassID") & ") and DelTF=0 and HeadNewsTF=0")(0)
			End If
		else
			SearchSql = ""
			Session("NewsTotalNum") = 0
		end if
else
	SearchSql = ""
	Session("NewsTotalNum") = 0
end if
if SearchSql <> "" then
	Set RsSearchObj = Server.CreateObject(G_FS_RS)
	RsSearchObj.Open SearchSql,Conn,1,1
	if RsSearchObj.Eof then
		PromptInfo = "刷新新闻成功<font color=red><b>" & Session("NewsTotalNum") & "</b></font>条新闻<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Session("NewsTotalNum")=""
		Call PromptFunction
	else
		RefreshNews RsSearchObj
		NewsNo = NewsNo + 1
		AlreadyRefreshID = RsSearchObj("ID")
		Response.Write("<meta http-equiv=""refresh"" content=""0;url=RefreshNewsSave.asp?NumClass=" & NumClass & "&NewsNo=" & NewsNo & "&NewNum=" & NewNum & "&FromDate=" & FromDate & "&TentDate=" & TentDate & "&ClassID=" & ClassID & "&AlreadyRefreshID=" & AlreadyRefreshID & "&Types=" & Types & "&IssueSubClass="&Request("IssueSubClass")& """>")
		PromptInfo = "共有<font color=red><b>" & Session("NewsTotalNum") & "</b></font>条新闻需要刷新<br><br>正在刷新第<font color=red><b>" & NewsNo & "</b></font>条新闻"
		PromptInfo = PromptInfo & "按确定键返回！<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
		Call PromptFunction
	end if
	Set RsSearchObj = Nothing
else
	PromptInfo = "刷新新闻成功<font color=red><b>" & Session("NewsTotalNum") & "</b></font>条新闻<br><br><input name=""imageField"" type=""image"" src=""../../Images/Btn_Back.gif"" width=""75"" height=""21"" border=""0"">"
	Session("NewsTotalNum")=""
	Call PromptFunction
end if

Function AllChildClassIDList(ClassID)
	Dim TempRs
	Set TempRs = Conn.Execute("Select ClassID,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc")
	do while Not TempRs.Eof
		AllChildClassIDList = AllChildClassIDList & ",'" & TempRs("ClassID") & "'"
		AllChildClassIDList = AllChildClassIDList & AllChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

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
  <form method=post action=RefreshNews.asp>
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