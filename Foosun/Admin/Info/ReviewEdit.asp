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
if Not JudgePopedomTF(Session("Name"),"P010509") then Call ReturnError1()
Dim Action
Action = Request("Action")
Dim NewsID,ReviewID,RsReviewObj,Types
NewsID = Request("NewsID")
ReviewID = Request("ReviewID")
if ReviewID = "" then
	Response.Write("<script>alert('参数传递错误'),history.back();</script>")
	Response.End 
else
	Set RsReviewObj = Server.CreateObject(G_FS_RS)
	RsReviewObj.Open "Select * from FS_Review where ID=" & ReviewID,Conn,3,3
	if RsReviewObj.Eof then
		Set RsReviewObj = Nothing
		Response.Write("<script>alert('修改的评论不存在'),history.back();</script>")
		Response.End 
	else
		Dim RsNewsObj,NewsTitleStr
		Types = RsReviewObj("Types")
		NewsID = RsReviewObj("Newsid")
		if Types = "1" Then
			Set RsNewsObj = Conn.Execute("Select Title from FS_News where NewsID='" & RsReviewObj("NewsID") & "'")
			if Not RsNewsObj.Eof then
				NewsTitleStr = RsNewsObj("Title")
			else
				Set RsReviewObj = Nothing
				Set RsNewsObj = Nothing
				Conn.Execute("Delete from FS_Review where NewsID='" & NewsID & "'")
				Response.Write("<script>alert('评论所在的新闻已经删除'),history.back();</script>")
				Response.End 
			end if
		elseif  Types = "2" Then
			Set RsNewsObj = Conn.Execute("Select Name from FS_Download where DownloadID='" & NewsID & "'")
			if Not RsNewsObj.Eof then
				NewsTitleStr = RsNewsObj("Name")
			else
				Set RsReviewObj = Nothing
				Set RsNewsObj = Nothing
				Conn.Execute("Delete from FS_Review where NewsID='" & NewsID  & "'")
				Response.Write("<script>alert('评论所在的下载已经删除'),history.back();</script>")
				Response.End 
			end if
		End if
		if Action = "Submit" then
			Set RsNewsObj = Nothing
			RsReviewObj("title") = Request("Title")
			RsReviewObj("Content") = Request("Content")
			if Request("Audit") = "1" then
				RsReviewObj("Audit") = 1
			else
				RsReviewObj("Audit") = 0
			end if
			RsReviewObj.Update
			Set RsReviewObj = Nothing
			if Types = "1" Then
				Response.Redirect("Review.asp?NewsID=" & NewsID)
			elseif Types = "2" Then
				Response.Redirect("Review.asp?DownloadID=" & NewsID)
			End if
			Response.End
		end if
	end if
	Set RsNewsObj = Nothing
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>审核评论</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form name="ReviewForm" method="get" action="">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.ReviewForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input type="hidden" name="Action" value="Submit"><input type="hidden" name="NewsID" value="<% = NewsID %>"><input type="hidden" name="ReviewID" value="<% = ReviewID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="120" height="26"><div align="center">标题</div></td>
    <td><input type="text" style="width:100%;" name="Title" value="<% = RsReviewObj("Title") %>"></td>
  </tr>
  <tr>
      <td height="26">
<div align="center">新闻</div></td>
    <td><input disabled type="text" style="width:100%;" name="News" value="<% = NewsTitleStr %>"></td>
  </tr>
  <tr>
      <td height="26">
<div align="center">内容</div></td>
    <td><textarea name="Content" rows="6" style="width:100%;"><% = RsReviewObj("Content") %></textarea></td>
  </tr>
  <tr>
      <td height="26"><div align="center">用户</div></td>
    <td><input disabled style="width:100%;" type="text" name="UserID" value="<% = RsReviewObj("UserID") %>"></td>
  </tr>
  <tr>
      <td height="26">
<div align="center">评论时间</div></td>
    <td><input disabled style="width:100%;" type="text" name="AddTime" value="<% = RsReviewObj("AddTime") %>"></td>
  </tr>
  <tr>
      <td height="26">
<div align="center">IP地址</div></td>
    <td><input disabled style="width:100%;" type="text" name="IP" value="<% = RsReviewObj("IP") %>"></td>
  </tr>
  <tr>
      <td height="26"><div align="center">审核</div></td>
    <td><input type="checkbox" name="Audit" value="1" <% if RsReviewObj("Audit") = 1 then Response.Write("checked") %>></td>
  </tr>
</table>
</form>
</body>
</html>
<%
Set RsReviewObj = Nothing
Set Conn = Nothing
%>