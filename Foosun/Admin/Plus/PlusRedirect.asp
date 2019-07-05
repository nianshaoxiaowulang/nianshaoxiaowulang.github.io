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
if Not JudgePopedomTF(Session("Name"),"P080600") then Call ReturnError()
Dim RsPlusObj
Set RsPlusObj = Conn.Execute("Select * from FS_Plus where ID=" & Request("id") )
if Not RsPlusObj.Eof then
	Response.Redirect("" & RsPlusObj("Link") & "")
else
	Response.Write("没有发现插件地址")
end if
Set RsPlusObj = Nothing
Set Conn = Nothing
%>
