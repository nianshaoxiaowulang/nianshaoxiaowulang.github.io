<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/NoSqlHack.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if request("Location")<>"" and isnull(request("Location"))=false then
	Conn.Execute("update FS_Ads set ClickNum=ClickNum+1 where Location="&clng(request("Location"))&"")
	dim ClickObj,Getip,AdsRsObj,AdsRsSql
	
 	Getip=request.ServerVariables("REMOTE_ADDR")
	set AdsRsObj=server.createobject(G_FS_RS)
	AdsRsSql="select * from FS_AdsVisitList"
	AdsRsObj.open AdsRsSql,conn,1,3
	AdsRsObj.AddNew
	AdsRsObj("AdsLocation") = clng(request("Location"))
	AdsRsObj("VisitTime") = now()
	AdsRsObj("VisitIP") = Getip
	AdsRsObj("VisitType") = "1"
	AdsRsObj.update
	AdsRsObj.close
	set AdsRsObj=nothing
	
		Set ClickObj = Conn.Execute("select Url from FS_Ads where Location="&clng(request("Location"))&"")
	    if not ClickObj.eof then
		   Response.Redirect(ClickObj("Url"))
		end if
end if
Set Conn=nothing
%>