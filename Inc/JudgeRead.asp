<!--#include file="Const.asp" -->
<!--#include file="Cls_DB.asp" -->
<%
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="MemberCheck.asp" -->
<%
Dim GroupID,UserPoint
Dim ConfigDoMain
Set ConfigDoMain=conn.execute("select domain,IndexExtName from FS_config")
MemName = Request.Cookies("Foosun")("MemName")
Sub GetGroupID(Val)
	GroupID = Val
End Sub

Sub GetUserPoint(Val)
	UserPoint = Val
End Sub
Dim Rs,Rs1,ReadTF,Rs2
Set Rs=Conn.execute("select * from FS_Members where MemName='" & MemName & "'")
Set Rs1=Conn.execute("select * from FS_MemGroup where ID="&Rs("GroupID")&"")
Set Rs3=Conn.execute("select * from FS_MemGroup where PopLevel="&GroupID&"")
If clng(UserPoint)<>0 then
	If clng(UserPoint)>Rs("Point") Then
		Response.Write "<title>浏览</title>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" 
		Response.Write "<style>body{font-size:9pt;line-height:140%}</style>"
		Response.Write "<body>" 
		Response.Write "<meta http-equiv='Refresh' content='5; URL="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&"'>" & Chr(13)
		Response.Write("你的点数不足！5秒后自动<a href="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&">返回首页</a>...")  
		response.End
	End if
Else
	If Rs3("Point")>Rs("Point") Then
		Response.Write "<title>浏览</title>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" 
		Response.Write "<style>body{font-size:9pt;line-height:140%}</style>"
		Response.Write "<body>" 
		Response.Write "<meta http-equiv='Refresh' content='5; URL="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&"'>" & Chr(13)
		Response.Write("你的点数不足！5秒后自动<a href="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&">返回首页</a>...")  
		response.End
	End if
End if
%>
