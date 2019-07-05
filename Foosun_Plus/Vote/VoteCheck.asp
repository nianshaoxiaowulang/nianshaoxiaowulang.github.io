<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/NoSqlHack.asp" -->
<%
  dim DBC,conn
  set DBC=new databaseclass
  set conn=DBC.openconnection()
  set DBC=nothing

if isnull(request.cookies("ClickNums")) or request.cookies("ClickNums")="" then
response.cookies("ClickNums")="|"
response.cookies("ClickNums").Expires=Date+365
end if

Dim CloseTFObj,i

Set CloseTFObj = Conn.Execute("select State,EndTime from FS_Vote where VoteID='"&request("VoteID")&"'")
	if CloseTFObj("State")<>"1" then
		Response.Write("<script>alert(""这个项目已经关闭或是过期"");</script>")  
		call show()
		Response.End
		CloseTFObj.close
	end if

	if CloseTFObj("EndTime")<>"0"then 
		if datediff("s",now,formatdatetime(CloseTFObj("EndTime")))<0 then
			Conn.Execute("Update FS_Vote set State=2 where VoteID='"&request("VoteID")&"'")
			Response.Write("<script>alert(""这个项目已经过期"");</script>")  
			call show() 
			Response.End 
			CloseTFObj.close 
		end if
	end if
	CloseTFObj.close

	if instr(request.cookies("ClickNums"),"|"&request("VoteID")&"|")<>0 then 
		Response.Write("<script>alert(""这个项目你已经投过了"");</script>")  
		call show()
		response.end
	end if

	if request("style")="radio" then 
		if isnull(request("voted")) or request("voted")="" then 
			Response.Write("<script>alert(""请选择投票项目"");window.close();</script>")  
			response.end
		end if
		Conn.Execute("Update FS_VoteOption set ClickNum=ClickNum+1 where ID="&request("voted")&"")
	elseif request("style")="checkbox" then
		if request("voted").count=0 then
			Response.Write("<script>alert(""请选择投票项目"");window.close();</script>")  
			response.end
		end if
		for i=1 to request("voted").count
			Conn.Execute("Update VoteOption set ClickNum=ClickNum+1 where ID="&request("voted")(i)&"")
		next
	end if
	response.cookies("ClickNums")=request.cookies("ClickNums")&request("VoteID")&"|"
	response.cookies("ClickNums").Expires=Date+365
	call show()

	Function show
		response.write "<script>location.href=""VoteResult.asp?VoteID="&request("VoteID")&""";</script>"
		response.end
	End Function
	Set Conn=nothing
%>