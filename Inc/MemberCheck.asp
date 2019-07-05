<%
Dim MemName,PassWordm,MemID,RsMemObj,Urlm,MemIntegral,mGroupID,confimsn,MemberCheckConfigDoMain
set confimsn=conn.execute("select domain,Sitelock from FS_config")
MemberCheckConfigDoMain = confimsn("domain")
Set confimsn = Nothing
if Request.ServerVariables("SERVER_PORT")<>"80" then
	Urlm = "http://"&Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT")& Request.ServerVariables("URL")&"?"&request.QueryString
else
	Urlm = "http://"&Request.ServerVariables("SERVER_NAME")& Request.ServerVariables("URL")&"?"&request.QueryString
end if
MemName = Request.Cookies("Foosun")("MemName")
PassWordm = Request.Cookies("Foosun")("MemPassword")
MemID = Request.Cookies("Foosun")("MemID")
mGroupID = Request.Cookies("Foosun")("GroupID")
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& MemName &"' and password='"&PassWordm&"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
if not RsMemObj.EOF then
      if RsMemObj("Lock")=1 then
         Response.Write("<script>alert(""没有浏览权限，原因：您已经被锁定\n请与系统管理员联系"");location=""javascript:history.back()"";</script>")  
         Response.End
      end if
else
	RsMemObj.Close
	set RsMemObj = nothing
	Response.redirect""&MemberCheckConfigDoMain&"/Users/Login.asp?UrlAddress=" & Urlm
end if
%>