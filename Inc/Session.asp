<% 
Dim LoginUrl,Confimsn
Set Confimsn=conn.execute("select domain,Sitelock,IndexExtName,NumberContPoint,NumberLoginPoint from FS_config")
If Request.ServerVariables("SERVER_PORT")<>"80" then
	LoginUrl = "http://"&Request.ServerVariables("LOCAL_ADDR")& ":" & Request.ServerVariables("SERVER_PORT")& Request.ServerVariables("URL")&"?"&request.QueryString
Else
	LoginUrl = "http://"&Request.ServerVariables("LOCAL_ADDR")& Request.ServerVariables("URL")&"?"&request.QueryString
End if
Dim UserName,UserPass,UserID,RsSessionObj,TopLocation
UserName = Session("Name")
UserPass = Session("PassWord")
UserID = Session("AdminID")
If G_FS_Session = 0 then
	if Session("Name") = "" OR Session("PassWord") = "" OR Session("AdminID") = "" OR Session("UnRefreshs") = "" then
		TopLocation = "" & ConfiMSN("domain") & "/" & AdminDir & "Login.asp"
		Response.Write("<script>top.location='" & TopLocation & "'</script>")
		Response.end
	End If
ElseIf G_FS_Session = 1 then
	Dim SessionTFObj,SessionSQL
	if Session("AdminID") <> ""  then
		Set SessionTFObj = Server.CreateObject(G_FS_RS)
		SessionSQL = "Select ID,Name,PassWord From FS_Admin where Name = '"& UserName &"' and PassWord='"& UserPass &"' and Id="&Cint(Session("AdminID"))
		SessionTFObj.Open SessionSQL,Conn,1,1
		If SessionTFObj.eof then
			TopLocation = "" & ConfiMSN("domain") & "/" & AdminDir & "Login.asp"
			Response.Write("<script>top.location='" & TopLocation & "'</script>")
			Response.end
		End if
	Else
			TopLocation = "" & ConfiMSN("domain") & "/" & AdminDir & "Login.asp"
			Response.Write("<script>top.location='" & TopLocation & "'</script>")
			Response.end
	End If
Else
	  Response.Write("<script>alert(""错误:\n配置信息出错，请与系统管理员联系\n\n系统即将关闭窗口!"&Copyright&""");window.close();</script>")
      Response.End
End If
if Confimsn("Sitelock")="" then
	  Response.Write("<script>alert(""错误:\n产品信息没有找到\n\n系统即将关闭窗口!"&Copyright&""");window.close();</script>")
      Response.End
end if
%>

