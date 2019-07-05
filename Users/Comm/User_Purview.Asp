<% 
Dim LoginUrl,confimsn
Set confimsn=conn.execute("select domain from FS_config")
Dim MemName,PassWordm,MemID,mGroupID,sMemName
If Len(Trim(Request.Cookies("Foosun")("MemPassword")))=16 And session("MemName")="" And Request.Cookies("Foosun")("MemName")<>"" Then 
	session("MemName") = Replace(Request.Cookies("Foosun")("MemName"),"'","")
	Session("MemPassword") = Replace(Request.Cookies("Foosun")("MemPassword"),"'","")
	Session("MemID") = CInt(Request.Cookies("Foosun")("MemID"))
End If
MemName = session("MemName")
sMemName = Session("sName")
PassWordm = Session("MemPassword")
MemID = Session("MemID")
Dim RsMemObj
set RsMemObj = Server.CreateObject (G_FS_RS)
RsMemObj.Source="select * from FS_Members where MemName='"& MemName &"' and password='"&PassWordm&"'"
RsMemObj.Open RsMemObj.Source,Conn,1,1
if not RsMemObj.EOF then
      if cint(RsMemObj("Lock"))=1 then
         Response.Write("<script>alert(""没有浏览权限，原因：您已经被锁定\n请与系统管理员联系"&CopyRight&""");location=""javascript:history.back()"";</script>")  
         Response.End
      end if
else
   RsMemObj.Close
   set RsMemObj = nothing
   Dim TopLocation
   TopLocation = ""&confimsn("domain")&"/"&UserDir&"/Login.asp"
   Response.Write("<script>alert(""您还没登陆，或者登陆过期，请登陆后进行操作!!"&CopyRight&""");location="""& TopLocation &""";</script>")  
   Response.end
End If
Set confimsn=Nothing
Set RsMemObj=Nothing
%>
