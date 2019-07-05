<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim Conn,DBC
Set DBC = New DatabaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing

Dim AddrConn,AddrDBC,AddrConnStr
Set AddrDBC = New DatabaseClass
	AddrDBC.ConnStr = "DBQ=" + Server.MapPath(""&IPDataBaseConnStr&"") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set AddrConn = AddrDBC.OpenConnection()
Set AddrDBC = Nothing

Dim DummyPath
If SysRootDir <> "" then
	DummyPath = "/"& SysRootDir
Else
	DummyPath = ""
End If

	Dim Types,VisitIP,vSoft,vExplorer,vOS,EnVisitIP,RsCouObj,RsCouSql,vSource,ExpTime
	ExpTime = 24
	Types = Request("Type")
	VisitIP = request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If VisitIP = "" then
		VisitIP = request.ServerVariables("REMOTE_ADDR")
	End If
	vSoft = Request.ServerVariables("HTTP_USER_AGENT")	
	if vSource="" then
		vSource="直接输入网址进入的"
	else	
		vSource=Mid(vSource,8)
		vSource="http://"&Mid(vSource,1,instr(vSource,"/"))
	end if
	
	if instr(vSoft,"NetCaptor") then											
		vExplorer="NetCaptor"
	elseif instr(vSoft,"MSIE 6") then
		vExplorer="Internet Explorer 6.x"
	elseif instr(vSoft,"MSIE 5") then
		vExplorer="Internet Explorer 5.x"
	elseif instr(vSoft,"MSIE 4") then
		vExplorer="Internet Explorer 4.x"
	elseif instr(vSoft,"Netscape") then
		vExplorer="Netscape"
	elseif instr(vSoft,"Opera") then
		vExplorer="Opera"
	else
		vExplorer="Other"
	end if
	
	if instr(vSoft,"Windows NT 5.0") then										
		vOS="Windows 2000"
	elseif instr(vSoft,"Windows NT 5.1") then
		vOS="Windows XP"
	elseif instr(vSoft,"Windows NT 5.2") then
		vOS="Windows 2003"
	elseif instr(vSoft,"Windows NT") then
		vOS="Windows NT"
	elseif instr(vSoft,"Windows 9") then
		vOS="Windows 9x"
	elseif instr(vSoft,"unix") or instr(vSoft,"linux") or instr(vSoft,"SunOS") or instr(vSoft,"SunOS") or instr(vSoft,"BSD") or instr(vSoft,"Mac") then 
		vOS="Unix & Unix 类"
	else
		vOS="Other"
	end if
	EnAddress = EnAddr(EnIP(VisitIP))
	Set RsCouObj = Conn.Execute("Select ID from FS_FlowStatistic where IP='"&VisitIP&"'")
	If RsCouObj.eof then
		Response.Cookies("online") = false
	End if
	RsCouObj.Close
	Set RsCouObj = Nothing
	
	If Types = "Word" then
		If request.Cookies("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select * from FS_FlowStatistic where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = vOS
			RsCouObj("ExploreType") = vExplorer
			RsCouObj("IP") = Request.ServerVariables("Remote_Addr")
			RsCouObj("OSType") = vOS
			RsCouObj("Area") = EnAddress
			RsCouObj("Source") = vSource
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_FlowStatistic Set LoginNum=LoginNum+1 where IP='"&VisitIP&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
	Set TempObj = Conn.Execute("Select WebCountTime from FS_WebInfo")
	If IsSqlDataBase=0 then
		Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where VisitTime>#"&TempObj("WebCountTime")&"#")
	Else
		Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where VisitTime>'"&TempObj("WebCountTime")&"'")
	End if
		VisitAllNums = Clng(TempObjs(0))
	Set TempObjs = Conn.Execute("Select Count(ID) from FS_FlowStatistic where day(VisitTime) = '"&Day(Now())&"' and month(VisitTime)='"&Month(Now())&"' and year(VisitTime)='"&Year(Now())&"'")
		VisitTodayNum = Clng(TempObjs(0))
	TempObjs.Close
	Set TempObjs = Nothing
	TempObj.Close
	Set TempObj = Nothing
	ShowStr = "总访问量: " & VisitAllNums & " &nbsp;今日访问: " & VisitTodayNum&""
	ElseIf Types = "Pic" then
		If request.Cookies("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select * from FS_FlowStatistic where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = vOS
			RsCouObj("ExploreType") = vExplorer
			RsCouObj("IP") = Request.ServerVariables("Remote_Addr")
			RsCouObj("OSType") = vOS
			RsCouObj("Area") = EnAddress
			RsCouObj("Source") = vSource
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_FlowStatistic Set LoginNum=LoginNum+1 where IP='"&VisitIP&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
	ShowStr = "<img src='"&DummyPath&"/"&PlusDir&"/Count/Img/mc.gif' border=0>"
	Else
		If request.Cookies("online") <> "true" then
			Set RsCouObj = Server.CreateObject(G_FS_RS)
			RsCouSql = "Select * from FS_FlowStatistic where 1=0"
			RsCouObj.Open RsCouSql,Conn,3,3
			RsCouObj.AddNew
			RsCouObj("VisitTime") = Now()
			RsCouObj("OSType") = vOS
			RsCouObj("ExploreType") = vExplorer
			RsCouObj("IP") = Request.ServerVariables("Remote_Addr")
			RsCouObj("OSType") = vOS
			RsCouObj("Area") = EnAddress
			RsCouObj("Source") = vSource
			RsCouObj("LoginNum") = "1"
			RsCouObj.Update
			RsCouObj.Close
			Set RsCouObj = Nothing
		Else
			Conn.Execute("Update FS_FlowStatistic Set LoginNum=LoginNum+1 where IP='"&VisitIP&"' and day(VisitTime)='"&day(now())&"' and month(VisitTime)='"&month(now())&"' and year(VisitTime)='"&year(now())&"'")
		End If
		ShowStr = ""
	End If
	Response.Cookies("online") = "true"
	Response.Cookies("online").Expires = DateAdd("h", ExpTime, now()) 
	
	Response.Write "document.write(" & chr(34) & ShowStr & chr(34) & ")"
	
function EnIP(ip)
	ip=cstr(ip)
	ip1=left(ip,cint(instr(ip,".")-1))
	ip=mid(ip,cint(instr(ip,".")+1))
	ip2=left(ip,cint(instr(ip,".")-1))
	ip=mid(ip,cint(instr(ip,".")+1))
	ip3=left(ip,cint(instr(ip,".")-1))
	ip4=mid(ip,cint(instr(ip,".")+1))
	EnIP=cint(ip1)*256*256*256+cint(ip2)*256*256+cint(ip3)*256+cint(ip4)
end function

Function EnAddr(IP)
	Dim EnAddrObj
    Set EnAddrObj = AddrConn.Execute("select Country,City from Address where StarIP <= "&IP&" and EndIP >= "&IP&"")
	if Not EnAddrObj.Eof then
		EnAddr = EnAddrObj("Country")&EnAddrObj("City")
	else
		EnAddr = "未知区域"
	end if
	EnAddrObj.close
	Set EnAddrObj = Nothing
End Function
Set AddrConn = Nothing
Set Conn = Nothing
%>
