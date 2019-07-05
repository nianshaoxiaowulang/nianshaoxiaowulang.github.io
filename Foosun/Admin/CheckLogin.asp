<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
<!--#include file="../../Inc/Enpas.asp" -->
<%
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Dim DBC,Conn,RsLoginObj,RsLogObj
Dim UserName,UserPass,VerifyCode,System,SqlLog,SqlLogin,Url,TempUserPass
if Request("UrlAddress")<>"" then
	Url = Request("UrlAddress")
else
	Url = "main.asp"
end if
UserName = Replace(Trim(Request.Form("Name")),"'","''")
TempUserPass = Replace(Trim(Request.Form("Password")),"'","''")
VerifyCode = Replace(Trim(Request("VerifyCode")),"'","")
if UserName = "" or  TempUserPass = "" then
	Response.Write("<script>alert(""错误：\n请填写完整"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
end if

if  VerifyCode <> CStr(Session("GetCode"))  then 
	Response.Write("<script>alert(""错误：\n验证码错误"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
end if
if request("VerifyCode") = "" then
	Response.Write("<script>alert(""错误：\n请填写验证码"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
elseif Session("GetCode") = "9999" then
	Session("GetCode")=""
elseif Session("GetCode") = "" then
	Response.Write("<script>alert(""错误：\n请不要重复提交，如需重新登录请返回登录页面。\n返回后请刷新登录页面后重新输入正确的信息"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
elseif cstr(Session("GetCode"))<>cstr(Trim(Request("VerifyCode"))) then
	Response.Write("<script>alert(""错误：\n您输入的确认码和系统产生的不一致，请重新输入。\n返回后请刷新登录页面后重新输入正确的信息"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
end if

System = Request.ServerVariables("HTTP_USER_AGENT")
if Instr(System,"Windows NT 5.2") then
	System = "Win2003"
elseif Instr(System,"Windows NT 5.0") then
	System="Win2000"
elseif Instr(System,"Windows NT 5.1") then
	System = "WinXP"
elseif Instr(System,"Windows NT") then
	System = "WinNT"
elseif Instr(System,"Windows 9") then
	System = "Win9x"
elseif Instr(System,"unix") or instr(System,"linux") or instr(System,"SunOS") or instr(System,"BSD") then
	System = "类Unix"
elseif Instr(System,"Mac") then
	System = "Mac"
else
	System = "Other"
end if
Dim PassArr,TrueResult,CheckedResult
PassArr=split(SafePass,",")

If PassArr(0)=1 then
	If PassArr(3)="1" then
		TrueResult=Trim(Cstr(Cint(mid(Session("GetCode"),Cint(PassArr(1)),1))+Cint(mid(Session("GetCode"),Cint(PassArr(2)),1))))
	Else
		TrueResult=Trim(Cstr(Cint(mid(Session("GetCode"),Cint(PassArr(1)),1))*Cint(mid(Session("GetCode"),Cint(PassArr(2)),1))))
	End If
	If PassArr(4)="0" then
		CheckedResult=left(TempUserPass,Len(TrueResult))
		UserPass=mid(TempUserPass,Len(TrueResult)+1)
	ElseIf Cint(PassArr(4))>len(TempUserPass)-len(TrueResult) then
		CheckedResult=right(TempUserPass,Len(TrueResult))
		UserPass=left(TempUserPass,len(TempUserPass)-Len(TrueResult))
	Else
		CheckedResult=mid(TempUserPass,PassArr(4)+1,Len(TrueResult))
		UserPass=left(TempUserPass,PassArr(4))&mid(TempUserPass,Cint(PassArr(4))+len(TrueResult)+1)
	End If
Else
	UserPass=TempUserPass
End If
Session("GetCode")=""

Set RsLoginObj = server.CreateObject ("ADODB.RecordSet")
SqlLogin = "select * from FS_admin where Name='"&UserName&"' and  password='"&md5(UserPass,16)&"'"
RsLoginObj.Open SqlLogin,Conn,1,1
if Not RsLoginObj.EOF then
	if cint(RsLoginObj("Lock")) = 1 then
		Response.Write("<script>alert(""错误:\n您已经被锁定,请与管理联系\n"&Copyright&""");window.close();</script>")
		Response.End
	end if
	Session("Name") = RsLoginObj("name")
	Session("PassWord") = RsLoginObj("PassWord")
	Session("AdminID") = RsLoginObj("ID")
	Session("GroupID") = CStr(RsLoginObj("GroupID"))
	If Application("UseDatePath")="" or IsNull(Application("UseDatePath"))then
		Application.lock
		Application("UseDatePath")=conn.execute("select UseDatePath from FS_config")(0)
		Application.unlock
	End If
	Dim TempGetPopedomList,RsGroupObj
	if CStr(RsLoginObj("GroupID")) <> "0" then
		Set RsGroupObj = Conn.Execute("Select * from FS_AdminGroup where ID=" & CStr(RsLoginObj("GroupID")))
		if Not RsGroupObj.Eof then
			TempGetPopedomList = RsGroupObj("PopList")
			if IsNull(TempGetPopedomList) then TempGetPopedomList = ""
		else
			TempGetPopedomList = ""
		end if
		Set RsGroupObj = Nothing
	else
		TempGetPopedomList = ""
	end if
	Session("PopedomList") = TempGetPopedomList

	Set RsLogObj = Server.Createobject("adodb.recordset")
	SqlLog = "Select * from FS_Log"
	RsLogObj.open SqlLog,Conn,3,3
	RsLogObj.addnew
	RsLogObj("LogUser") = UserName
	RsLogObj("LogIP") = request.ServerVariables("Remote_Addr")
	RsLogObj("OS") = System
	RsLogObj("Result") = 1
	RsLogObj("LoginTime") = now()
	RsLogObj.update
	RsLogObj.close
	set RsLogObj = Nothing
	If CBool(Request.Form("AutoGet")) or Request.Form("AutoGet")<>"" Then
        Response.Cookies("FoosunCookie")("AdminName")=Session("Name")
        Response.Cookies("FoosunCookie").Expires=Date()+365
    Else
        Response.Cookies("FoosunCookie")("AdminName")=""
        Response.Cookies("FoosunCookie").Expires=Date()-1
    End If
	If TrueResult=CheckedResult then
		Response.Redirect(Url)
		Response.End
	Else
		Response.Write("<script>alert(""错误:\n请检查用户名和密码的正确性\n"&Copyright&""");location.href=""Login.asp"";</script>")
		Response.End
	End If
else
	Set RsLogObj = Server.Createobject("adodb.recordset")
	SqlLog = "Select * from FS_Log"
	RsLogObj.open SqlLog,Conn,3,3
	RsLogObj.AddNew
	RsLogObj("LogUser") = Request.Form("Name")
	RsLogObj("LogIP") = request.ServerVariables("Remote_Addr")
	RsLogObj("OS") = System
	RsLogObj("ErrorPas") = Request.Form("Password")
	RsLogObj("Result") = 0
	RsLogObj("LoginTime") = Now()
	RsLogObj.update
	RsLogObj.close
	set RsLogObj = Nothing
	Response.Write("<script>alert(""错误:\n请检查用户名和密码的正确性\n"&Copyright&""");location.href=""Login.asp"";</script>")
	Response.End
end if
set Conn = Nothing
Set RsLoginObj = Nothing
%>