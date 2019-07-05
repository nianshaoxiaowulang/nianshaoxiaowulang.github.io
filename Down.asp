<% Option Explicit %>
<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<%
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
'==============================================================================
Dim DBC,Conn
Set DBC=new DataBaseClass
Set Conn=DBC.OpenConnection
Set DBC=Nothing
Dim ConfigDoMain
Set ConfigDoMain=conn.execute("select domain,IndexExtName from FS_config")
Dim ResponseBodyStr,ResponseStr,ErrorStr,RsAddressObj,FileURL
Dim Server_Name,Server_V1,Server_V2
Dim OnlyFileUrlTF '只有文件地址
OnlyFileUrlTF = False
ResponseBodyStr = "<title>下载</title>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<style>body{font-size:9pt;line-height:140%}</style>" & Chr(13)
ResponseBodyStr = ResponseBodyStr & "<body>" & Chr(13)
ErrorStr = "<meta http-equiv='Refresh' content='5; URL="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&"'>" & Chr(13)
ErrorStr = ErrorStr & ResponseBodyStr & Chr(13)
ErrorStr = ErrorStr & "<b>错误!&nbsp;</b>读取地址时出错&nbsp;5秒后自动<a href="&ConfigDoMain("domain")&"/index."&ConfigDoMain("IndexExtName")&">返回首页</a>..." & Chr(13)
FileURL = Request("FileUrl")
if Request("ID")="" And FileURL = "" then
	Response.Write ErrorStr
	Set Conn = Nothing
	Response.End
end if
if FileUrl = "" then
	Set RsAddressObj = Server.CreateObject(G_FS_RS)
	RsAddressObj.Open "Select Url from FS_DownLoadAddress where ID='" & trim(Replace(Replace(Request("ID"),"'",""),Chr(39),"")) & "'",Conn,1,1
	if Not RsAddressObj.Eof then
		FileURL = RsAddressObj("Url")
	else
		RsAddressObj.Close
		Set RsAddressObj = Nothing
		Set Conn = Nothing
		Response.Write ErrorStr
		Response.End
	end if
	RsAddressObj.Close
	OnlyFileUrlTF = False
else
	OnlyFileUrlTF = True
end if
'防盗链
Dim DownLoadConfigObj
Set DownLoadConfigObj = Conn.Execute("Select * from FS_DownLoadConfig")
if DownLoadConfigObj("Lock") = 1 then
	Server_Name = Len(Request.ServerVariables("SERVER_NAME"))
	Server_V1 = Left(Replace(Cstr(Request.ServerVariables("HTTP_REFERER")),"http://",""),Server_Name)
	Server_V2 = Left(Cstr(Request.ServerVariables("SERVER_NAME")),Server_Name)
	if Server_V1 <> Server_V2 and Server_V1 <> "" and Server_V2 <> "" then
		Set DownLoadConfigObj = Nothing
		Set Conn = Nothing
		Response.write("<script>alert('没有权限');history.back();</script>")
		Response.End
	end if
end if
'判断IP限制
Dim RequestIPAddress,IPList,IPType,Flag,DownLoadTF
RequestIPAddress = Request.ServerVariables("REMOTE_ADDR")
IPList = DownLoadConfigObj("IPList")
IPType = DownLoadConfigObj("IPType")
Flag = CheckIPAddress(IPList,RequestIPAddress)
'Response.Write(Flag)
'Response.End
if Not IsNull(IPList) And IPList <> "" then
	if Flag = True then
		if IPType = 1 then 
			DownLoadTF = False
		else
			DownLoadTF = True
		end if
	else
		if IPType = 1 then 
			DownLoadTF = True
		else
			DownLoadTF = False
		end if
	end if
else
	DownLoadTF = True
end if

if DownLoadTF then
	if OnlyFileUrlTF = False then
		RsAddressObj.Open "Select ClickNum from FS_DownLoad where DownLoadID='" & trim(Replace(Replace(Request("DownID"),"'",""),Chr(39),"")) & "'",Conn,1,2
		if Not RsAddressObj.eof then
			RsAddressObj("ClickNum") = CLng(RsAddressObj("ClickNum")) + 1
			RsAddressObj.UpDate
		else
			RsAddressObj.Close
			Set RsAddressObj = Nothing
			Set Conn = Nothing
			Response.Write ErrorStr
			Response.End
		end if
	end if
	Set RsAddressObj = Nothing
	if InStr(LCase(FileURl),"http://") = 0 then
		FileURl = ConfigDoMain("domain") & FileUrl
	end if
	Response.Redirect FileURL
else
	Response.write("<script>alert('没有权限,或者IP被锁定');history.back();</script>")
end if
Response.End

Set DownLoadConfigObj = Nothing
Set Conn = Nothing
Set ConfigDoMain = Nothing
Function CheckIPAddress(IPList,IPAddress)
	Dim TempArray,i,j,AddressArray,BeginAddressArray,EndAddressArray,IPAddressArray
	IPAddressArray = Split(IPAddress,".")
	if UBound(IPAddressArray) <> 3 then
		CheckIPAddress = False
		Exit Function
	end if
	if IsNull(IPList) then
		CheckIPAddress = False
	else
		if IPList <> "" then
			TempArray = Split(IPList,"$")
			for i = LBound(TempArray) to UBound(TempArray)
				AddressArray = Split(TempArray(i),"-")
				if UBound(AddressArray) = 1 then
					BeginAddressArray = Split(AddressArray(0),".")
					EndAddressArray = Split(AddressArray(1),".")
					if (UBound(BeginAddressArray) = 3) and (UBound(EndAddressArray) = 3) then
						for j = LBound(BeginAddressArray) to UBound(BeginAddressArray)
								'Response.Write(EndAddressArray(j) = BeginAddressArray(j))
							if (EndAddressArray(j) = BeginAddressArray(j)) then
								if EndAddressArray(j) <> IPAddressArray(j) then
									if (CInt(IPAddressArray(j)) >= CInt(BeginAddressArray(j))) And (CInt(IPAddressArray(j)) <= CInt(EndAddressArray(j))) then
										CheckIPAddress = True
										Exit Function
									end if
								end if
							else
								if (CInt(IPAddressArray(j)) >= CInt(BeginAddressArray(j))) And (CInt(IPAddressArray(j)) <= CInt(EndAddressArray(j))) then
									CheckIPAddress = True
									Exit Function
								end if
							end if
						Next
					end if
				end if
				'Response.End
			Next
			CheckIPAddress = False
		else
			CheckIPAddress = False
		end if
	end if
End Function
%>