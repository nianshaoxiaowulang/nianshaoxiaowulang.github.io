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
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Option Explicit
Response.buffer=true
NumCode
Function NumCode()
	Response.Expires = -1
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	On Error Resume Next
	Dim zNum,i,j
	Dim Ados,Ados1
	Randomize timer
	zNum = cint(8999*Rnd+1000)
	Session("GetCode") = zNum
	Dim zimg(4),NStr
	NStr=cstr(zNum)
	For i=0 To 3
		zimg(i)=cint(mid(NStr,i+1,1))
	Next
	Dim Pos
	Set Ados=Server.CreateObject("Adodb.Stream")
	Ados.Mode=3
	Ados.Type=1
	Ados.Open
	Set Ados1=Server.CreateObject("Adodb.Stream")
	Ados1.Mode=3
	Ados1.Type=1
	Ados1.Open
	Ados.LoadFromFile(Server.mappath("../Images/Login/body.Fix"))
	Ados1.write Ados.read(1280)
	For i=0 To 3
		Ados.Position=(9-zimg(i))*320
		Ados1.Position=i*320
		Ados1.write ados.read(320)
	Next	
	Ados.LoadFromFile(Server.mappath("../Images/Login/head.fix"))
	Pos=lenb(Ados.read())
	Ados.Position=Pos
	For i=0 To 9 Step 1
		For j=0 To 3
			Ados1.Position=i*32+j*320
			Ados.Position=Pos+30*j+i*120
			Ados.write ados1.read(30)
		Next
	Next
	Response.ContentType = "image/BMP"
	Ados.Position=0
	Response.BinaryWrite Ados.read()
	Ados.Close:set Ados=nothing
	Ados1.Close:set Ados1=nothing
	If Err Then Session("GetCode") = 9999
End Function
%>