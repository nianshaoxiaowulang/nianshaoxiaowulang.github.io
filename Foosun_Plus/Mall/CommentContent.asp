<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Function.asp" -->
<%
Dim confimsn
Set confimsn=conn.execute("select ReviewShow,domain from FS_config")
if Request("Pid")="" Then
	Response.Write("")
	Response.End
end if
Dim Pid,ReviewList,Content1,RsReview,sql,TempRsNewsObj1
Pid = Replace(Replace(Trim(Request("Pid")),"'","''"),Chr(39),"")
	Set TempRsNewsObj1 = Conn.Execute("Select ReviewTF from FS_Shop_Products where id=" & Pid)
	if cint(TempRsNewsObj1("ReviewTF")) = 0 then
		response.Write("")
		response.end
	end if
	Set RsReview1 = server.CreateObject (G_FS_RS)
	Sql1 = "select * from FS_Review where Types = 3 and NewsID='" & Pid &"' and isv=1 And Audit=1  order by ID desc"
	RsReview1.Open Sql1,Conn,1,1
	set RsReview = server.CreateObject (G_FS_RS)
	Sql = "select top 10 * from FS_Review where Types = 3 and Newsid='" & Pid &"' and isv=1 And Audit=1 order by ID desc"
	RsReview.Open Sql,Conn,1,3
	ReviewList="<table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""4""  bgcolor=""#E8E8E8""> <tr bgcolor=""#F7F7F7"" align=center><td width=75><strong>作者</strong></td><td><strong>简单内容(共<font color=red>"&RsReview1.recordcount&"</font>个评论)</strong> <a href="&confimsn("domain")&"/"& PlusDir &"/"& MallDir &"/Comment.asp?Pid="&Pid&"><u>查看全部内容</u></a></td><td align=left width=68><strong>发表日期</strong></td></tr>"
	while Not RsReview.Eof
		if len(RsReview("Content"))>30 then
			Content1=""& RsReview("Content") &".."
		else
			Content1=""& RsReview("Content") &""
		end if
			iF  RsReview("UserID") <> "匿名" Then
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td  width=75><a href="&confimsn("Domain")&"/"& UserDir &"/ReadUser.asp?UserName="& RsReview("UserID") &">"& RsReview("UserID") &"</a></td><td >"& Content1 &"</td><td  width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			Else
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td  width=75>"& RsReview("UserID") &"</td><td >"& Content1 &"</td><td  width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			End if
		RsReview.movenext
	Wend
	ReviewList=ReviewList & "</table>"
Response.write "document.write ('"& ReviewList &"')"
RsReview1.close
Set RsReview1=nothing
RsReview.close
Set RsReview=nothing
%>