<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<% 
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'==============================================================================
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="Inc/Function.asp" -->
<%
set confimsn=conn.execute("select ReviewShow,domain from FS_config")
if Request("Newsid")="" and  Request("Downloadid")="" Then
	Response.Write("����Ĳ���")
	Response.End
end if
Dim Newsid,ReviewList,Content1,RsReview,sql,TempRsNewsObj1,Downloadid
Newsid = Replace(Replace(Trim(Request("Newsid")),"'","''"),Chr(39),"")
Downloadid = Replace(Replace(Trim(Request("Downloadid")),"'","''"),Chr(39),"")
if Newsid<>"" Then
	Set TempRsNewsObj1 = Conn.Execute("Select ShowReviewTF,ReviewTF from FS_News where Newsid='" & Newsid & "'")
	if cint(TempRsNewsObj1("ShowReviewTF")) = 0 then
		response.Write("")
		response.end
	end if
	Set RsReview1 = server.CreateObject (G_FS_RS)
	Sql1 = "select * from FS_Review where Types = 1 and Newsid='" & Newsid &"' and isv=1 order by ID desc"
	RsReview1.Open Sql1,Conn,1,1
	set RsReview = server.CreateObject (G_FS_RS)
	Sql = "select top 10 * from FS_Review where Types = 1 and Newsid='" & Newsid &"' and isv=1 order by ID desc"
	RsReview.Open Sql,Conn,1,3
	ReviewList="<table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""4""  bgcolor=""#E8E8E8""> <tr bgcolor=""#F7F7F7"" align=center><td width=75><strong>���߱���</strong></td><td><strong>������(��<font color=red>"&RsReview1.recordcount&"</font>������)</strong> <a href="&confimsn("domain")&"/NewsReview.asp?Newsid="&Newsid&"><u>�鿴ȫ������</u></a></td><td align=left width=68><strong>��������</strong></td></tr>"
	while Not RsReview.Eof
		if len(RsReview("Content"))>30 then
			Content1=""& RsReview("Content") &".."
		else
			Content1=""& RsReview("Content") &""
		end if
		if confimsn("ReviewShow") = 1 then
			if RsReview("Audit") = 1 then
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td  width=75>"& RsReview("UserID") &"</td><td >"& Content1 &"</td><td  width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			else
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td  width=75>"& RsReview("UserID") &"</td><td ><font color=red>����Ա��û�����</font></td><td  width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			end if
		else
			ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td  width=75>"& RsReview("UserID") &"</td><td >"& Content1 &"</td><td  width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
		end if
		RsReview.movenext
	Wend
	ReviewList=ReviewList & "</table>"
elseif  Downloadid<>"" Then
	Set TempRsNewsObj1 = Conn.Execute("Select ShowReviewTF,ReviewTF from FS_download where Downloadid='" & Downloadid & "'")
	if cint(TempRsNewsObj1("ShowReviewTF")) = 0 then
		response.Write("")
		response.end
	end if
	set RsReview1 = server.CreateObject (G_FS_RS)
	Sql1 = "select * from FS_Review where Types = 2 and Newsid='" & Downloadid &"' and isv=1 order by ID desc"
	RsReview1.Open Sql1,Conn,1,1
	set RsReview = server.CreateObject (G_FS_RS)
	Sql = "select top 10 * from FS_Review where Types = 2 and Newsid='"& Downloadid &"' and isv=1 order by ID desc"
	RsReview.Open Sql,Conn,1,3
	ReviewList="<table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""4""  bgcolor=""#E8E8E8""> <tr bgcolor=""#F7F7F7"" align=center><td width=75><strong>����</strong></td><td><strong>������(��<font color=red>"&RsReview1.recordcount&"</font>������)</strong> <a href="&confimsn("domain")&"/NewsReview.asp?Downloadid="&Downloadid&"><u>�鿴ȫ������</u></a></td><td align=left width=68><strong>��������</strong></td></tr>"
	while Not RsReview.Eof
		if len(RsReview("Content"))>30 then
			Content1=""& Left(RsReview("Content"),30) &".."
		else
			Content1=""& RsReview("Content") &""
		end if
		if confimsn("ReviewShow") = "1" then
			if RsReview("Audit") = "1" then
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td nowrap width=75><a href="&confimsn("Domain")&"/"& UserDir &"/ReadUser.asp?UserName="& RsReview("UserID") &">"& RsReview("UserID") &"</a></td><td nowrap>"&Content1&"</td><td nowrap width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			else
				ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td nowrap width=75><a href="&confimsn("Domain")&"/"& UserDir &"/ReadUser.asp?UserName="& RsReview("UserID") &">"& RsReview("UserID") &"</a></td><td nowrap><font color=red>����Ա��û�����</font></td><td nowrapwidth=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
			end if
		else
			ReviewList=ReviewList & "<tr  bgcolor=""#FFFFFF"" align=center><td nowrap width=75><a href="&confimsn("Domain")&"/"& UserDir &"/ReadUser.asp?UserName="& RsReview("UserID") &">"& RsReview("UserID") &"</a></td><td nowrap>"&Content1&"</td><td nowrap width=100>"&month(RsReview("AddTime"))&"-"&day(RsReview("AddTime"))&" "&Hour(RsReview("AddTime"))&":"&minute(RsReview("AddTime"))&"</td></tr>"
		end if
		RsReview.movenext
	wend
	ReviewList=ReviewList & "</table>"
End if
Response.write "document.write ('"& ReviewList &"')"
RsReview1.close
Set RsReview1=nothing
RsReview.close
Set RsReview=nothing
%>