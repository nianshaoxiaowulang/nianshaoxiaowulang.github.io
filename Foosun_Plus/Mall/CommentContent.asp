<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
	ReviewList="<table width=""96%"" border=""0"" cellspacing=""1"" cellpadding=""4""  bgcolor=""#E8E8E8""> <tr bgcolor=""#F7F7F7"" align=center><td width=75><strong>����</strong></td><td><strong>������(��<font color=red>"&RsReview1.recordcount&"</font>������)</strong> <a href="&confimsn("domain")&"/"& PlusDir &"/"& MallDir &"/Comment.asp?Pid="&Pid&"><u>�鿴ȫ������</u></a></td><td align=left width=68><strong>��������</strong></td></tr>"
	while Not RsReview.Eof
		if len(RsReview("Content"))>30 then
			Content1=""& RsReview("Content") &".."
		else
			Content1=""& RsReview("Content") &""
		end if
			iF  RsReview("UserID") <> "����" Then
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