<!--#include file="inc_common.asp"-->
<!--#include file="UBB.asp"-->
<body topmargin="0" leftmargin="0">
<%
dim n
	n=10		'nΪҪ��ʾ�������������Լ��޸�֮
sql="Select top "&n&" * from [topic] where checked=1 and whisper=0 order by usertime desc"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

if rs.eof and rs.bof then
	Response.Write "û�����ԡ�"
	rs.close
	set rs=nothing
else
	dim usertitle
	while not rs.eof
	if rs("usertitle")="" then
		usertitle="�ޱ���"
	else
		usertitle=HTMLencode(rs("usertitle"))
	end if
	Response.Write "<A HREF='index.asp' title='����鿴����' target='_parent'>"&usertitle&"</A>--"&rs("usertime")&"<br>"
	rs.movenext
	wend
	rs.close
	set rs=nothing
end if
%>