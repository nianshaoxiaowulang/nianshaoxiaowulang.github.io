<% Option Explicit %>
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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="../Comm/User_Purview.Asp" -->
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=5 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="../images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>���ӹ���</TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
                <TD width="100%" height="159" valign="top"> 
                  <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      
                    <td width="62%"><a href="GBook.asp">�ҷ��������</a> �� <a href="All_GBook.asp">���Ӳ鿴</a> 
                      �� <a href="Write_GBook.asp"><font color="#FF0000">��������</font></a> 
                      �� <a href="GBook.asp?Action=Q">�ѻظ�������</a> �� <a href="GBook.asp?Action=Q"></a><a href="GBook.asp?Action=UnQ">δ�ظ�������</a></td>
                      <form name="form1" method="post" action="ALL_GBook.asp">
                      <td width="38%"><input name="Keyword" type="text" id="Keyword">
                        <input type="submit" name="Submit2" value="����"> </td>
                    </form>
                    </tr>
                  </table>
                  
                <strong>�鿴��������</strong><br>
                <br>
                <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <form method=POST action="GBook.asp" name=Form1  onsubmit="return Cim()">
                    <tr bgcolor="#E8E8E8"> 
                      <td width="6%"> <div align="center"><strong>����</strong></div></td>
                      <td width="42%"><strong>����</strong></td>
                      <td width="23%"><strong>����ʱ��</strong></td>
                      <td width="16%"><strong>�ظ�ʱ��</strong></td>
                      <td width="10%"><strong>������</strong></td>
                    </tr>
     <%
    dim RsCon,strpage,select_count,select_pagecount
	strpage=request.querystring("page")
	if len(strpage)=0 then
		strpage="1"
	end if
	Dim QS
	IF Request("Action")="Q" then
		QS = " and isQ=1"
	ElseIf Request("Action")="UnQ" then
		QS = " and isQ=0"
	Else
		QS = ""
	End if
	Dim Ks
	If Request("Keyword")<>"" then
		Ks = " and (Content like '%"& Replace(Request("Keyword"),"'","") &"%' or Title like '%"& Replace(Request("Keyword"),"'","") &"%')"
	Else
		Ks = ""
	End if
	Set RsCon = Server.CreateObject (G_FS_RS)
	RsCon.Source="select * from FS_GBook where QID=0 "& QS & Ks &" order by Orders,Qtime desc,Addtime desc"
	RsCon.Open RsCon.Source,Conn,1,1
	'Response.Write(RsCon.Source)
	'Response.end
	If RsCon.eof then
		   RsCon.close
		   set RsCon=nothing
		   Response.Write"<TR><TD colspan=""5"" bgcolor=FFFFFF>û�м�¼��</TD></TR>"
	Else
			RsCon.pagesize=15
			RsCon.absolutepage=cint(strpage)
			select_count=RsCon.recordcount
			select_pagecount=RsCon.pagecount
			for i=1 to RsCon.pagesize
				if RsCon.eof then
					exit for
				end if
					  If RsCon("isadmin")=0 then
					  if i mod 2 = 0 then
					%>
					<tr bgcolor="#EEEEEE"> 
					  <%Else%>
					 <tr bgcolor="#FFFFFF"> 
					  <%End If%>
					  <td> <div align="center"><img src="images/face<% = RsCon("FaceNum")%>.gif"></div></td>
						  <td>
							<%If RsCon("Orders")=1 then%>
							<img src="Images/ztop.gif" alt="�̶���" width="18" height="15"> 
							<%Else
								iF RsCon("Isadmin")=1 then%>
								<img src="Images/lhotfolder.gif" alt="����ֻ�й���Ա�ɼ�" width="18" height="12"> 
								<%Else%>
								<img src="Images/hotfolder.gif" alt="һ������" width="18" height="12"> 
							  <%End if
							 End if%>
							<a href="ReadBook.asp?id=<% = RsCon("ID")%>"> </a><a href="ReadBook.asp?id=<% = RsCon("ID")%>"> 
							<% = RsCon("Title")%>
							</a></td>
						  <td> <% = RsCon("Addtime")%> </td>
						  <td> <font color="#FF0000"> 
							<%
							If RsCon("Qtime")=RsCon("Addtime") Or RsCon("Qtime")="" Or RsCon("isQ")=0 then
								Response.Write("")
							Else
								Response.Write RsCon("Qtime")
							End if
							%>
							</font> </td>
						  <td> <%
						  If RsCon("UserID")=0 then
								Response.Write("<font color=#990000>����Ա</font>")
						  Else
								Set MemberObj = Conn.execute("Select MemName From FS_Members Where id="&Replace(Replace(RsCon("UserID"),"'",""),Chr(39),""))
									If Not MemberObj.eof then
										Response.Write("<a href=../ReadUser.Asp?UserName="&MemberObj("MemName")&">"& MemberObj("MemName")&"</a>")
									Else
										Response.Write("�û��ѱ�ɾ��")
									End if
						   End If
						 %></td>
							</tr>
					<%
					ElseIf RsCon("isAdmin")=1 and RsCon("UserID")=Session("MemID") then
					  if i mod 2 = 0 then
					%>
					<tr bgcolor="#EEEEEE"> 
					  <%Else%>
					 <tr bgcolor="#FFFFFF"> 
					  <%End If%>
					  <td> <div align="center"><img src="images/face<% = RsCon("FaceNum")%>.gif"></div></td>
						  <td>
							<%If RsCon("Orders")=1 then%>
							<img src="Images/ztop.gif" alt="�̶���" width="18" height="15"> 
							<%Else
								iF RsCon("Isadmin")=1 then%>
								<img src="Images/lhotfolder.gif" alt="����ֻ�й���Ա�ɼ�" width="18" height="12"> 
								<%Else%>
								<img src="Images/hotfolder.gif" alt="һ������" width="18" height="12"> 
							  <%End if
							 End if%>
							<a href="ReadBook.asp?id=<% = RsCon("ID")%>"> </a><a href="ReadBook.asp?id=<% = RsCon("ID")%>"> 
							<% = RsCon("Title")%>
							</a></td>
						  <td> <% = RsCon("Addtime")%> </td>
						  <td> <font color="#FF0000"> 
							<%
							If RsCon("Qtime")=RsCon("Addtime") Or RsCon("Qtime")=""  Or RsCon("isQ")=0  then
								Response.Write("")
							Else
								Response.Write RsCon("Qtime")
							End if
							%>
							</font> </td>
						  <td> <%
						  If RsCon("UserID")=0 then
								Response.Write("<font color=#990000>����Ա</font>")
						  Else
								Dim MemberObj
								Set MemberObj = Conn.execute("Select MemName From FS_Members Where id="&Replace(Replace(RsCon("UserID"),"'",""),Chr(39),""))
									If Not MemberObj.eof then
										Response.Write("<a href=../ReadUser.Asp?UserName="&MemberObj("MemName")&">"& MemberObj("MemName")&"</a>")
									Else
										Response.Write("�û��ѱ�ɾ��")
									End if
						   End If
						 %></td>
							</tr>
							<%End if
							%>
							<%
							RsCon.MoveNext
						Next
						%>
                  </form>
                </table> 
                  
<%
	Response.write"<br>&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
	if int(strpage)>1 then
		Response.Write"&nbsp;<a href=?page=1&Action="&Request("Action")&">��һҳ</a>&nbsp;"
		Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Action="&Request("Action")&">��һҳ</a>&nbsp;"
	end if
	if int(strpage)<select_pagecount then
		Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Action="&Request("Action")&">��һҳ</a>"
		Response.Write"&nbsp;<a href=?page="& select_pagecount &"&Action="&Request("Action")&">���һҳ</a>&nbsp;"
	end if
	Response.Write"<br>"
	Rscon.close
	Set Rscon=nothing
End if
%> 
</TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> 
      <div align="center">
        <hr size="1" noshade color="#FF6600">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
Set Conn=nothing
%><script language="JavaScript" type="text/JavaScript">
function Cim(){
	if (window.confirm('��ȷ��Ҫ����?')){
	 	return true;
	 } 
	 return false;		
}
</script>