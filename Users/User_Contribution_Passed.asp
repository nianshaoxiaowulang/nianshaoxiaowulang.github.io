<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
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
Dim DBC,conn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()

Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
If request("action")="del" then
	  Conn.execute("Delete From FS_Contribution where ContID='"&Replace(Replace(Request("ContID"),Chr(39),""),"'","")&"'")
	  Conn.execute("Update FS_Members set ConNum=ConNum-1 where MemName='"&Replace(session("MemName"),"'","")&"'")
	  Response.Write("<script>alert(""���ɾ���ɹ���"&CopyRight&""");location=""Mycon.asp"";</script>")
	  Response.End
End if
Dim RsUserObj
Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(Replace(session("MemName"),"'",""),Chr(39),"")&"' and Password = '"& Replace(Replace(session("MemPassword"),"'",""),Chr(39),"") &"'")
If RsUserObj.eof then
	Response.Write("<script>alert(""���ش���"&CopyRight&""");location=""Login.asp"";</script>")  
    Response.End
End if
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>Ͷ�����</TD>
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
                    
                  <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=5 width="100%" border=1>
                    <TBODY>
                        <TR> 
                          
                        <TD height="446" vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="15%" height="26"> <div align="left"> 
                                    <font color="#000000"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="Add_Contribution.asp">��ҪͶ��</a> 
                                    </font> </div></TD>
                                <TD width="17%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution.asp">δ���Ͷ��</a></TD>
                                <TD width="43%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution_Passed.asp"><font color="#FF0000">�����Ͷ��</font></a></TD>
                             <TD width="25%"> <div align="center"></div></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <hr size="1" noshade>
                          <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#E6E6E6">
                            <tr bgcolor="#FFFFFF"> 
                              <td width="100%" >
<table width="100%">
<%
Dim RsCon,strpage,select_count,select_pagecount
	strpage=request.querystring("page")
		if len(strpage)=0 then
		strpage="1"
		end if
Dim Rsconfig
Set Rsconfig=conn.execute("select * from FS_config")
Set RsCon = Server.CreateObject (G_FS_RS)
RsCon.Source="select * from FS_News where Author='"& replace(session("MemName"),"'","") &"' order by adddate desc"
RsCon.Open RsCon.Source,Conn,1,1
If RsCon.eof then
	   RsCon.close
	   set RsCon=nothing
	   Response.Write"<TR><TD>û�м�¼��</TD></TR>"
Else
		RsCon.pagesize=20
		RsCon.absolutepage=cint(strpage)
		select_count=RsCon.recordcount
		select_pagecount=RsCon.pagecount
        for i=1 to RsCon.pagesize
		if RsCon.eof then
		exit for
		end if
		dim rsclass
		set rsclass=conn.execute("select * from FS_Newsclass where Classid='"& Rscon("Classid") &"'")
		response.Write("<tr><td width=60>[<font color=#666666>"&year(Rscon("AddDate"))&"-"&month(Rscon("AddDate"))&"-"&day(Rscon("AddDate"))&"</font>]</td><td>��<a href="&Rsconfig("Domain")&rsclass("SaveFilePath")&"/"&rsclass("classename")&"/"&rscon("FileName")&"."&rscon("FileExtName")&" target=_blank>"&Left(Rscon("title"),25)&"</a></td></tr>")
	RsCon.movenext
next
%> 
</table>
<%  response.write"<br>&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
		if int(strpage)>1 then
		    response.Write"&nbsp;<a href=?page=1>��һҳ</a>&nbsp;"
		    response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&">��һҳ</a>&nbsp;"
			end if
			if int(strpage)<select_pagecount then
			response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&">��һҳ</a>"
			response.Write"&nbsp;<a href=?page="& select_pagecount &">���һҳ</a>&nbsp;"
			end if
			response.Write"<br>"
			Rscon.close
			set rscon=nothing
end if
%> </td>
                            </tr>
                          </table></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
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
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>

