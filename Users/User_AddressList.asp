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
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
If Trim(Request("UserName"))<>"" then
	  Dim MembersObj
	  Set MembersObj = Conn.execute("Select MemName,Name from FS_Members where MemName='"&Replace(Replace(Request("UserName"),"'",""),Chr(39),"")&"'")
	  if MembersObj.eof then
		Response.Write("<script>alert(""�˻�Ա�Ѿ������ڣ����ܱ�����Աɾ���ˣ�"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	  Else
	  	  Dim FrisObj
	  	  Set FrisObj = Conn.execute("Select MemName from FS_Friend where MemName='"&Replace(Replace(Request("UserName"),"'",""),Chr(39),"")&"'")
		  If Not FrisObj.eof then
				Response.Write("<script>alert(""�˻�Ա�Ѿ������ĺ����ˣ�"&CopyRight&""");location=""javascript:history.back()"";</script>")  
				Response.End
		  Else
			  Dim RsaddObj,addSQL
			  Set RsaddObj = server.createobject(G_FS_RS)
			  addSQL = "select * from FS_Friend where 1=0"
			  RsaddObj.open addSQL,conn,1,3
			  RsaddObj.AddNew
			  RsaddObj("FriendName") = MembersObj("MemName")
			  RsaddObj("RealName") = MembersObj("Name")
			  RsaddObj("MemName") = Session("MemName")
			  RsaddObj.Update
			  Set RsaddObj=Nothing
			  Response.Write("<script>alert(""���Ϊ���ѳɹ���"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			  Response.End
		  End if
	  End if
End if
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(Replace(session("MemName"),"'",""),Chr(39),"")&"' and Password = '"& Replace(Replace(session("MemPassword"),"'",""),Chr(39),"") &"'")
	If RsUserObj.eof then
		Response.Write("<script>alert(""���ش���"&CopyRight&""");location=""Login.asp"";</script>")  
		Response.End
	End if
	If Request.Form("action")="dels" Then
		If trim(Request.Form("FriendId"))<>"" Then
			Conn.execute("Delete From FS_Friend Where id in("&Request.Form("FriendId")&")")
			Response.Write("<script>alert(""����ɾ�����ɹ���"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			Response.End
		Else
			Response.Write("<script>alert(""��ѡ��ɾ���ĺ��ѣ�"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			Response.End
		End if
	End If
	If Request.Form("action")="add" Then
		If trim(Request.Form("FriendName"))="" Then
			Response.Write("<script>alert(""����д������"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			Response.End
		End if
		Dim MemberObj,AddFriendObj,Sql
		Set MemberObj=Conn.execute("select * from FS_Members where MemName= '"& Replace(Request.Form("FriendName"),"'","")&"'")
		If MemberObj.EOF Then
			Response.Write("<script>alert(""û�д��û���"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			Response.End
		Else
			Set AddFriendObj = Server.CreateObject(G_FS_RS)
			Sql = "select * from FS_Friend where 1=0"
			AddFriendObj.Open Sql,Conn,1,3
			AddFriendObj.Addnew
			AddFriendObj("FriendName") = Replace(Request.Form("FriendName"),"'","")
			AddFriendObj("RealName") = Replace(Request.Form("RealName"),"'","")
			AddFriendObj("MemName") = Session("MemName")
			AddFriendObj.update
			AddFriendObj.Close
			Set AddFriendObj=nothing
			Response.Write("<script>alert(""������ӳɹ���"&CopyRight&""");location=""User_AddressList.asp"";</script>")  
			Response.End
		End if
	End If
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
class=f4>�����б�</TD>
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
                          
                        <TD height="207" vAlign=top> 
                          <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="6%" height="26"> <div align="left"><font color="#000000"> 
                                    </font> <font color="#000000"> </font> </div>
                                  <a href="User_Message.asp?action=Inbox"><img src="Images/o_inbox.gif" width="40" height="40" border="0"></a> 
                                </TD>
                                <TD width="6%"><a href="User_Message.asp?action=Outbox"><img src="Images/M_outbox.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="6%"><a href="User_Message.asp?action=Recycle"><img src="Images/M_recycle.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="6%"><a href="User_AddressList.asp"><img src="Images/M_address.gif" width="40" height="40" border="0"></a></TD>
                                <TD width="2%"><span class="f41"><a href="User_WriteMessage.asp"><img src="Images/m_write.gif" width="40" height="40" border="0"></a></span></TD>
                                <TD width="68%"><div align="center"></div></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <hr size="1" noshade>
                          
                            <table width="100%" height="89" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                            <form name="form1" method="post" action="">
                              <tr bgcolor="#E7E7E7"> 
                                <td width="40%" height="26">�û���</td>
                                <td width="37%" height="26">����</td>
                                <td width="23%">����</td>
                              </tr>
							  <%
							  Dim RsFriendobj
							  Set RsFriendobj = Conn.execute("select Top 100 * from FS_Friend where memName='"& Session("MemName")&"' order by id desc")
							  Do while not RsFriendobj.eof 
							  %>
                              <tr bgcolor="#FFFFFF"> 
                                <td height="26"><a href=ReadUser.Asp?UserName=<% = RsFriendobj("FriendName") %> target="_blank"><% = RsFriendobj("FriendName") %></a></td>
                                <td><%
								If  trim(RsFriendobj("RealName"))="" then
									Response.Write("----")
								Else
									Response.Write RsFriendobj("RealName")
								End if
								 %></td>
                                <td>
<input name="FriendId" type="checkbox" id="FriendId" value="<% = RsFriendobj("Id") %>"></td>
                              </tr>
							  <%
								  RsFriendobj.movenext
							  Loop
							  %>
                              <tr bgcolor="#FFFFFF"> 
                                <td height="31" colspan="3">
<div align="right"> 
                                    <input name="action" type="hidden" id="action" value="dels">
                                    <input type="submit" name="Submit" value="ɾ������">
                                  </div></td>
                              </tr>
                            </form>
                          </table>
                         
                            
                          <br>
                          <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                            <form name="form2" method="post" action="">
                              <tr bgcolor="#FFFFFF"> 
                                <td width="8%"><font color="#FF3300">���Ӻ���</font></td>
                                <td>�û�ID�� 
                                  <input name="FriendName" type="text" id="FriendName">
                                  ��ע������ 
                                  <input name="RealName" type="text" id="RealName"> 
                                  <input type="submit" name="Submit2" value="�ύ"> 
                                  <input name="action" type="hidden" id="action" value="add"></td>
                              </tr>
                            </form>
                          </table>
                          
                          <p>&nbsp;</p></TD>
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

