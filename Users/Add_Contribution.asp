<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
Dim RsUserObj
Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(session("MemName"),"'","")&"' and Password = '"& Replace(session("MemPassword"),"'","") &"'")
If RsUserObj.eof then
    Response.Write("<script>alert(""���ش���"&CopyRight&""");location=""Login.asp"";</script>")  
    Response.End
Else
	Dim TempStr,PopManage
	TempStr = ClassList	
	Function ClassList()
		Dim Rs,Temp
		Set Rs = server.createobject(G_FS_RS)
		Rs.source = "select * from FS_NewsClass where ParentID='0' and Contribution=1"
		Rs.open Rs.source,Conn,1,1
		If Rs.Eof then
			ClassList = "<font color=red>��ʱû������Ͷ�����Ŀ</font>"
			Exit Function	
		End if
		Do while Not Rs.Eof
			Temp = ""
			ClassList = ClassList & ""
			if Rs("ChildNum") = 0 then
				Temp = Temp & " �� <img src=images/NewsPr1.gif> "
			else
				Temp = Temp & " �� <img src=images/NewsPr.gif> "
			end if
			if PopManage=true then
			ClassList = ClassList & Temp & "<A href=Add_ManageAdd.asp?ClassID="&Trim(Rs("ClassID"))&"><font color=red>"&Rs("ClassCName")&"</font></A>" & "<br>"
			else
			ClassList = ClassList & Temp & "<A href=Add_UserAdd.asp?ClassID="&Trim(Rs("ClassID"))&"><font color=red>"&Rs("ClassCName")&"</font></A>" & "<br>"
			end if
			ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
			Rs.MoveNext	
		loop
		Rs.Close
		Set Rs = Nothing
	End Function
	Function ChildClassList(ClassID,Temp)
		Dim TempRs,TempStr
		TempStr = Temp & " �� "
		Set TempRs = Conn.Execute("Select * from FS_NewsClass where ParentID = '" & ClassID & "' and Contribution=1")
		do while Not TempRs.Eof
			if TempRs("ChildNum") = 0 then
				if PopManage=true then
				ChildClassList = ChildClassList & TempStr & "�� <img src=images/NewsPr1.gif> " & "<A href=Add_ManageAdd.asp?ClassID="&Trim(TempRs("ClassID"))&"><font color=red>"&TempRs("ClassCName")&"</font></A>" & "<br>"
				else
				ChildClassList = ChildClassList & TempStr & "�� <img src=images/NewsPr1.gif> " & "<A href=Add_UserAdd.asp?ClassID="&Trim(TempRs("ClassID"))&"><font color=red>"&TempRs("ClassCName")&"</font></A>" & "<br>"
				end if
			else
				if PopManage=true then
				ChildClassList = ChildClassList & TempStr & "�� <img src=images/NewsPr.gif> " & "<A href=Add_ManageAdd.asp?ClassID="&Trim(TempRs("ClassID"))&"><font color=red>"&TempRs("ClassCName")&"</font></A>" & "<br>"
				else
				ChildClassList = ChildClassList & TempStr & "�� <img src=images/NewsPr.gif> " & "<A href=Add_UserAdd.asp?ClassID="&Trim(TempRs("ClassID"))&"><font color=red>"&TempRs("ClassCName")&"</font></A>" & "<br>"
				end if
			end if
			ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
			TempRs.MoveNext
		loop
		TempRs.Close
		Set TempRs = Nothing
	End Function
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
		  <TD height="262" vAlign=top> 
			<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
					  border=0>
			  <TBODY>
				<TR> 
				  <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
					  <TBODY>
						<TR> 
						  <TD width=26><IMG 
								  src="images/Favorite.OnArrow.gif" border=0></TD><TD class=f4>Ͷ��</TD>
						</TR>
					  </TBODY>
					</TABLE></TD>
				</TR>
				<TR> 
				  <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
					  <TBODY>
						<TR> 
						  <TD bgColor=#ff6633 height=4><IMG height=1 src="" width=1></TD>
						</TR>
					  </TBODY>
					</TABLE></TD>
				</TR>
				<TR> 
					
				  <TD width="100%" height="238" valign="top"> 
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
							  
							<TD height="233" vAlign=top> 
							  <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
					  background="" border=0>
								<TBODY>
								  <TR> 
									<TD width="15%" height="26"> 
									  <div align="left"> <font color="#000000"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="Add_Contribution.asp"><font color="#FF0000">��ҪͶ��</font></a> 
										</font> </div></TD>
									<TD width="17%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution.asp">δ���Ͷ��</a></TD>
									<TD width="43%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution_Passed.asp">�����Ͷ��</a></TD>
									<TD width="25%"> 
									  <div align="center"></div></TD>
								  </TR>
								</TBODY>
							  </TABLE>
							  <hr size="1" noshade>
							  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#E6E6E6">
								<tr bgcolor="#FFFFFF"> 
								  <td width="100%" height="57" >ѡ��Ͷ����Ŀ��<br>
									<br>
									<% =TempStr %> </td>
								</tr>
							  </table> </TD>
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
End If
%>
<script language="JavaScript" type="text/JavaScript">
function deltf(){if(confirm("��ȷ��Ҫɾ��?")){return true;}return false;}
</script>

