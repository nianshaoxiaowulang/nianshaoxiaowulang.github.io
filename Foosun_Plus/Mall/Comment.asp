<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Md5.asp" -->
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
 if Request("Pid")="" Or Not IsNumeric(Request("Pid")) Then
		Response.Write("<script>alert(""����\n����Ĳ���,����"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
end if
Dim TempRsNewsObj,TempFlag,Downloadid,Pid
TempFlag = true
Pid=Replace(Replace(Trim(Request("Pid")),"'",""),Chr(39),"")
If Pid <> "" Then
	Set TempRsNewsObj = Conn.Execute("Select ReviewTF from FS_Shop_Products where id=" & Pid )
	if Not TempRsNewsObj.Eof then
		if cint(TempRsNewsObj("ReviewTF")) = 0 then
			TempFlag = False
		end if
	else
		TempFlag = False
	end if
	if TempFlag = False then
		Response.Write("<script>alert(""����\n����Ʒ����������"");location.href=""javascript:window.close()"";</script>")
		Response.End
	end if
End if
if request.Form("action")="add" then
	if  request.Form("NoName")="" then
		if request.Form("MemName")="" then
			Response.Write("<script>alert(""����\n����д�����û����������û���ѡ��"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if 
		if request.Form("Password")="" then
			Response.Write("<script>alert(""����\n����д�������룡"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		set Rs = server.CreateObject(G_FS_RS)
		Sql = "select * from FS_Members where MemName='" &Replace(Replace(trim(request("MemName")),"'",""),Chr(39),"")&"' and Password='"&MD5(Replace(Replace(trim(request("Password")),"'","''"),Chr(39),""),16)&"'"
		Rs.Open Sql,Conn,1,3
		if rs.eof then
			Response.Write("<script>alert(""����\nû������û�,�������������������д��"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
	    Else
			Session("MemName") = Rs("MemName")
			Session("MemPassWord") = Rs("Password")
			Session("MemID") = Rs("ID")
			Session("GroupID") = Rs("GroupID")
			Session("Point") = Rs("Point")
			Response.Cookies("Foosun")("MemName") = Rs("MemName")
			Response.Cookies("Foosun")("MemPassword") = Rs("Password")
			Response.Cookies("Foosun")("MemID") = Rs("ID")
			Response.Cookies("Foosun")("GroupID") = Rs("GroupID")
			Response.Cookies("Foosun")("Point") = Rs("Point")
			Session("RePassWord") = Replace(Replace(trim(request("Password")),"'","''"),Chr(39),"")
			Dim Rscon
			set Rscon= conn.execute("select NumberContPoint,NumberLoginPoint from FS_Config")
			conn.execute("update FS_members set LoginNum=LoginNum+1,Point=Point+"&clng(Rscon("NumberLoginPoint"))&",LastLoginIP='"&trim(Request.ServerVariables("Remote_ADDR"))&"',LastLoginTime='" & date() & "' where MemName='"&Rs("MemName")&"'")'�û���½һ�Σ�����+1��
			Rscon.close
			set Rscon=nothing
		end if 
	end if
		if request.Form("RevContent")="" then
			Response.Write("<script>alert(""����\n�������������ݣ�"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		if Len(request.Form("RevContent"))>300 then
			Response.Write("<script>alert(""����\n���۲��ܴ���300���ַ���"");location.href=""javascript:history.go(-1)"";</script>")
			Response.End
		end if
		Dim Rscon1
		Set Rscon1= conn.execute("select ReviewShow from FS_Config")
		set Rs1 = server.CreateObject(G_FS_RS)
		Sql1 = "select * from FS_Review where 1=0"
		Rs1.Open Sql1,Conn,1,3
		Rs1.addnew
		if Request.Form("NoName")="" then
			Rs1("UserID")=Replace(request("MemName"),"'","''")
		else
			Rs1("UserID")="����"
		end if
		Rs1("NewsID")=Replace(Request.form("PID"),"'","''")
		Rs1("Types") = 3
		Rs1("Content")=Request.form("RevContent")
		If Rscon1("ReviewShow")=0 then
			Rs1("Audit") = 1
		Else
			Rs1("Audit") = 0
		End if
		Rs1("IP")=Request.ServerVariables("Remote_Addr")
		Rs1("AddTime")=now()
		Rs1("Isv")=1
		Rs1.update
		Response.Redirect("Comment.asp?Pid="& Pid)
		Response.end 
end if
strpage=request.querystring("page")
		if len(strpage)=0 then
		strpage="1"
		end if
Set RsConfigObj = Conn.execute("select SiteName,Copyright,Domain,UseDatePath From FS_Config")
set Rs = server.CreateObject(G_FS_RS)
Sql = "select * from FS_Review where Newsid='" &Pid &"' and  Types = 3  and isv=1 and Audit=1 order by ID desc"
Rs.Open Sql,Conn,1,1
%>
<html>
<title><% = RsConfigObj("SiteName") %>___��Ʒ�û�����</title>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<body bgcolor="#FFFFFF">
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#D7D7D7" class="tabbgcolor">
  <tr class="tabbgcolorliWhite"> 
    <td colspan="2" bgcolor="#FFFFFF"> <TABLE width="100%" border=0 cellpadding="5" cellspacing="0">
        <TBODY>
          <TR> 
            <TD width=26><IMG 
                              src="../../<%=UserDir%>/images/Favorite.OnArrow.gif" border=0></TD>
            <TD 
class=f4>��Ʒ����</TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
        <TBODY>
          <TR> 
            <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
          </TR>
        </TBODY>
      </TABLE></td>
  </tr>
  <tr class="tabbgcolorliWhite"> 
    <td width="78%" colspan="2" bgcolor="#FFFFFF"> <%
if Rs.eof and Rs.bof then
	Response.write "<p align='center'> δ�ҵ�����</p>"
	Else
	rs.pagesize=20
	rs.absolutepage=cint(strpage)
	select_count=rs.recordcount
	select_pagecount=rs.pagecount
	%> <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <%
		 for i=1 to rs.pagesize
		if rs.eof then
		exit for
		end if
		%>
        <tr> 
          <td height="17" colspan="2" bgcolor="#F5F5F5">�����ԣ�<font color="#0000FF"><%=Replace(rs("IP"),Mid(rs("IP"),InstrRev(rs("IP"),".")+1),"**")%></font>��<font color="#FF0000"> 
            <%
		 set Rs2 = server.CreateObject(G_FS_RS)
		 Sql2 = "select * from FS_Members where MemName='" &Replace(rs("Userid"),"'","''")&"'"
		 Rs2.Open Sql2,Conn,1,3
		 if rs("Userid")="����" or rs("Userid")="" then
		     Members="�����û�"
		 else
		     Members= "<a href=../../"& UserDir &"/ReadUser.asp?UserName="&rs2("MemName")&" target=_blank>"&rs("Userid")&"</a>"
		 end if		  
		 %>
            </font><strong> 
            <% = Members%>
            </strong>��<%=rs("AddTime")%>�� ��Ʒ�� 
            <%
			set Rs1 = server.CreateObject(G_FS_RS)
			Sql1 = "select * from FS_Shop_Products where ID=" & Replace(request("PId"),"'","''")
			Rs1.Open Sql1,Conn,1,1
			Dim RsClassObj
			Set RsClassObj = Conn.execute("Select ClassID,ClassEname,SaveFilePath From FS_NewsClass Where ClassID='"& Replace(Replace(Rs1("ClassID"),"'",""),Chr(39),"")&"'")
			%> <a href=<%=RsConfigObj("Domain")&RsClassObj("SaveFilePath")&"/"&RsClassObj("ClassEname")&Rs1("Product_SavPath")&"/"&Rs1("Product_FileName")&"."&Rs1("Product_ExName")&""%> target="_blank"><font color="#FF0000"><%=rs1("Product_Name")%></font></a> <%
			rs1.close
			set rs1=nothing
			Set RsClassObj = Nothing
			%>
            ����ĵ����ۣ�</td>
        </tr>
        <tr> 
          <td height="39" colspan="2" valign="top"> <%
		if conn.execute("select ReviewShow from FS_Config")(0) = 1 then
			if RS("Audit") = 1 then
			  Response.Write(rs("Content"))
			else
			  Response.Write("<font color=""red"">����Ա��û����˴�����,��ʱ����ʾ��</font>")
			end if
		else
			  Response.Write(rs("Content"))
		end if
		  %> </td>
        </tr>
        <%
	  rs.movenext
	 next
	%>
      </table>
      <%
	   response.write"&nbsp;&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
		if int(strpage)>1 then
		   response.Write"&nbsp;&nbsp;&nbsp;<a href=?page=1&Pid="&Request("Pid")&">��һҳ</a>&nbsp;"
		   response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Pid="&Request("Pid")&">��һҳ</a>&nbsp;"
		end if
		if int(strpage)<select_pagecount then
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Pid="&Request("Pid")&">��һҳ</a>"
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="& select_pagecount &"&Pid="&Request("Pid")&">���һҳ</a>&nbsp;"
		end if
		response.Write"<br>"
end if	   
	   %> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> </td>
        </tr>
      </table></td>
  </tr>
  <tr class="tabbgcolorliWhite"> 
    <td colspan="2" bgcolor="#FFFFFF"><form name="form1" method="post" action="">
        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
          <TBODY>
            <TR> 
              <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
            </TR>
          </TBODY>
        </TABLE>
        <table width="100%" border="0" cellpadding="5" cellspacing="1" class="tabbgcolor">
          <tr bgcolor="#F7F7F7"> 
            <td width="15%"> <div align="right"> 
                <input name="Pid" type="hidden" id="Pid" value="<%=trim(Request("Pid"))%>">
                <input name="action" type="hidden" id="action" value="add">
                ��Ա���ƣ�</div></td>
            <td width="85%"> <input name="MemName" type="text" id="MemName" value="<%=session("MemName")%>"> 
              <input name="NoName" type="checkbox" id="NoName" value="1">
              ���� <font color="#FF0000">��</font><a href="../../<%=UserDir%>/sRegister.asp"><font color="#FF0000">ע���û�</font></a> 
              <a href="../../<%=UserDir%>/User_GetPassword.asp">���������룿</a>����<a href="../../<%=UserDir%>/User_Comments.asp"><font color="#0000FF">�鿴�ҵ�����</font></a> 
            </td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td> <div align="right">���룺</div></td>
            <td> <input name="Password" type="password" id="Password" value="<%=Session("RePassWord")%>"> </td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td> <div align="right">�������ݣ�<br>
                (���300���ַ�) </div></td>
            <td> <textarea name="RevContent" cols="60" rows="6" id="RevContent"></textarea></td>
          </tr>
          <tr bgcolor="#F7F7F7"> 
            <td colspan="2" align="center"> <input type="submit" name="Submit" value=" �� �� "> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit2" value=" �� �� "></td>
          </tr>
        </table>
      </form>
      <br>
      �� �������ϵ��£����ء�ȫ���˴�ί�����ά����������ȫ�ľ������͡����������ӹ���������涨�����л����񹲺͹����������йط��ɷ��档 <br>
      �� �Ͻ�����Σ�����Ұ�ȫ���𺦹������桢�ƻ������Žᡢ�ƻ������ڽ����ߡ��ƻ�����ȶ������衢�̰�����������������ݵ���Ʒ �� <br>
      �� �û�����Լ���ʹ�ñ�վ��������е���Ϊ�е��������Σ�ֱ�ӻ��ӵ��µģ��� <br>
      �� ����̳������Ȩ������ɾ�����Ͻ��̳�е��������ݡ� <br>
      �� ���������е����°�Ȩ��ԭ�����ߺͱ�վ��ͬ���У��κ�����Ҫת�����������£���������ԭ�����߻�վ��Ȩ�� <br>
      �� �����ύ�߷��Դ�������������뱾��վ�����޹ء� </td>
  </tr>
  <tr class="tabbgcolorliWhite">
    <td height="15" colspan="2" bgcolor="#EEEEEE"> 
      <div align="center"><% = RsConfigObj("Copyright")%></div></td>
  </tr>
</table>
</body></html>
<%
Set RsConfigObj = Nothing
%>