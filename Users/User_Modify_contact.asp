<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
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
	If Request.Form("action")="Update" then
		If Trim(Request.Form("Name"))="" Or Trim(Request.Form("Tel"))=""  Or Trim(Request.Form("Address"))=""  Then
			Response.Write("<script>alert(""����д������"&CopyRight&""");location=""User_Modify_Contact.asp"";</script>")  
			Response.End
		End if
		Dim RsUpdateObj,SqlUpdate
		Set RsUpdateObj = server.CreateObject (G_FS_RS)
		SqlUpdate = "select * from FS_members where id="& Clng(Replace(Replace(Request.Form("id"),"'",""),Chr(39),""))
		RsUpdateObj.Open SqlUpdate,Conn,1,3
		RsUpdateObj("Name")=NoCSSHackInput(Replace(Request.Form("Name"),"'",""))
		If Replace(Request.Form("Sex"),"'","")="0" Then
			RsUpdateObj("Sex")=0
		Else
			RsUpdateObj("Sex")=1
		End if
		RsUpdateObj("Telephone")=NoCSSHackInput(Replace(Request.Form("tel"),"'",""))
		RsUpdateObj("Msn")=NoCSSHackInput(Replace(Request.Form("Msn"),"'",""))
		RsUpdateObj("Oicq")=NoCSSHackInput(Replace(Request.Form("qq"),"'",""))
		RsUpdateObj("Province")=NoCSSHackInput(Replace(Request.Form("Province"),"'",""))
		RsUpdateObj("City")=NoCSSHackInput(Replace(Request.Form("City"),"'",""))
		RsUpdateObj("Address")=NoCSSHackInput(Replace(Request.Form("address"),"'",""))
		RsUpdateObj("postcode")=NoCSSHackInput(Replace(Request.Form("postcode"),"'",""))
		RsUpdateObj("HomePage")=NoCSSHackInput(Replace(Request.Form("HomePage"),"'",""))
		RsUpdateObj("Birthday")=NoCSSHackInput(Replace(Request.Form("Birthday"),"'",""))
		RsUpdateObj.Update
		RsUpdateObj.Close
		Set RsUpdateObj=Nothing
		Response.Write("<script>alert(""������ϵ���ϳɹ���"&CopyRight&""");location=""User_Modify_Contact.asp"";</script>")  
		Response.End
	End if
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(session("MemName"),"'","")&"' and Password = '"& Replace(session("MemPassword"),"'","") &"'")
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
class=f4>�޸���ϵ��ʽ</TD>
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
              <form method=POST action="" name="UserForm1">
                <TD width="100%" height="159"> <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                              <TBODY>
                                <TR> 
                                  <TD width="95%" height="68"><div align="left"><font color="#000000"> 
                                      </font> 
                                      <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#E7E7E7">
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">��Ա��ţ�</div></td>
                                          <td><font color="#FF0000"> 
                                            <% = RsUserObj("UserNo") %>
                                            &nbsp;</font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td width="22%"> <div align="right">�û�����</div></td>
                                          <td width="78%"> <% = RsUserObj("MemName") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">������</div></td>
                                          <td> <font color="#666666"> 
                                            <input name="name" type="text" id="name" value="<% = RsUserObj("name") %>">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font><font color="#666666">&nbsp; </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">�Ա�</div></td>
                                          <td><input type="radio" name="Sex" value="0" <%if RsUserObj("Sex")=0 Then Response.Write("checked")%>>
                                            �� 
                                            <input type="radio" name="Sex" value="1" <%if RsUserObj("Sex")=1 Then Response.Write("checked")%>>
                                            Ů</td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">�绰��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="tel" type="text" id="tel" value="<% = RsUserObj("Telephone") %>">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">MSN��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="msn" type="text" id="msn" value="<% = RsUserObj("msn") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">QQ��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="qq" type="text" id="qq" value="<% = RsUserObj("Oicq") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">ʡ�ݣ�</div></td>
                                          <td><select name="Province" id="Province">
                                              <option value="�Ĵ�" <%if RsUserObj("Province")="�Ĵ�" Then Response.Write("selected")%>>�Ĵ�</option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="�㶫"  <%if RsUserObj("Province")="�㶫" Then Response.Write("selected")%>> 
                                              �㶫 </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="���"  <%if RsUserObj("Province")="���" Then Response.Write("selected")%>> 
                                              ��� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="�ӱ�"  <%if RsUserObj("Province")="�ӱ�" Then Response.Write("selected")%>> 
                                              �ӱ� </option>
                                              <option value="ɽ��"  <%if RsUserObj("Province")="ɽ��" Then Response.Write("selected")%>> 
                                              ɽ�� </option>
                                              <option value="ɽ��"  <%if RsUserObj("Province")="ɽ��" Then Response.Write("selected")%>> 
                                              ɽ�� </option>
                                              <option value="������"  <%if RsUserObj("Province")="������" Then Response.Write("selected")%>> 
                                              ������ </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="�Ϻ�"  <%if RsUserObj("Province")="�Ϻ�" Then Response.Write("selected")%>> 
                                              �Ϻ� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="�ຣ"  <%if RsUserObj("Province")="�ຣ" Then Response.Write("selected")%>> 
                                              �ຣ </option>
                                              <option value="�½�"  <%if RsUserObj("Province")="�½�" Then Response.Write("selected")%>> 
                                              �½� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="���ɹ�"  <%if RsUserObj("Province")="���ɹ�" Then Response.Write("selected")%>> 
                                              ���ɹ� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="�㽭"  <%if RsUserObj("Province")="�㽭" Then Response.Write("selected")%>> 
                                              �㽭 </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                              <option value="̨��"  <%if RsUserObj("Province")="̨��" Then Response.Write("selected")%>> 
                                              ̨�� </option>
                                              <option value="���"  <%if RsUserObj("Province")="���" Then Response.Write("selected")%>> 
                                              ��� </option>
                                              <option value="����"  <%if RsUserObj("Province")="����" Then Response.Write("selected")%>> 
                                              ���� </option>
                                            </select> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right"><span class="f41">����</span>��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="city" type="text" id="city" value="<% = RsUserObj("City") %>">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">��ַ��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="address" type="text" id="address" value="<% = RsUserObj("address") %>" size="35">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font><font color="#666666">&nbsp; </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">�������룺</div></td>
                                          <td><font color="#666666"> 
                                            <input name="postcode" type="text" id="postcode" value="<% = RsUserObj("postcode") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">��վ��ַ��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="HomePage" type="text" id="HomePage" value="<% = RsUserObj("HomePage") %>" size="35">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">�������ڣ�</div></td>
                                          <td><font color="#666666"> 
                                            <input name="Birthday" type="text" id="Birthday" value="<% = RsUserObj("Birthday") %>"  readonly>
                                            <input type="button" name="Submit4" value="ѡ������" onClick="OpenWindowAndSetValue('Comm/SelectDate.asp',280,110,window,document.UserForm1.Birthday);document.UserForm1.Birthday.focus();">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td>&nbsp;</td>
                                          <td> <input type="submit" name="Submit" value="�ύ�޸�"> 
                                            <input name="id" type="hidden" id="id" value="<% = RsUserObj("ID") %>"> 
                                            <input name="action" type="hidden" id="action" value="Update">
<script language="JavaScript" type="text/JavaScript">
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
</script></td>
                                        </tr>
                                      </table>
                                      <font color="#000000"> </font> </div>
                                    <span class="f41"> </span> </TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
              </form>
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

