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
		If request.Form("SelfIntro")<>"" then
			If len(request.Form("SelfIntro"))>300 then
				Response.Write("<script>alert(""���˽��ܲ��ܳ���300���ַ�����"&CopyRight&""");location=""javascript:history.back()"";</script>")  
				Response.End
			End if
		End if
		If request.Form("UnderWrite")<>"" then
			If len(request.Form("UnderWrite"))>300 then
				Response.Write("<script>alert(""ǩ�����ܳ���300���ַ�����"&CopyRight&""");location=""javascript:history.back()"";</script>")  
				Response.End
			End if
		End if
		Dim RsUpdateObj,SqlUpdate
		Set RsUpdateObj = server.CreateObject (G_FS_RS)
		SqlUpdate = "select * from FS_members where id="& Clng(Replace(Replace(Request.Form("id"),"'",""),Chr(39),""))
		RsUpdateObj.Open SqlUpdate,Conn,1,3
		If Replace(Request.Form("OpenInfTF"),"'","")="0" Then
			RsUpdateObj("OpenInfTF")=0
		Else
			RsUpdateObj("OpenInfTF")=1
		End if
		If Replace(Request.Form("SubInfTF"),"'","")="0" Then
			RsUpdateObj("SubInfTF")=0
		Else
			RsUpdateObj("SubInfTF")=1
		End if
		RsUpdateObj("SelfIntro")=NoCSSHackInput(Request.Form("SelfIntro"))
		RsUpdateObj("HeadPic")=NoCSSHackInput(Replace(Request.Form("HeadPic"),"'",""))
		RsUpdateObj("EduLevel")=NoCSSHackInput(Request.Form("EduLevel"))
		RsUpdateObj("Vocation")=NoCSSHackInput(Replace(Request.Form("Vocation"),"'",""))
		RsUpdateObj("UnderWrite")=NoCSSHackInput(Request.Form("UnderWrite"))
		RsUpdateObj.Update
		RsUpdateObj.Close
		Set RsUpdateObj=Nothing
		Response.Write("<script>alert(""������ϵ���ϳɹ���"&CopyRight&""");location=""User_Modify_other.asp"";</script>")  
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
class=f4>�޸�������ʽ</TD>
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
                                          <td> <div align="right">���˽��ܣ�</div></td>
                                          <td> <font color="#666666"> 
                                            <textarea name="SelfIntro" cols="50" rows="6" id="SelfIntro"><% = RsUserObj("SelfIntro") %></textarea>
                                            </font><font color="#FF0000">&nbsp; 
                                            </font><font color="#666666">&nbsp; 
                                            ֧��HTML�﷨�����300���ַ�</font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">ͷ���ַ��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="HeadPic" type="text" id="HeadPic" value="<% = RsUserObj("HeadPic") %>" size="30">
                                            </font>&nbsp; 
                                            <a href="#" onclick="openScript('SelectHeadPic.asp?action=a',650,400)" title="ͷ��Ԥ���б�"><font color="#FF0000">[ͷ��ѡ��]</font></a> 
                                            </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">ְҵ��</div></td>
                                          <td><font color="#666666"> 
                                            <input name="Vocation" type="text" id="Vocation" value="<% = RsUserObj("Vocation") %>" size="30">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">ѧ����</div></td>
                                          <td><font color="#666666"> 
                                            <input name="EduLevel" type="text" id="EduLevel" value="<% = RsUserObj("EduLevel") %>" size="30">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">�Ƿ񿪷����ϣ�</div></td>
                                          <td><input type="radio" name="OpenInfTF" value="1" <%if RsUserObj("OpenInfTF")=1 then response.Write("checked")%>>
                                            ��
<input type="radio" name="OpenInfTF" value="0" <%if RsUserObj("OpenInfTF")=0 then response.Write("checked")%>>
                                            �� </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">�Ƿ��ı�վ��Ϣ��</div></td>
                                          <td bgcolor="#FFFFFF"><input type="radio" name="SubInfTF" value="1" <%if RsUserObj("SubInfTF")=1 then response.Write("checked")%>>
                                            �� 
                                            <input type="radio" name="SubInfTF" value="0" <%if RsUserObj("SubInfTF")=0 then response.Write("checked")%>>
                                            �� </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">����ǩ����</div></td>
                                          <td><font color="#666666"> 
                                            <textarea name="UnderWrite" cols="50" rows="6" id="UnderWrite"><% = RsUserObj("UnderWrite") %></textarea>
                                            &nbsp; &nbsp; ֧��HTML�﷨</font><font color="#666666"> 
                                            �����300���ַ� </font></td>
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
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no' );
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

