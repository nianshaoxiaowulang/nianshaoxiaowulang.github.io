<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
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
Dim DBC,conn
Set DBC = new databaseclass
Set conn = DBC.openconnection()
Set DBC = nothing

Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright from FS_Config")
'-------------------�ж��û����͵����ʼ��Ƿ�ע��
Dim RsMemberObj,RsMemberObj1
Set RsMemberObj = Conn.Execute("Select MemName,Email from FS_members where MemName = '" & replace(request.Form("Username"),"'","") &"'")    
If Not RsMemberObj.Eof then
	Response.Write("<script>alert(""�û����Ѿ����ڣ�������ѡ��"&CopyRight&""");location=""javascript:history.back()"";</script>")
	Response.End
End if
Set RsMemberObj1 = Conn.Execute("Select MemName,Email from FS_members where Email ='" & trim(replace(request.Form("email"),"'","")) &"'")
If Not RsMemberObj1.Eof then
	Response.Write("<script>alert(""�����ʼ��Ѿ����ڣ�������ѡ��"&CopyRight&""");location=""javascript:history.back()"";</script>")
	Response.End
End if
Dim Action
Action = "CheckLogin.ASP?UrlAddress=" & Request("UrlAddress")
Function GetCode()
	Dim TestObj
	On Error Resume Next
	Set TestObj = Server.CreateObject("Adodb.Stream")
	Set TestObj = Nothing
	If Err Then
		Dim TempNum
		Randomize timer
		TempNum = cint(8999*Rnd+1000)
		Session("GetCode") = TempNum
		GetCode = Session("GetCode")		
	Else
		GetCode = "<img src=""Comm/GetCode.asp"" onclick='this.src=this.src;' style='cursor:pointer'>"		
	End If
End Function
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> ��д�ʺ���Ϣ</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
  <script language="javascript" src="Comm/MyScript.js"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=160>
      <TABLE cellSpacing=0 cellPadding=0 width=102 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/favorite.left.help.jpg" 
        width=190></TD></TR>
        <TR>
          <TD>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="15%" height=30>&nbsp;</TD>
                                
                              <TD width="85%" class="f4"><strong>��д��ϵ����</strong></TD>
                              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD height=10><IMG height=1 src="" 
width=1></TD></TR></TBODY></TABLE>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="12%" height=30>&nbsp;</TD>
                      <TD width="88%">
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD> ͬ��ע��Э��</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>��д�ʺ���Ϣ</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><font color="#FF0000">��д��ϵ����</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>ע��ɹ�</TD>
                                    </TR>
                                    <TR> 
                                      <TD><IMG height=5 src="images/SelfService.aspx" 
                              width=1></TD>
                                      <TD></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>��д��ϵ����</TD>
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
              <form method=POST action="sRegister_Success.asp" name=UserForm1 onSubmit="return checkdata1()">
                <TD width="100%" height="74"> <div align="left"> <br>
                    <table width="98%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td width="19%"><div align="right"> 
                            <input name="Username" type="hidden" id="Username" value="<% = NoCSSHackInput(Request.Form("Username"))%>">
                            <input name="sPassword" type="hidden" id="sPassword" value="<% = Request.Form("sPassword")%>">
                            <input name="PassQuestion" type="hidden" id="PassQuestion" value="<% = Request.Form("PassQuestion")%>">
                            <input name="PassAnswer" type="hidden" id="PassAnswer" value="<% = Request.Form("PassAnswer")%>">
                            ������</div></td>
                        <td width="29%"> <font color="#FF0000"> 
                          <input name="sName" type="text" id="sName">
                          * </font></td>
                        <td width="52%">����д������ʵ������</td>
                      </tr>
                      <tr> 
                        <td><div align="right">�Ա�</div></td>
                        <td> <input type="radio" name="Sex" value="0">
                          �� 
                          <input type="radio" name="Sex" value="1">
                          Ů <font color="#FF0000">*</font></td>
                        <td>����ѡ���Ա�</td>
                      </tr>
                      <tr> 
                        <td><div align="right">�������ڣ�</div></td>
                        <td> <select name="yyear" id="yyear">
						<%
						For I=1949 to Year(Now)-6
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          �� 
                          <select name="mmonth" id="mmonth">
						<%
						For I=1 to 12
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          �� 
                          <select name="dday" id="dday">
						<%
						For I=1 to 31
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          ��</td>
                        <td>����д������ʵ���գ���������ȡ�����롣</td>
                      </tr>
                      <tr> 
                        <td><div align="right">֤�����</div></td>
                        <td><font color="#FF0000"> 
                          <select name="VerGetType" id="VerGetType">
                            <option value="���֤" selected>���֤</option>
                            <option value="ѧ��֤">ѧ��֤</option>
                            <option value="����֤">����֤</option>
                            <option value="����">����</option>
                          </select>
                          </font> </td>
                        <td rowspan="2">��Ч֤����Ϊȡ���ʺŵ�����ֶΣ����Ժ�ʵ�ʺŵĺϷ���ݣ����������ʵ��д��<br>
                          �ر����ѣ���Ч֤��һ���趨�����ɸ���</td>
                      </tr>
                      <tr> 
                        <td><div align="right">֤�����룺</div></td>
                        <td><font color="#FF0000"> 
                          <input name="VerGetCode" type="text" id="VerGetCode">
                          </font> </td>
                      </tr>
                      <tr> 
                        <td><div align="right">У���룺</div></td>
                        <td><font color="#FF0000"> 
                          <input name="Ver" type="text" id="Ver" size="10">
                          * 
                          <% = GetCode() %>
                          </font> </td>
                        <td>�뽫ͼ�������������������У��ò��������ڷ�ֹע�����<br>
                          <font color="#FF0000">����㳤ʱ��û�в�����������֤��ˢ����֤�����</font></td>
                      </tr>
                      <tr> 
                        <td><div align="right">�绰��</div></td>
                        <td><font color="#FF0000"> 
                          <input name="tel" type="text" id="tel">
                          * </font></td>
                        <td>������д������룬�м���&quot;,&quot;����</td>
                      </tr>
                      <tr> 
                        <td><div align="right">���棺</div></td>
                        <td><font color="#FF0000"> 
                          <input name="fax" type="text" id="fax">
                          </font></td>
                        <td>������д������룬�м���&quot;,&quot;����</td>
                      </tr>
                      <tr>
                        <td><div align="right">ʡ�ݣ�</div></td>
                        <td colspan="2">
						     <select name="Province" id="Province">
                            <option value="�Ĵ�" selected>�Ĵ�</option>
                            <option value="����" > ���� </option>
                            <option value="�㶫" > �㶫 </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="���" > ��� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="�ӱ�" > �ӱ� </option>
                            <option value="ɽ��" > ɽ�� </option>
                            <option value="ɽ��" > ɽ�� </option>
                            <option value="������" > ������ </option>
                            <option value="����" > ���� </option>
                            <option value="�Ϻ�" > �Ϻ� </option>
                            <option value="����" > ���� </option>
                            <option value="�ຣ" > �ຣ </option>
                            <option value="�½�" > �½� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="���ɹ�" > ���ɹ� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="�㽭" > �㽭 </option>
                            <option value="����" > ���� </option>
                            <option value="����" > ���� </option>
                            <option value="̨��" > ̨�� </option>
                            <option value="���" > ��� </option>
                            <option value="����" > ���� </option>
                          </select>
                          ���У�
                          <input name="City" type="text" id="City" size="12"></td>
                      </tr>
                      <tr> 
                        <td><div align="right">��ַ��</div></td>
                        <td colspan="2"><font color="#FF0000"> 
                          <input name="address" type="text" id="address" size="35">
                          * </font> �������ϸ��д����Ŀ</td>
                      </tr>
                      <tr> 
                        <td><div align="right">�������룺</div></td>
                        <td colspan="2"><input name="PostCode" type="text" id="PostCode"></td>
                      </tr>
                      <tr> 
                        <td><div align="right">�����ʼ���</div></td>
                        <td colspan="2"><font color="#FF0000"> 
                          <input name="email" type="text" id="email" value="<% = Request.Form("email")%>">
                          *</font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2"><input  type=submit name="Submit3" value="��һ��" style="cursor:hand;"></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">ע�⣺<br>
                          1����*����Ŀ������д������ע�᲻�ܼ����� <br>
                          2���Ƽ���ʹ������2G�����������@126.com����� <a href="http://reg.126.com/reg1.jsp" target="_blank"><font color="#FF0000">����ע��</font></a> 
                        </td>
                      </tr>
                    </table>
                  </div></TD></form>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> <hr size="1" noshade color="#FF6600">
      <div align="center">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
<BR>
</BODY></HTML>
<script language="JavaScript" type="text/JavaScript">
function CheckName(gotoURL) {
   var ssn=UserForm.Username.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=UserForm.email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
</script>

