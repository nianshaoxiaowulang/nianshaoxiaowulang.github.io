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
Dim DBC,Conn,URLS
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim Rsconfig,Sqlconfig
%>
<!--#include file="../../Inc/Session.asp" -->
<HTML>
<HEAD>
<TITLE>FSCMS</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
.STYLE1 {
	color: #3366FF;
	font-weight: bold;
}
-->
</style>
</HEAD>
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<style>
a{text-decoration: none;} /* �������»���,��Ϊunderline */ 
a:link {color: #000000;} /* δ���ʵ����� */
a:visited {color: #000000;} /* �ѷ��ʵ����� */
a:hover{color: #FF0000;} /* ����������� */ 
a:active {color: #FF0000;} /* ����������� */
/*BodyCSS����*/
BODY {
scrollbar-face-color: #f6f6f6;
scrollbar-highlight-color: #ffffff; scrollbar-shadow-color: #cccccc; scrollbar-3dlight-color: #cccccc; scrollbar-arrow-color: #000000; scrollbar-track-color: #EFEFEF; scrollbar-darkshadow-color: #ffffff;
}
td	{font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px; color: #000000; text-decoration:none ; text-decoration:none ; }
</style><BODY bgcolor="#FFFFFF" LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 scroll=yes>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E0E0E0">
        <tr> 
          <td bgcolor="#FFFFFF"><table width="98%" height="30" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/main_bg.gif">
              <tr>
                <td>��<font color="#FF0000"><strong>��ӭʹ�����У԰������ϵͳ(<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a>)</strong>V
                  <%
				  Dim RscObj
				  set RscObj = Conn.execute("select Version from FS_Config")
				  Dim RsVersion
				  RsVersion = RscObj("Version")
				  Response.write RsVersion
				  Rscobj.close
				  Set Rscobj=nothing
				  %>
                </font><strong>������<font color="#FF0000">��Ȩ�ţ�20051026</font></strong></td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="1"></td>
              </tr>
            </table>
            <table width="98%" height="45" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="43"> <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">
                    <tr> 
                      <td width="47%">��ǰ�汾:<font color="#FF0000"> 
                        <% = RsVersion %>
                        </font>����</td>
                      <td width="24%" align="right">�ٷ��汾:</td>
                      <td width="29%"><iframe id='NewsContent' src='http://www.foosun.cn/ver/ver.asp' frameborder=0 scrolling=no width='100' height='14'></iframe></td>
                    </tr>
                  </table></td>
                <td width="56%" height="43" bgcolor="#FAFAFA"><iframe id='NewsId' src='http://www.foosun.cn/ver/Foosun_News.asp' frameborder=0 scrolling=no width='100%' height='26'></iframe> 
                  <!--�������»�Ա�������ݽ���-->
                </td>
              </tr>
            </table> 
            <div align="center">
              <table width="98%" height="30" border="0" cellpadding="0" cellspacing="0" background="Images/main_bg.gif">
                <tr> 
                  <td>��<font color="#006699"><strong>��������Ϣ</strong></font></td>
                </tr>
              </table>
              <%
				Dim theInstalledObjects(23)
				theInstalledObjects(0) = "MSWC.AdRotator"
				theInstalledObjects(1) = "MSWC.BrowserType"
				theInstalledObjects(2) = "MSWC.NextLink"
				theInstalledObjects(3) = "MSWC.Tools"
				theInstalledObjects(4) = "MSWC.Status"
				theInstalledObjects(5) = "MSWC.Counters"
				theInstalledObjects(6) = "IISSample.ContentRotator"
				theInstalledObjects(7) = "IISSample.PageCounter"
				theInstalledObjects(8) = "MSWC.PermissionChecker"
				theInstalledObjects(9) = G_FS_FSO
				theInstalledObjects(10) = G_FS_CONN
					
				theInstalledObjects(11) = "SoftArtisans.FileUp"
				theInstalledObjects(12) = "SoftArtisans.FileManager"
				theInstalledObjects(13) = "JMail.SMTPMail"
				theInstalledObjects(14) = "CDONTS.NewMail"
				theInstalledObjects(15) = "Persits.MailSender"
				theInstalledObjects(16) = "LyfUpload.UploadFile"
				theInstalledObjects(17) = "Persits.Upload.1"
				theInstalledObjects(18) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
				theInstalledObjects(19)	= "Persits.Jpeg"				'AspJpeg
				theInstalledObjects(20) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
				theInstalledObjects(21) = "sjCatSoft.Thumbnail"
				theInstalledObjects(22) = "Microsoft.XMLHTTP"
				theInstalledObjects(23) = "Adodb.Stream"
				%>
            </div>
            <table width="98%" height="221" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
              <tr bgcolor="#FAFAFA"> 
                <td height="32">�����������ͣ�<font face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</font></td>
                <td height="32">��վ������·��<font face="Verdana, Arial, Helvetica, sans-serif">��<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="32">�����ط���������������<font face="Verdana, Arial, Helvetica, sans-serif">IP</font>��ַ<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></font></font></td>
                <td width="56%" height="32">������������ϵͳ<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("OS")%></font></font></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="32">���ű���������<span class="small2">��</span><font face="Verdana, Arial, Helvetica, sans-serif"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>��</font></td>
                <td width="56%" height="37">��<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">WEB</font></span>�����������ƺͰ汾<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></font></font></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="32">���ű���ʱʱ��<span class="small2">��</span><font color=#0076AE><%=Server.ScriptTimeout%></font> 
                  ��</td>
                <td width="56%" height="32">��<font face="Verdana, Arial, Helvetica, sans-serif">CDONTS</font>���֧��<span class="small2">��</span> 
				<%
				On Error Resume Next
				Server.CreateObject("CDONTS.NewMail")
				if err=0 then 
					response.write("<font color=#0076AE>��</font>")
				else
					response.write("<font color=red>��</font>")
				end if	 
				err=0
				%>
                </td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="32">������·��<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SCRIPT_NAME")%></font></td>
                <td width="56%" height="32">��<font face="Verdana, Arial, Helvetica, sans-serif"><span class="small2">Jmail</span></font>�������֧��<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">��</font></span> 
                  <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
                  <font color="red">��</font> 
                  <%else%>
                  <font color="0076AE"> ��</font> 
                  <%end if%>
                </td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="32">�����ط�������������Ķ˿�<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SERVER_PORT")%></font></td>
                <td height="32">��Э������ƺͰ汾<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SERVER_PROTOCOL")%></font></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="32">�������� <font face="Verdana, Arial, Helvetica, sans-serif">CPU</font> 
                  ����<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></font> 
                  ����</td>
                <td height="32">���ͻ��˲���ϵͳ�� 
                  <%
					dim thesoft,vOS
					thesoft=Request.ServerVariables("HTTP_USER_AGENT")
					if instr(thesoft,"Windows NT 5.0") then
						vOS="Windows 2000"
					elseif instr(thesoft,"Windows NT 5.2") then
						vOs="Windows 2003"
					elseif instr(thesoft,"Windows NT 5.1") then
						vOs="Windows XP"
					elseif instr(thesoft,"Windows NT") then
						vOs="Windows NT"
					elseif instr(thesoft,"Windows 9") then
						vOs="Windows 9x"
					elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
						vOs="��Unix"
					elseif instr(thesoft,"Mac") then
						vOs="Mac"
					else
						vOs="Other"
					end if
					response.Write(vOs)
					%>
                </td>
              </tr>
            </table>
            <table width="98%" height="30" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/main_bg.gif">
              <tr> 
                <td>��<font color="#006699"><strong>ʹ�ñ��������ȷ�ϵķ�����������������������Ҫ��</strong></font></td>
              </tr>
            </table>
            <table width="98%" height="105" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
              <tr bgcolor="#FAFAFA"> 
                <td width="44%" height="25">��<font face="Verdana, Arial, Helvetica, sans-serif">JRO.JetEngine(ACCESS</font><font face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                  </font> ���ݿ�<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">)</font>��</span> 
                  <%
					On Error Resume Next
					Server.CreateObject("JRO.JetEngine")
					if err=0 then 
					  response.write("<font color=#0076AE>��</font>")
					else
					  response.write("<font color=red>��</font>")
					end if	 
					err=0
					%>
                </td>
                <td width="56%" height="25">�����ݿ�ʹ��<span class="small2">��</span> 
                  <%
					On Error Resume Next
					Server.CreateObject(G_FS_CONN)
					if err=0 then 
					  response.write("<font color=#0076AE>��,����ʹ�ñ�ϵͳ</font>")
					else
					  response.write("<font color=red>��,����ʹ�ñ�ϵͳ</font>")
					end if	 
					err=0
					%>
                </td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="25">��<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">FSO</font></span>�ı��ļ���д<span class="small2">��</span> 
                  <%
					On Error Resume Next
					Server.CreateObject(G_FS_FSO)
					if err=0 then 
					  response.write("<font color=#0076AE>��,����ʹ�ñ�ϵͳ</font>")
					else
					  response.write("<font color=red>��������ʹ�ô�ϵͳ</font>")
					end if	 
					err=0
				   %>
                </td>
                <td height="25">��Microsoft.XMLHTTP 
                  <%If Not IsObjInstalled(theInstalledObjects(22)) Then%>
                  <font color="red">��</font> 
                  <%else%>
                  <font color="0076AE"> ��</font> 
                  <%end if%>
                  (�Ǳ���) ��Adodb.Stream 
                  <%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
                  <font color="red">��</font> 
                  <%else%>
                  <font color="0076AE"> ��</font> 
                  <%end if%>
                </td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="25" colspan="2">���ͻ���������汾�� 
                  <%
		    Dim Agent,Browser,version,tmpstr
		    Agent=Request.ServerVariables("HTTP_USER_AGENT")
		    Agent=Split(Agent,";")
		    If InStr(Agent(1),"MSIE")>0 Then
				Browser="MS Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
			response.Write(""&Browser&"  "&version&"")
		  %>
                  [��ҪIE5.5������,�������������Windows 2000��Windows 2003 Server]</td>
              </tr>
            </table>
            <table width="98%" height="30" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/main_bg.gif">
              <tr> 
                <td>��<font color="#006699"><strong>��ϵ����</strong></font></td>
              </tr>
            </table>
            <table width="98%" height="158" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CDCDCD">
              <tr bgcolor="#FAFAFA"> 
                <td height="20"> <div align="center"> ��Ʒ����</div></td>
                <td height="20">���Ĵ���Ѷ�Ƽ���չ���޹�˾</td>
                <td> <div align="center">�����޸�</div></td>
                <td>��<font color="#0076AE">���IM</font><a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a><a href="http://www.skyim.com/" target="_blank" class="STYLE1"></a></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="13%" height="20"> <div align="center">�ܻ��绰</div></td>
                <td width="31%" height="20">��028-85098980 66026180 </td>
                <td width="17%"> <div align="center">��Ʒ��ѯ</div></td>
                <td width="39%">��028-85098980-601\602\603<br>
                  ��<font color="#0076AE">QQ��159410��394226379 ��66252421 </font></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="13%" height="20"> <div align="center">����֧��</div></td>
                <td width="31%" height="20">��028-85098980-607��606</td>
                <td width="17%"> <div align="center">�ͷ��绰</div></td>
                <td width="39%">��028-85098980-608</td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="13%" height="20"> <div align="center">��������</div></td>
                <td width="31%" height="20">��028-66026180-603</td>
                <td width="17%"> <div align="center">�����ʼ�</div></td>
                <td width="39%">��Service@Foosun.cn</td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td width="13%" height="20"> <div align="center">�ٷ���վ</div></td>
                <td width="31%" height="20">��<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a></td>
                <td width="17%"> <div align="center">��ʾվ��</div></td>
                <td width="39%">��<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="20"> <div align="center">��������</div></td>
                <td height="20">��<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a>                </td>
                <td> <div align="center">��������</div></td>
                <td>��<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a></td>
              </tr>
              <tr bgcolor="#FAFAFA"> 
                <td height="20" colspan="4" bgcolor="#F0F0F0"><div align="center">&copy;2004-2005 
                  CopyRight<a href="http://www.skyim.com/" target="_blank"><font color="#FF0000"><strong>www.skyim.com</strong></font></a>��All 
                Rights Reserved</div></td>
              </tr>
            </table>

		  </td>
        </tr>
      </table>
</td>
  </tr>
</table>
</BODY>
</HTML>
<%
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>