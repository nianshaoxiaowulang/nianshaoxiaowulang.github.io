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
	Dim RsConfigObj
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,QPoint from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="../Comm/User_Purview.Asp" -->
<%
If request("Action")="Del" Then
	If trim(Request("GID"))<>"" Then
		Dim BookStr,ParaArray,i,NumStr,NumParaArray
		BookStr = Request("GID")
		  if Right(BookStr,1)="," then
			BookStr = Left(BookStr,Len(BookStr)-1)
		  end if
		  if Left(BookStr,1)="," then
			BookStr = Right(BookStr,Len(BookStr)-1)
		  end if
		  ParaArray = Split(BookStr,",")
		For i = LBound(ParaArray) to UBound(ParaArray)
			Dim GBListObj
			Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where id="&Clng(ParaArray(i)))
			If Clng(GBListObj("UserID"))<>Clng(Session("MemID")) Then
				Response.Write("<script>alert(""��ûȨ��ɾ������"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
				Response.End
			End if
			Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where id="&Clng(ParaArray(i)))
		    Conn.execute("Update FS_Members Set Point = Point-"&RsConfigObj("QPoint")&"  where ID="&GBListObj("UserID"))
			Conn.execute("Delete From FS_GBook Where id="&Clng(ParaArray(i)))
		Next
		Response.Write("<script>alert(""ɾ���ɹ���"&CopyRight&""");location=""Gbook.asp"";</script>")  
		Response.End
	Else
		Response.Write("<script>alert(""��ѡ��ɾ�������ӣ�"&CopyRight&""");location=""Gbook.asp"";</script>")  
		Response.End
	End if
End If
%>
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
                      <form name="form1" method="post" action="GBook.asp">
                      <td width="38%"><input name="Keyword" type="text" id="Keyword">
                        <input type="submit" name="Submit2" value="����"> </td>
                    </form>
                    </tr>
                  </table>
                  
                <strong>�ҷ��������</strong><br>
                <br>
                <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <form method=POST action="GBook.asp" name=Form1  onsubmit="return Cim()">
                    <tr bgcolor="#E8E8E8"> 
                      <td width="3%"><strong>ID</strong></td>
                      <td width="6%"> <div align="center"><strong>����</strong></div></td>
                      <td width="44%"><strong>����</strong></td>
                      <td width="22%"><strong>����ʱ��</strong></td>
                      <td width="18%"><strong>�ظ�ʱ��</strong></td>
                      <td width="7%"><strong>�޸�</strong></td>
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
	RsCon.Source="select * from FS_GBook where UserID="& Replace(session("MemID"),"'","") & QS & Ks &" and QID=0 order by Orders,Qtime desc,Addtime desc"
	RsCon.Open RsCon.Source,Conn,1,1
	If RsCon.eof then
		   RsCon.close
		   set RsCon=nothing
		   Response.Write"<TR><TD colspan=""6"" bgcolor=FFFFFF>û�м�¼��</TD></TR>"
	Else
			RsCon.pagesize=15
			RsCon.absolutepage=cint(strpage)
			select_count=RsCon.recordcount
			select_pagecount=RsCon.pagecount
			for i=1 to RsCon.pagesize
			if RsCon.eof then
				exit for
			end if
	if i mod 2 = 0 then
	%>
                    <tr bgcolor="#EEEEEE"> 
                      <%Else%>
                    <tr bgcolor="#FFFFFF"> 
                      <%End If%>
                      <td> <input name="GID" type="checkbox" id="GID" value="<% = RsCon("ID")%>"></td>
                      <td> <div align="center"><img src="images/face<% = RsCon("FaceNum")%>.gif"></div></td>
                      <td>
                        <%If RsCon("Orders")=1 then%>
                        <img src="Images/ztop.gif" alt="�̶���" width="18" height="15"> 
                        <%Else%>
                        <img src="Images/hotfolder.gif" alt="һ������" width="18" height="12"> 
                        <%End if%>
                        <a href="ReadBook.asp?id=<% = RsCon("ID")%>"> 
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
                      <td><a href="Modify_GBook.asp?Id=<% = RsCon("ID")%>">�޸�</a></td>
                    </tr>
                    <%
		RsCon.MoveNext
	Next
	%>
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="6"><input name="Action" type="radio" value="Del" checked>
                        ɾ�� �� 
                        <input type="submit" name="Submit" value="ִ�в���"></td>
                    </tr>
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