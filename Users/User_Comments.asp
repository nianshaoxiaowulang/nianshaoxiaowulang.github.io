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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
	Dim DBC,conn,sConn
	Set DBC = new databaseclass
	Set Conn = DBC.openconnection()
	Dim I,RsConfigObj
	Set RsConfigObj = Conn.Execute("Select Domain,SiteName,UserConfer,Copyright,isEmail,isChange,UseDatePath from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<%
Dim PlusDir1
If PlusDir <> "" then
	PlusDir1 = PlusDir & "/"
Else
	PlusDir1 = ""
End if
If Request("Action") = "Del" Then
	If Request("RID") = "" then
		Response.Write("<script>alert(""����Ĳ�����"&CopyRight&""");location=""User_Comments.asp"";</script>")  
		Response.End
	Else
		Conn.execute("Delete From FS_Review where ID in("&Replace(Request("RID"),"'","")&") and UserID='"&Session("MemName")&"'")
		Response.Write("<script>alert(""ɾ���ɹ���"&CopyRight&""");location=""User_Comments.asp"";</script>")  
		Response.End
		Set ListObj = Nothing
	End If 
End if
If Request.Form("Action") = "DelAll" Then
	Conn.execute("Delete From FS_Review where Audit=0 and UserID ='"&Replace(Replace(Session("MemName"),"'",""),Chr(39),"")&"'")
	Response.Write("<script>alert(""�ղؼ�����ɹ���"&CopyRight&""");location=""User_Comments.asp"";</script>")  
	Response.End
End if
Dim strpage
strpage=request.querystring("page")
if len(strpage)=0 then
	strpage="1"
end if
Dim RsFobj,RsFSQL
Set RsFobj = Server.CreateObject(G_FS_RS)
Dim Keywords
If Request("Keyword")<> "" then
	Keywords = " and Content Like '%" &Request.Form("Keyword") & "%'"
Else
	Keywords = ""
End if
Dim Tp
If Request("Types")= "1" then
	Tp = " and types=1"
ElseIf Request("Types")= "2" then
	Tp = " and types=2"
ElseIf Request("Types")= "3" then
	Tp = " and types=3"
Else
	Tp = " "
End if
RsFSQL = "Select * From FS_Review where UserID='"& Replace(Session("MemName"),"'","")&"' "& Keywords & Tp &" Order by ID Desc"
RsFobj.Open RsFSQL,Conn,1,3
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> ��Ա����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
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
                      <TD width=26><IMG src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD width="402" class=f4><p>�ҷ��������</p></TD>
                      <TD width="103" class=f4><div align="right">������</div></TD>
                      <form name="form1" method="post" action="User_Comments.asp"><TD width="404" class=f4>
                          <input name="Keyword" type="text" id="Keyword" value="<% = Request("Keyword")%>">
                          <input type="submit" name="Submit2" value="����">
                        </TD></form>
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
              <form method=POST action="User_Comments.asp" name=BuyForm  onsubmit="return Cim()">
                <TD width="100%" height="103" valign="top"> 
                  <div align="left"> <a href="User_Comments.asp">��������</a> | <a href="User_Comments.asp?types=2">��������</a> 
                    | <a href="User_Comments.asp?types=1">��������</a> | <a href="User_Comments.asp?types=3">��Ʒ����</a> 
                    <font color="#006600"> <strong><br>
                    <br>
                    </strong> </font> 
                    <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top><TABLE width="100%" border=0 cellPadding=5 cellSpacing=1 
                  background="" bgcolor="#CCCCCC" class=bgup>
                              <TBODY>
                                <TR bgcolor="#E8E8E8"> 
                                  <TD width="29" height="26"> <div align="center"><font color="#000000">ѡ��</font> 
                                      <span class="f41"> </span> </div></TD>
                                  <TD width="289"><div align="center">����</div></TD>
                                  <TD width="254">�鿴</TD>
                                  <TD width="150"><div align="center">����</div></TD>
                                  <TD width="171"><div align="center">״̬|ɾ��</div></TD>
                                </TR>
                <%
				If Not RsFobj.eof then
					Dim select_count,select_pagecount
					RsFobj.pagesize=20
					RsFobj.absolutepage=cint(strpage)
					select_count=RsFobj.recordcount
					select_pagecount=RsFobj.pagecount
						for i=1 to RsFobj.pagesize
								if RsFobj.eof then
									exit for
								end if
								%>
                                <TR bgcolor="#FFFFFF"> 
                                  <TD height="26"><div align="center"> 
                                      <input name="RID" type="checkbox" id="RID" value="<% = RsFobj("ID")%>">
                                    </div></TD>
                                  <TD><% 
								  If Len(RsFobj("Content")) > 50 then
									  Response.Write Left(RsFobj("Content"),50) &".."
								  Else
									  Response.Write RsFobj("Content")
								  End if
								  %></TD>
                                  <TD>
								  <%
								  If RsFobj("types")=3 then
									  Dim RsProductsObj
									  Set RsProductsObj = Conn.execute("Select ID,Product_Name,Products_AddTime From FS_Shop_Products where ID="& Clng(RsFobj("NewsID")) &"")
									  If RsProductsObj.Eof then
										 Response.Write("<font color=red>�˲�Ʒ�Ѿ���ɾ��</font>")
									  Else
								  %>
								  <a href="../<% = PlusDir &"/"& MallDir%>/Comment.asp?PId=<% =Clng(RsFobj("NewsID"))%>"  Title="��Ʒ���ƣ�<%=RsProductsObj("Product_Name")%>
�ϼ����ڣ�<%=RsProductsObj("Products_AddTime")%>">�鿴����Ʒ������</a> 
								  <%
									  Set RsProductsObj = Nothing
									  End if
								  Else
									  Dim RsNewsObj
									  Set RsNewsObj = Conn.execute("Select NewsID,Title,addDate,Author From FS_News where NewsID='"& RsFobj("NewsID") &"'")
									  If RsNewsObj.Eof then
										 Response.Write("<font color=red>�������Ѿ���ɾ��</font>")
									  Else
								  %>
								  <a href="../NewsReview.asp?NewsId=<% =RsFobj("NewsID")%>" Title="���ű��⣺<%=RsNewsObj("Title")%>
�������ڣ�<%=RsNewsObj("AddDate")%>
�������ߣ�<%=RsNewsObj("Author")%>">�鿴������/���ص�����</a> 
								  <%
									  Set RsNewsObj = Nothing
									  End if
								  End if
								  %>
                                  </TD>
                                  <TD><div align="center"> 
                                      <% = RsFobj("Addtime")%>
                                    </div></TD>
                                  <TD><div align="center"> 
								  <%
								  If RsFobj("Audit") = 0 then
								  	Response.Write("<font color=red>δ���</font>")
								  Else
								  	Response.Write("�����")
								  End if
								  %>
                                      | <%If RsFobj("Audit") = 0 then%><a href=User_CommentsModify.asp?Id=<% = RsFobj("ID")%>>�޸�</A><%Else%><font color="#999999">�޸�</font><%End if%> | <a href="User_Comments.asp?Action=Del&RID=<%=RsFobj("Id")%>"  onClick="return Cim1()">ɾ��</a></div></TD>
                                </TR>
                                <%
									RsFobj.MoveNext
								Next
				 Else
					Response.Write("<tr bgcolor=""#FFFFFF""><td  colspan=""5"" bgcolor=#ffffff><font color=red>û�м�¼</font>&nbsp;&nbsp;</td></tr>")
				 End if
								%>
                              </TBODY>
                            </TABLE> </TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <table width="95%" border="0" align="center" cellpadding="5" cellspacing="0">
                      <tr>
                        <td> 
                          <input name="Action" type="radio" value="Del" checked>
                          ɾ������
<input type="radio" name="Action" value="DelAll">
                          �������
<input type="submit" name="Submit" value="ִ�в���"></td>
                      </tr>
                    </table>
                    <strong></strong></div></TD>
              </form>
            </TR>
          </TBODY>
        </TABLE>
<%
	   response.write"&nbsp;&nbsp;��<b>"& select_pagecount &"</b>ҳ<b>" & select_count &"</b>����¼����ҳ�ǵ�<b>"& strpage &"</b>ҳ��"
		if int(strpage)>1 then
		   Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page=1&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">��һҳ</a>&nbsp;"
		   Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">��һҳ</a>&nbsp;"
		end if
		if int(strpage)<select_pagecount then
			Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">��һҳ</a>"
			Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="& select_pagecount &"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">���һҳ</a>&nbsp;"
		end if
		Response.Write"<br>"
	   %> </TD>
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
%>
<script language="JavaScript" type="text/JavaScript">
function Cim(){
	if (window.confirm('��ȷ��Ҫ����?')){
	 	return true;
	 } 
	 return false;		
}
function Cim1(){
	if (window.confirm('��ȷ��Ҫɾ����?')){
	 	return true;
	 } 
	 return false;		
}
</script>
