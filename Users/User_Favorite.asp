<% Option Explicit %>
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
	If Request("NID") = "" then
		Response.Write("<script>alert(""请选择！"&CopyRight&""");location=""User_Favorite.asp"";</script>")  
		Response.End
	Else
	    Conn.execute("Delete From FS_Favorite where ID in("&Replace(Request("NID"),"'","")&") and UserID ="&Replace(Replace(Session("MemID"),"'",""),Chr(39),""))
		Response.Write("<script>alert(""删除成功！"&CopyRight&""");location=""User_Favorite.asp"";</script>")  
		Response.End
	End If 
End if
If Request.Form("Action") = "DelAll" Then
	    Conn.execute("Delete From FS_Favorite where isTF=0 and UserID ="&Replace(Replace(Session("MemID"),"'",""),Chr(39),""))
		Response.Write("<script>alert(""收藏夹清除成功！"&CopyRight&""");location=""User_Favorite.asp"";</script>")  
		Response.End
End if
Dim RsFobj,RsFSQL
Set RsFobj = Server.CreateObject(G_FS_RS)
RsFSQL = "Select ID,PID,UserID,Addtime From FS_Favorite where isTF=0 and UserID="& Replace(Session("MemID"),"'","")&" Order By AddTime desc,id desc"
RsFobj.Open RsFSQL,Conn,1,3
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
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
                      <TD class=f4>收 藏 夹</TD>
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
              <form method=POST action="User_Favorite.asp" name=BuyForm  onsubmit="return Cim()">
                <TD width="100%" height="159" valign="top"> 
                  <div align="left"> <font color="#006600"><strong> </strong> 
                    </font> 
                    <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top><TABLE width="100%" border=0 cellPadding=5 cellSpacing=1 
                  background="" bgcolor="#CCCCCC" class=bgup>
                              <TBODY>
                                <TR bgcolor="#E8E8E8"> 
                                  <TD width="50" height="25"> <div align="center"><strong><font color="#000000">选择</font> 
                                      </strong></div></TD>
                                  <TD width="390"><div align="center"><strong>新闻名称</strong></div></TD>
                                  <TD width="146"><div align="center"><strong>新闻日期</strong></div></TD>
                                  <TD width="173"><div align="center"><strong>收藏日期</strong></div></TD>
                                  <TD width="134"><div align="center"><strong>发送给好友|删除</strong></div></TD>
                                </TR>
                                <%
								Do while Not RsFobj.Eof 
								Dim RsPobj 
								Set RsPobj = Conn.execute("Select * from FS_News where ID="&Replace(RsFobj("PID"),"'",""))
								If RsPobj.eof then
										Response.Write("<tr bgcolor=""#FFFFFF""><td><div align=center><input name=""NID"" type=""checkbox"" id=""NID"" value="""&RsFobj("ID")&"""></div></td><td  colspan=""4"" bgcolor=#ffffff><font color=red>此新闻已经被管理员删除</font>&nbsp;&nbsp;</td></tr>")
								Else
								Dim RsClassObj
								Set RsClassObj = Conn.execute("Select ClassEName,SaveFilePath from FS_NewsClass Where ClassID='"&RsPobj("ClassID")&"'")
								%>
                                <TR bgcolor="#FFFFFF"> 
                                  <TD height="26"><div align="center"> 
                                      <input name="NID" type="checkbox" id="NID" value="<% = RsFobj("ID")%>">
                                    </div></TD>
                                  <TD> <%
								  If RsConfigObj("UseDatePath")=1 then
								  	  iF RsPobj("HeadNewsTF")=0 Then
								  %> <a href="<%=RsConfigObj("Domain") &  RsClassObj("SaveFilePath") &"/"& RsClassObj("ClassEName") & RsPobj("Path") & "/" & RsPobj("FileName") & "." & RsPobj("FileExtName")%>" target="_blank"> 
                                    <% = RsPobj("Title")%>
                                    </a> <%
									  Else
								  %> <a href="<%=RsPobj("HeadNewsPath")%>" target="_blank"> 
                                    <% = RsPobj("Title")%>
                                    </a> <%
									  End If
								  Else
								  	  iF RsPobj("HeadNewsTF")=0 Then
								  %> <a href="<%=RsConfigObj("Domain") &  RsClassObj("SaveFilePath")&"/" & RsClassObj("ClassEName") & RsPobj("FileName") & "." & RsPobj("FileExtName")%>" target="_blank"> 
                                    <% = RsPobj("Title")%>
                                    </a> <%
									  Else
								  %> <a href="<%=RsPobj("HeadNewsPath")%>" target="_blank"> 
                                    <% = RsPobj("Title")%>
                                    </a> <%
									  End If
								  End If
								  %> </TD>
                                  <TD><div align="center"> 
                                      <% = RsPobj("AddDate")%>
                                    </div></TD>
                                  <TD> <div align="center"> 
                                      <% = RsFobj("AddTime")%>
                                    </div></TD>
                                  <TD><div align="center"><a href="../SendMail.asp?Newsid=<% = RsPobj("Newsid")%>">发送给好友</a> 
                                      | <a href="User_Favorite.asp?Action=Del&NID=<%=RsFobj("Id")%>"  onClick="return Cim1()">删除</a></div></TD>
                                </TR>
                                <%
								End if
									RsFobj.MoveNext
								Loop
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
                          删除新闻
						  <input type="radio" name="Action" value="DelAll">
                          清空收藏夹 
                          <input type="submit" name="Submit" value="执行操作"></td>
                      </tr>
                    </table>
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
Set Conn=nothing
%>
<script language="JavaScript" type="text/JavaScript">
function Cim(){
	if (window.confirm('您确定要操作?')){
	 	return true;
	 } 
	 return false;		
}
function Cim1(){
	if (window.confirm('您确定要删除吗?')){
	 	return true;
	 } 
	 return false;		
}
</script>
