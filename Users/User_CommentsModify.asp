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
If Request.Form("Action") = "Update" then
	iF Len(Request.Form("Content"))>300 and Len(Request.Form("Content"))<=1 then
		Response.Write("<script>alert(""错误：\n评论内容应该大于1小于300个字符！"");location.href=""javascript:history.go(-1)"";</script>")
		Response.End
	End if
	Dim RsFobj1,RsFSQL1
	Set RsFobj1 = Server.CreateObject(G_FS_RS)
	RsFSQL1 = "Select ID,Content,Audit From FS_Review where ID="& Replace(Replace(Request.Form("ID"),"'",""),Chr(39),"")
	RsFobj1.Open RsFSQL1,Conn,1,3
	RsFobj1("Content") = Request.Form("Content")
	RsFobj1("Audit") = 0
	RsFobj1.Update
	RsFobj1.Close
	Set RsFobj1 =nothing
	Response.Write("<script>alert(""修改成功！"&Copyright&""");location.href=""User_Comments.asp"";</script>")
	Response.End
End if
Dim RsFobj,RsFSQL
Set RsFobj = Server.CreateObject(G_FS_RS)
RsFSQL = "Select * From FS_Review where ID="& Replace(Replace(Request("ID"),"'",""),Chr(39),"")
RsFobj.Open RsFSQL,Conn,1,1
iF RsFobj("Audit")=1 Then
	Response.Write("<script>alert(""错误：\n审核后的评论不允许修改！"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
End If
iF RsFobj("UserID")<>Session("MemName") Then
	Response.Write("<script>alert(""错误：\n你没权限修改此评论！"");location.href=""javascript:history.go(-1)"";</script>")
	Response.End
End If
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
      <form name="form2" method="post" action=""><TD vAlign=top> 
        
          <TABLE cellSpacing=0 cellPadding=5 width="98%" align=center 
                  border=0>
            <TBODY>
              <TR> 
                <TD width="100%"> <TABLE width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD width=20><IMG src="images/Favorite.OnArrow.gif" border=0></TD>
                        <TD width="923" class=f4><p>修改我发表的评论</p></TD>
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
                <TD width="100%" height="103" valign="top"> <div align="left"> 
                    <strong> 
                    <textarea name="Content" cols="60" rows="6" id="Content"><% = RsFobj("Content")%></textarea>
                    </strong></div></TD>
              </TR>
              <TR> 
                <TD height="31" valign="top">
<input type="submit" name="Submit" value="修改评论">
                  <input name="Action" type="hidden" id="Action" value="Update">
                  <input name="ID" type="hidden" id="ID" value="<% = RsFobj("id")%>"></TD>
              </TR>
            </TBODY>
          </TABLE>
        </TD></form>
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
