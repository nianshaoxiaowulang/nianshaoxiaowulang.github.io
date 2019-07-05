<% Option Explicit %>
<!--#include file="../../Inc/Function.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent,QPoint from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="../Comm/User_Purview.Asp" -->
<%
If Request.Form("action")="add" then
		If Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/","")="" or request.form("Content")="" then
			Response.Write("<script>alert(""请填写标题和内容"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Len(request.form("Title"))>30 Or Len(request.form("Title"))<3 then
			Response.Write("<script>alert(""标题不能超过30字符小于3个字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Len(request.form("Content"))>RsConfigObj("MaxContent") then
			Response.Write("<script>alert(""内容不能超过"&RsConfigObj("MaxContent")&"字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Cint(Session("MemID"))=0 then
			Response.Write("<script>alert(""错误的权限！！！"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set Rs = server.createobject(G_FS_RS)
	  Sql1 = "select * from FS_GBook where 1=0"
	  Rs.open sql1,conn,1,3
	  Rs.addnew
	  Rs("Title")=NoCSSHackInput(Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/",""))
	  Rs("Content")=NoCSSHackContent(Request.Form("Content"))
	  Rs("AddTime")=Now()
	  Rs("QTime")=Now()
	  Rs("UserID")=Session("MemID")
	  Rs("FaceNum")=NoCSSHackInput(Replace(request.form("FaceNum"),"'",""))
	  Rs("isQ")=0
	  Rs("isLock")=0
	  Rs("Orders")=2
	  Rs("EditQ")=""
	  Rs("QID")=0
	  If Request.Form("isAdmin")<>"" then
		  Rs("isAdmin")=1 
	  Else
		  Rs("isAdmin")=0 
	  End if
	  '增加积分
	  Conn.execute("Update FS_Members Set Point = Point+"&RsConfigObj("QPoint")&" Where Id="&Replace(Replace(Session("MemId"),"'",""),Chr(39),""))
	  Rs.update
	  Response.Write("<script>if (confirm(""发表成功,是否继续?"")==false) window.location=""GBook.asp""; else window.location=""Write_GBook.asp"";</script>")
	  Response.End
	  Rs.close
	  Set Rs=nothing
End if
Dim NewsContent
NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
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
class=f4>发表帖子</TD>
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
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      
                    <td width="62%"><a href="GBook.asp">我发表的帖子</a> ｜ <a href="All_GBook.asp">帖子查看</a> 
                      ｜ <a href="Write_GBook.asp"><font color="#FF0000">发表帖子</font></a> 
                      ｜ <a href="GBook.asp?Action=Q">已回复的帖子</a> ｜ <a href="GBook.asp?Action=Q"></a><a href="GBook.asp?Action=UnQ">未回复的帖子</a></td>
                  <form name="form2" method="post" action="Write_GBook.asp">
                      <td width="38%"><input name="Keyword" type="text" id="Keyword">
                        <input type="submit" name="Submit2" value="搜索"> </td>
                    </form>
                    </tr>
                  </table>
                  
                <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <form action="" method="POST" name="NewsForm">
                    <tr bgcolor="#F2F2F2"> 
                      <td width="16%"> 
                        <div align="right">帖子标题：</div></td>
                      <td width="84%"> 
                        <input name="Title" type="text" id="Title" size="30">
                        <input name="isAdmin" type="checkbox" id="isAdmin" value="1">
                        管理员可见</td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td> 
                        <div align="right">当前表情：</div></td>
                      <td> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td> <input name="FaceNum" type="radio" value="1" checked> 
                              <img src="Images/face1.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="2"> 
                              <img src="Images/face2.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="3"> 
                              <img src="Images/face3.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="4"> 
                              <img src="Images/face4.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="5"> 
                              <img src="Images/face5.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="6"> 
                              <img src="Images/face6.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="7"> 
                              <img src="Images/face7.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="8"> 
                              <img src="Images/face8.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="9"> 
                              <img src="Images/face9.gif" width="22" height="22"></td>
                          </tr>
                          <tr> 
                            <td> <input type="radio" name="FaceNum" value="10"> 
                              <img src="Images/face10.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="11"> 
                              <img src="Images/face11.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="12"> 
                              <img src="Images/face12.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="13"> 
                              <img src="Images/face13.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="14"> 
                              <img src="Images/face14.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="15"> 
                              <img src="Images/face15.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="16"> 
                              <img src="Images/face16.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="17"> 
                              <img src="Images/face17.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="18"> 
                              <img src="Images/face18.gif" width="22" height="22"> 
                            </td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td colspan="2"> 
                        <div align="right"></div>
                        <iframe id='NewsContent' src='../Editer/BookNewsEditer.asp' frameborder=0 scrolling=no width='100%' height='320'></iframe></td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td> 
                        <div align="right"></div></td>
                      <td> 
                        <input name="submitggg" type="button" onClick="SubmitFun();" value="发表帖子"> 
                        <input name="reset" type="reset" value="重新填写"> <input name="Content" type="hidden" id="Content" value="<% = NewsContent %>"> 
                        <input name="action" type="hidden" id="action" value="add"> 
                      </td>
                    </tr>
                  </form>
                </table>
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
%>
<script>
function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.NewsForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.NewsForm.Content.value=='') document.NewsForm.Content.value=TempContentArray[i];
			else document.NewsForm.Content.value=document.NewsForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.NewsForm.submit();
}
</script>
