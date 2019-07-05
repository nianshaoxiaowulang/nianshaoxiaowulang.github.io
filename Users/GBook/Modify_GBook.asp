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
	Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent from FS_Config")
	Set DBC = Nothing
%>
<!--#include file="../Comm/User_Purview.Asp" -->
<%
If Request.Form("action")="add" then
		If Len(request.form("Content"))>RsConfigObj("MaxContent") then
			Response.Write("<script>alert(""内容不能超过"&RsConfigObj("MaxContent")&"字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If  Cint(Session("MemID"))=0 then
			Response.Write("<script>alert(""错误的权限！！！"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set Rs = server.createobject(G_FS_RS)
	  Sql1 = "select * from FS_GBook where id="&Replace(Replace(Request.Form("Id"),"'",""),Chr(39),"")
	  Rs.open sql1,conn,1,3
	  Rs("Title")=NoCSSHackInput(Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/",""))
	  Rs("Content")=NoCSSHackContent(Request.Form("Content"))
	  Rs("FaceNum")=NoCSSHackInput(Replace(request.form("FaceNum"),"'",""))
	  Rs("EditQ")= "<br><br><div align=right><font color=#003399>[此信息作者于<"&Now&"> 编辑过]</font></div> "
	  If Request.Form("isAdmin")<>"" then
		  Rs("isAdmin")=1 
	  Else
		  Rs("isAdmin")=0 
	  End if
	  Rs.update
	  		If Request("GetAction")<>"" then
				Response.Write("<script>alert(""修改成功"&CopyRight&""");location=""ReadBook.asp?id="&Request.Form("Sid")&""";</script>")  
			Else
				Response.Write("<script>alert(""修改成功"&CopyRight&""");location=""GBook.asp"";</script>")  
			End if
			Response.End
	  Rs.close
	  Set rs=nothing
End if
Dim RsModifyObj,ModifySQL
	  Set RsModifyObj = server.createobject(G_FS_RS)
	  ModifySQL = "select * from FS_GBook where ID="&Replace(Replace(Request("Id"),"'",""),Chr(39),"")
	  RsModifyObj.open ModifySQL,conn,1,1
	  If Cint(RsModifyObj("UserID"))<>Cint(Session("MemID")) Then
			Response.Write("<script>alert(""您没有编辑此帖子的权限"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
	  End if
Dim NewsContent
NewsContent = Replace(Replace(RsModifyObj("Content"),"""","%22"),"'","%27")
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
                        <input name="Title" type="text" id="Title" value="<% = RsModifyObj("Title")%>" size="30">
                        <input name="isAdmin" type="checkbox" id="isAdmin" value="1" <% If  RsModifyObj("isAdmin")=1 then Response.Write("Checked")%>>
                        管理员可见
                        <input name="GetAction" type="hidden" id="GetAction" value="<% = Request("GetAction")%>">
                        <input name="Sid" type="hidden" id="Sid" value="<% = Request("Sid")%>"></td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td bgcolor="#F2F2F2"> 
                        <div align="right">当前表情：</div></td>
                      <td> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td> <input name="FaceNum" type="radio" value="1" <%If RsModifyObj("FaceNum")=1 then response.Write("Checked")%>> 
                              <img src="Images/face1.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="2" <%If RsModifyObj("FaceNum")=2 then response.Write("Checked")%>> 
                              <img src="Images/face2.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="3" <%If RsModifyObj("FaceNum")=3 then response.Write("Checked")%>> 
                              <img src="Images/face3.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="4" <%If RsModifyObj("FaceNum")=4 then response.Write("Checked")%>> 
                              <img src="Images/face4.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="5" <%If RsModifyObj("FaceNum")=5 then response.Write("Checked")%>> 
                              <img src="Images/face5.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="6" <%If RsModifyObj("FaceNum")=6 then response.Write("Checked")%>> 
                              <img src="Images/face6.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="7" <%If RsModifyObj("FaceNum")=7 then response.Write("Checked")%>> 
                              <img src="Images/face7.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="8" <%If RsModifyObj("FaceNum")=8 then response.Write("Checked")%>> 
                              <img src="Images/face8.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="9" <%If RsModifyObj("FaceNum")=9 then response.Write("Checked")%>> 
                              <img src="Images/face9.gif" width="22" height="22"></td>
                          </tr>
                          <tr> 
                            <td> <input type="radio" name="FaceNum" value="10" <%If RsModifyObj("FaceNum")=10 then response.Write("Checked")%>> 
                              <img src="Images/face10.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="11" <%If RsModifyObj("FaceNum")=11 then response.Write("Checked")%>> 
                              <img src="Images/face11.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="12" <%If RsModifyObj("FaceNum")=12 then response.Write("Checked")%>> 
                              <img src="Images/face12.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="13" <%If RsModifyObj("FaceNum")=13 then response.Write("Checked")%>> 
                              <img src="Images/face13.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="14" <%If RsModifyObj("FaceNum")=14 then response.Write("Checked")%>> 
                              <img src="Images/face14.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="15" <%If RsModifyObj("FaceNum")=15 then response.Write("Checked")%>> 
                              <img src="Images/face15.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="16" <%If RsModifyObj("FaceNum")=16 then response.Write("Checked")%>> 
                              <img src="Images/face16.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="17" <%If RsModifyObj("FaceNum")=17 then response.Write("Checked")%>> 
                              <img src="Images/face17.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="18" <%If RsModifyObj("FaceNum")=18 then response.Write("Checked")%>> 
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
                        <input name="submitggg" type="button" onClick="SubmitFun();" value="修改帖子"> 
                        <input name="reset" type="reset" value="重新填写"> <input name="Content" type="hidden" id="Content" value="<% = NewsContent %>"> 
                        <input name="action" type="hidden" id="action" value="add">
                        <input name="ID" type="hidden" id="ID" value="<% = RsModifyObj("ID")%>"> 
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
Set RsModifyObj = Nothing
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
