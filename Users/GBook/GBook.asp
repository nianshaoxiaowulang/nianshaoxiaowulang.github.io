<% Option Explicit %>
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
				Response.Write("<script>alert(""您没权限删除此帖"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
				Response.End
			End if
			Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where id="&Clng(ParaArray(i)))
		    Conn.execute("Update FS_Members Set Point = Point-"&RsConfigObj("QPoint")&"  where ID="&GBListObj("UserID"))
			Conn.execute("Delete From FS_GBook Where id="&Clng(ParaArray(i)))
		Next
		Response.Write("<script>alert(""删除成功！"&CopyRight&""");location=""Gbook.asp"";</script>")  
		Response.End
	Else
		Response.Write("<script>alert(""请选择删除的帖子！"&CopyRight&""");location=""Gbook.asp"";</script>")  
		Response.End
	End if
End If
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
class=f4>帖子管理</TD>
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
                      
                    <td width="62%"><a href="GBook.asp">我发表的帖子</a> ｜ <a href="All_GBook.asp">帖子查看</a> 
                      ｜ <a href="Write_GBook.asp"><font color="#FF0000">发表帖子</font></a> 
                      ｜ <a href="GBook.asp?Action=Q">已回复的帖子</a> ｜ <a href="GBook.asp?Action=Q"></a><a href="GBook.asp?Action=UnQ">未回复的帖子</a></td>
                      <form name="form1" method="post" action="GBook.asp">
                      <td width="38%"><input name="Keyword" type="text" id="Keyword">
                        <input type="submit" name="Submit2" value="搜索"> </td>
                    </form>
                    </tr>
                  </table>
                  
                <strong>我发表的帖子</strong><br>
                <br>
                <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <form method=POST action="GBook.asp" name=Form1  onsubmit="return Cim()">
                    <tr bgcolor="#E8E8E8"> 
                      <td width="3%"><strong>ID</strong></td>
                      <td width="6%"> <div align="center"><strong>表情</strong></div></td>
                      <td width="44%"><strong>标题</strong></td>
                      <td width="22%"><strong>发表时间</strong></td>
                      <td width="18%"><strong>回复时间</strong></td>
                      <td width="7%"><strong>修改</strong></td>
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
		   Response.Write"<TR><TD colspan=""6"" bgcolor=FFFFFF>没有记录。</TD></TR>"
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
                        <img src="Images/ztop.gif" alt="固顶帖" width="18" height="15"> 
                        <%Else%>
                        <img src="Images/hotfolder.gif" alt="一般帖子" width="18" height="12"> 
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
                      <td><a href="Modify_GBook.asp?Id=<% = RsCon("ID")%>">修改</a></td>
                    </tr>
                    <%
		RsCon.MoveNext
	Next
	%>
                    <tr bgcolor="#FFFFFF"> 
                      <td colspan="6"><input name="Action" type="radio" value="Del" checked>
                        删除 　 
                        <input type="submit" name="Submit" value="执行操作"></td>
                    </tr>
                  </form>
                </table> 
                  
<%
	Response.write"<br>&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
	if int(strpage)>1 then
		Response.Write"&nbsp;<a href=?page=1&Action="&Request("Action")&">第一页</a>&nbsp;"
		Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Action="&Request("Action")&">上一页</a>&nbsp;"
	end if
	if int(strpage)<select_pagecount then
		Response.Write"&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Action="&Request("Action")&">下一页</a>"
		Response.Write"&nbsp;<a href=?page="& select_pagecount &"&Action="&Request("Action")&">最后一页</a>&nbsp;"
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
	if (window.confirm('您确定要操作?')){
	 	return true;
	 } 
	 return false;		
}
</script>