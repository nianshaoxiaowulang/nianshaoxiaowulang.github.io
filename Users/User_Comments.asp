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
	If Request("RID") = "" then
		Response.Write("<script>alert(""错误的参数！"&CopyRight&""");location=""User_Comments.asp"";</script>")  
		Response.End
	Else
		Conn.execute("Delete From FS_Review where ID in("&Replace(Request("RID"),"'","")&") and UserID='"&Session("MemName")&"'")
		Response.Write("<script>alert(""删除成功！"&CopyRight&""");location=""User_Comments.asp"";</script>")  
		Response.End
		Set ListObj = Nothing
	End If 
End if
If Request.Form("Action") = "DelAll" Then
	Conn.execute("Delete From FS_Review where Audit=0 and UserID ='"&Replace(Replace(Session("MemName"),"'",""),Chr(39),"")&"'")
	Response.Write("<script>alert(""收藏夹清除成功！"&CopyRight&""");location=""User_Comments.asp"";</script>")  
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
                      <TD width="402" class=f4><p>我发表的评论</p></TD>
                      <TD width="103" class=f4><div align="right">搜索：</div></TD>
                      <form name="form1" method="post" action="User_Comments.asp"><TD width="404" class=f4>
                          <input name="Keyword" type="text" id="Keyword" value="<% = Request("Keyword")%>">
                          <input type="submit" name="Submit2" value="搜索">
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
                  <div align="left"> <a href="User_Comments.asp">所有评论</a> | <a href="User_Comments.asp?types=2">下载评论</a> 
                    | <a href="User_Comments.asp?types=1">新闻评论</a> | <a href="User_Comments.asp?types=3">商品评论</a> 
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
                                  <TD width="29" height="26"> <div align="center"><font color="#000000">选择</font> 
                                      <span class="f41"> </span> </div></TD>
                                  <TD width="289"><div align="center">评论</div></TD>
                                  <TD width="254">查看</TD>
                                  <TD width="150"><div align="center">日期</div></TD>
                                  <TD width="171"><div align="center">状态|删除</div></TD>
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
										 Response.Write("<font color=red>此产品已经被删除</font>")
									  Else
								  %>
								  <a href="../<% = PlusDir &"/"& MallDir%>/Comment.asp?PId=<% =Clng(RsFobj("NewsID"))%>"  Title="产品名称：<%=RsProductsObj("Product_Name")%>
上架日期：<%=RsProductsObj("Products_AddTime")%>">查看此商品的评论</a> 
								  <%
									  Set RsProductsObj = Nothing
									  End if
								  Else
									  Dim RsNewsObj
									  Set RsNewsObj = Conn.execute("Select NewsID,Title,addDate,Author From FS_News where NewsID='"& RsFobj("NewsID") &"'")
									  If RsNewsObj.Eof then
										 Response.Write("<font color=red>此新闻已经被删除</font>")
									  Else
								  %>
								  <a href="../NewsReview.asp?NewsId=<% =RsFobj("NewsID")%>" Title="新闻标题：<%=RsNewsObj("Title")%>
新闻日期：<%=RsNewsObj("AddDate")%>
新闻作者：<%=RsNewsObj("Author")%>">查看此新闻/下载的评论</a> 
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
								  	Response.Write("<font color=red>未审核</font>")
								  Else
								  	Response.Write("已审核")
								  End if
								  %>
                                      | <%If RsFobj("Audit") = 0 then%><a href=User_CommentsModify.asp?Id=<% = RsFobj("ID")%>>修改</A><%Else%><font color="#999999">修改</font><%End if%> | <a href="User_Comments.asp?Action=Del&RID=<%=RsFobj("Id")%>"  onClick="return Cim1()">删除</a></div></TD>
                                </TR>
                                <%
									RsFobj.MoveNext
								Next
				 Else
					Response.Write("<tr bgcolor=""#FFFFFF""><td  colspan=""5"" bgcolor=#ffffff><font color=red>没有记录</font>&nbsp;&nbsp;</td></tr>")
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
                          删除评论
<input type="radio" name="Action" value="DelAll">
                          清空所有
<input type="submit" name="Submit" value="执行操作"></td>
                      </tr>
                    </table>
                    <strong></strong></div></TD>
              </form>
            </TR>
          </TBODY>
        </TABLE>
<%
	   response.write"&nbsp;&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
		if int(strpage)>1 then
		   Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page=1&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">第一页</a>&nbsp;"
		   Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">上一页</a>&nbsp;"
		end if
		if int(strpage)<select_pagecount then
			Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">下一页</a>"
			Response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="& select_pagecount &"&Keyword="&Request("Keyword")&"&Types="& Request("Types")&">最后一页</a>&nbsp;"
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
