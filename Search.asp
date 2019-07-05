<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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
'==============================================================================
Dim ConfigObj
Set ConfigObj=Conn.execute("Select Copyright,siteName From FS_Config")
'Response.Write(ReSQL)
'Response.end
strpage=request.querystring("Page")
if len(strpage)=0 then
	strpage="1"
end if
Dim Temp_DummyDir,Appraise
If SysRootDir <> "" then
	Temp_DummyDir = "/" & SysRootDir
Else
	Temp_DummyDir = ""
End If
dim DatePathStr
%>
<html>
<head>
<title><% = ConfigObj("SiteName")%>___关键字:<%=request("KeyWord")%>  搜索结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.MainMenuBGStyle {
	background-repeat: no-repeat;
	background-position: right center;
}
-->
</style>
</head>
<link href="CSS/FS_css.css" rel="stylesheet">
<body>
<table width="95%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="tabbgcolor">
  <tr class="tabbgcolorliWhite"> 
    <td width="78%" colspan="2" bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="6">
        <tr> 
          <td><TABLE width="100%" border=0 cellpadding="5" cellspacing="0">
              <TBODY>
                <TR> 
                  <TD width=26><IMG 
                              src="<%=UserDir%>/images/Favorite.OnArrow.gif" border=0></TD>
                  <TD bgcolor="#FFFFFF" 
class=f4><font color="#FF3300">搜索结果</font></TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
              <TBODY>
                <TR> 
                  <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
                </TR>
              </TBODY>
            </TABLE></td>
        </tr>
        <tr> 
          <td>以下为关键字<font color="#FF0000"><strong><%=request("KeyWord")%></strong></font>的相关内容</td>
        </tr>
      </table>
<%
'判断开始
Dim Keyword,Condition,Types,ReSQL
Keyword = Replace(Replace(request("keyword"),"'",""),Chr(39),"")
Condition = Replace(Replace(request("condition"),"'",""),Chr(39),"")
Types = Replace(Replace(request("Types"),"'",""),Chr(39),"")
Dim Rs
Set Rs = server.CreateObject(G_FS_RS)
If Keyword<>"" then
	If  Types = "News" then
		Dim N1
		If Condition = "title" Then
			N1 = " and Title like '%"& Keyword &"%'"
		ElseiF Condition = "content" Then
			N1 = " and Content like '%"& Keyword &"%'"
		Else 
			N1 = " and Author like '%"& Keyword &"%'"
		End if
		ReSQL = "select * from FS_News where DelTF=0 and AuditTF=1 "& N1 &" Order by Id desc"
	 ElseIf Types="DownLoad" then
	 	Dim k1
		If Condition = "title" Then
			k1 = " and Name like '%"& Keyword &"%'"
		ElseiF Condition = "content" Then
			k1 = " and description like '%"& Keyword &"%'"
		Else
			k1 = " and Provider like '%"& Keyword &"%'"
		End if
		ReSQL = "select * from FS_Download where AuditTF=1 "& k1 &" Order by Id desc"
	 ElseIf  Types="Mall" then
	 	Dim M1
		If Condition = "title" Then
			M1 = " and Product_Name like '%"& Keyword &"%'"
		ElseiF Condition = "content" Then
			M1 = " and Products_description like '%"& Keyword &"%'"
		Else
			M1 = " and Products_MakeCompany like '%"& Keyword &"%'"
		End if
		ReSQL = "select * from FS_Shop_Products where IsLock=0 "& M1 &" Order by Id desc"
	 End if
Else
	Response.Write("<script>alert(""请输入关键字!!!!\n\nPowered By FoosunCMS"");location=""javascript:history.back()"";</script>")  
	Response.End
End if
Rs.Open ReSQL,Conn,1,1
if Rs.eof then
	Response.write "<p align='center'> 未找到符合你条件的<font color=#ff0000>"& Request("keyword") &"</font>信息</p>"
Else
	if Types = "News" Then
		Rs.pagesize=40
		Rs.absolutepage=cint(strpage)
		select_count=Rs.recordcount
		select_pagecount=Rs.pagecount
		for i=1 to Rs.pagesize
			if Rs.eof then
			exit for
			end if
			response.write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			dim ClassCName1,RssearchObj
			Set RssearchObj = Conn.Execute("Select SaveFilePath,ClassEName,ClassCName,ClassID,FileExtName from FS_NewsClass where Classid = '" & Replace(Rs("ClassID"),"'","") &"'")
			if Not RssearchObj.eof then
				If Application("UseDatePath")="1" then DatePathStr=Rs("Path") else DatePathStr=""
				response.write "<tr>"
				response.write "<td height=""22"" bgcolor=""#FFFFFF"">·"
				ClassCName1="<a href="& Temp_DummyDir & RssearchObj("SaveFilePath")&"/"& RssearchObj("ClassEName") &"/index."&RssearchObj("FileExtName")&" target=""_blank""><font color=#336699><b>["& RssearchObj("ClassCName") &"]</b></font></a> "
				if Rs("HeadNewsTF")=1 then
					response.write ClassCName1 & "<a href=" & Rs("HeadNewsPath") & " target=""_blank""><font color=""#FF0000""><B>"& Rs("title") &"</B></font></a>[标]  <font color=#999999>" & Rs("AddDate")&"</font>"
				else
					response.write ClassCName1 & "<a href=" & Temp_DummyDir & RssearchObj("SaveFilePath") & "/" & RssearchObj("ClassEName") & DatePathStr & "/" &Rs("FileName") & "." & Rs("FileExtName") & " target=""_blank"">" & Rs("title") & "</a>  <font color=#999999>" & Rs("AddDate") & "</font>" 
				end if
				response.write "</td>"
				response.write "</tr>"
			 end if
			response.write "</table>" 
		  rs.movenext
		 next
	ElseIf Types = "DownLoad" Then
			Appraise = Array("","★","★★","★★★","★★★★","★★★★★","★★★★★★")
		%> 
      <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
        <tr align="center" bgcolor="#EAEAEA"> 
          <td width="16%"><div align="left"><strong>栏目</strong></div></td>
          <td width="25%"><div align="left"><strong>名称</strong></div></td>
          <td width="22%"><div align="left"><strong>版本号</strong></div></td>
          <td width="21%"><div align="left"><strong>系统类型</strong></div></td>
          <td width="16%"><div align="left"><strong>评价</strong></div></td>
        </tr>
        <%
			Rs.pagesize=50
			Rs.absolutepage=cint(strpage)
			Select_count=Rs.recordcount
			Select_pagecount=Rs.pagecount
			for i=1 to Rs.pagesize
				if Rs.eof then
					exit for
				end if
				if i mod 2 = 1 then
				%>
        <tr bgcolor="#ffffff"> 
          <%
				else
		%>
        <tr bgcolor="#eeeeee"> 
          <%
				end if
				
				Set RssearchObj = Conn.Execute("Select ClassCName,ClassEName,SaveFilePath,FileExtName from FS_NewsClass where Classid = '" & Rs("ClassID") &"'")
				ClassCName1="<a href="& Temp_DummyDir & RssearchObj("SaveFilePath")&"/"& RssearchObj("ClassEName") &"/index."&RssearchObj("FileExtName")&" target=""_blank"">["& RssearchObj("ClassCName") &"]</a> "
				%>
          <td height="31">·<%=ClassCName1%></td>
          <td><a href="<%=Temp_DummyDir&RssearchObj("SaveFilePath")&"/"&RssearchObj("ClassEName")&"/"&Rs("filename")&"."&Rs("fileextname")%>" target="_blank"><%=Rs("name")%></a></td>
          <td><%=Rs("Version")%></td>
          <td><%=Rs("SystemType")%></td>
          <td><%=Appraise(Cint(Rs("Appraise")))%></td>
        </tr>
        <%
			rs.movenext
		 	next
		%>
      </table>
      <%
	Elseif Types = "Mall" Then
	%>
      <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
        <tr align="center" bgcolor="#EAEAEA"> 
          <td width="16%"><div align="left"><strong>栏目</strong></div></td>
          <td width="32%"><div align="left"><strong>产品名称</strong></div></td>
          <td width="25%"><div align="left"><strong>产品序号</strong></div></td>
          <td width="16%"><div align="left"><strong>现在价格</strong></div></td>
          <td width="11%"><div align="left"><strong>购买</strong></div></td>
        </tr>
        <%
			Rs.pagesize=20
			Rs.absolutepage=cint(strpage)
			Select_count=Rs.recordcount
			Select_pagecount=Rs.pagecount
			for i=1 to Rs.pagesize
				if Rs.eof then
					exit for
				end if
				if i mod 2 = 1 then
				%>
        <tr bgcolor="#ffffff"> 
          <%
				else
		%>
        <tr bgcolor="#eeeeee"> 
          <%
				end if
				
				Set RssearchObj = Conn.Execute("Select ClassCName,ClassEName,SaveFilePath,FileExtName from FS_NewsClass where Classid = '" & Rs("ClassID") &"'")
				ClassCName1="<a href="& Temp_DummyDir & RssearchObj("SaveFilePath")&"/"& RssearchObj("ClassEName") &"/index."&RssearchObj("FileExtName")&" target=""_blank"">「"& RssearchObj("ClassCName") &"」</a> "
				%>
          <td height="31">·<%=ClassCName1%></td>
          <td><a href="<%=Temp_DummyDir&RssearchObj("SaveFilePath")&"/"&RssearchObj("ClassEName")&Rs("Product_SavPath")&"/"&Rs("Product_FileName")&"."&Rs("Product_ExName")%>" target="_blank"><%=Rs("Product_Name")%></a></td>
          <td><%=Rs("Products_serial")%></td>
          <td><%=Rs("Products_MemberPrice")%>RMB</td>
          <td><a href=<% = UserDir%>/Mall/BuyProduct.asp?Pid=<% =Rs("Id")%>>加入购物车</a></td>
        </tr>
        <%
			rs.movenext
		 	next
		%>
      </table>
	<%
	End if
End if
	%> <br> <br> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> <%
	   response.write"&nbsp;&nbsp;共<b>"& select_pagecount &"</b>页<b>" & select_count &"</b>条记录，本页是第<b>"& strpage &"</b>页。"
		if int(strpage)>1 then
		    response.Write"&nbsp;&nbsp;&nbsp;<a href=?page=1&condition="&Condition&"&keyword="&keyword&"&Types="&Types&">第一页</a>&nbsp;"
		    response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)-1)&"&condition="&Condition&"&keyword="&keyword&"&Types="&Types&">上一页</a>&nbsp;"
			end if
			if int(strpage)<select_pagecount then
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="&cstr(cint(strpage)+1)&"&condition="&Condition&"&keyword="&keyword&"&Types="&Types&">下一页</a>"
			response.Write"&nbsp;&nbsp;&nbsp;<a href=?page="& select_pagecount &"&condition="&Condition&"&keyword="&keyword&"&Types="&Types&">最后一页</a>&nbsp;"
			end if
			response.Write"<br>"
       %></td>
        </tr>
      </table></td>
  </tr>
  <tr class="tabbgcolorliWhite">
    <td colspan="2" bgcolor="#F2F2F2">
<div align="center">
<% = ConfigObj("Copyright")%></div></td>
  </tr>
</table>
</body>
</html>
<%
Set ConfigObj = Nothing
%>