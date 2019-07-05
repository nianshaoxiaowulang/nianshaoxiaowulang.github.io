<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040104") then Call ReturnError1()
Dim ClassSql,RsClassObj

Dim PConfigObj,ShowRows1,ShowRows2,HaveValueTF
Set PConfigObj = Conn.execute("Select isShop From FS_Config")
If PConfigObj("isShop")=1 then
	HaveValueTF=1
	ShowRows1=10
	ShowRows2=8
Else
	HaveValueTF=0
	ShowRows1=8
	ShowRows2=7
End If
ClassSql = "Select * from FS_NewsClass where parentID='0' and DelFlag=0"
Set RsClassObj = Conn.Execute(ClassSql)

Function GetChildClassList(ClassID,Str)
	Dim Sql,RsTempObj,TempImageStr,ImageStr,iScheck
	TempImageStr = "<img src=""../../Images/Folder/Node.gif""><img src=""../../Images/Folder/folderclosed.gif"">"
	Sql = "Select * from FS_NewsClass where ParentID='" & ClassID & "' and DelFlag=0"
	ImageStr = Str & "<img src=""../../Images/Folder/HR.gif"">"
	Set RsTempObj = Conn.Execute(Sql)
	do while Not RsTempObj.Eof
		if InStr(1, rs1("PopLIst"),""&RsTempObj("classID")&"" ,1)<>0 then iScheck=" checked"
		GetChildClassList = GetChildClassList & "<tr><td><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr align=""left""  class=""TempletItem""><td>" & ImageStr & TempImageStr & "</td><td><input name=""PopList"" type=""checkbox"" id=""News"&RsTempObj("Classid")&""" value="""& RsTempObj("Classid")&""""&iScheck&">" & RsTempObj("ClassCName") & "</td></tr></table></td></tr>"
		GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"),ImageStr)
		iScheck = ""
		RsTempObj.MoveNext
	loop
	Set RsTempObj = Nothing
End Function

if Request.Form("Action") = "Submit" then
	Dim Sql ,Rs
	Sql = "Select * from FS_AdminGroup where id = " & Request.Form("ID")
	Set Rs = Server.CreateObject(G_FS_RS)
	Rs.Open Sql,Conn,3,3
	Rs("PopList") = Request.Form("PopList")
	Rs.Update
	Rs.Close
	Set Rs = Nothing
end if
	Dim Sql1,Rs1
	Sql1 = "Select * from FS_AdminGroup where id="&request("ID")
	Set Rs1 = Server.CreateObject(G_FS_RS)
	Rs1.Open Sql1,Conn,3,3

%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>设置管理组权限</title>
</head>
<body topmargin="2" leftmargin="2"> 
<form name="PopForm" method="post" action=""> 
	<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999"> 
		<tr bgcolor="#EEEEEE"> 
			<td height="26" colspan="5" valign="middle"> <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0"> 
					<tr> 
						<td width=35 align="center" alt="设置" onClick="Modify();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">设置</td> 
						<td width=2 class="Gray">|</td> 
						<td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td> 
						<td>&nbsp; 
							<input name="id" type="hidden" id="id2" value="<%=request("id")%>"> 
							<input type="hidden" name="Action" value="Submit"> </td> 
					</tr> 
				</table></td> 
		</tr> 
	</table> 
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"> 
		<tr> 
			<td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
           
					
          <tr> 
						
            <td height="20" colspan="7"><div align="center"><strong>说明：</strong><font color="red"><strong>粗体为一级权限，红色为二级权限，黑色为三级权限</strong></font></div></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="<%=ShowRows1%>" align="center"><strong><font color=red> 
							
              <input name="PopList" type="checkbox"  value="P010000" <%if InStr(1, rs1("PopLIst"),"P010000" ,1)<>0 then response.Write("checked") %>>
               
							</font>信息管理</strong></td>
            <td width="14%"> <input  name="PopList" type="checkbox" value="P010100" <%if InStr(1, rs1("PopLIst"),"P010100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">添加栏目</font><br></td>
            <td width="14%"> <input  name="PopList" type="checkbox" value="P010200" <%if InStr(1, rs1("PopLIst"),"P010200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">修改栏目</font></td>
            <td width="14%"> <input  name="PopList" type="checkbox" value="P010300" <%if InStr(1, rs1("PopLIst"),"P010300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000"> 删除栏目</font></td>
            <td width="14%"> <input  name="PopList" type="checkbox" value="P010400" <%if InStr(1, rs1("PopLIst"),"P010400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">栏目操作</font></td>
            <td width="14%" height="20"><input  name="PopList" type="checkbox" value="P010513" <%if InStr(1, rs1("PopLIst"),"P010513" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000"> 栏目初始化</font></td>
            <td width="14%"><input  name="PopList" type="checkbox" value="P010514" <%if InStr(1, rs1("PopLIst"),"P010514" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000"> 新闻批量转移</font></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="3"><input  name="PopList" type="checkbox" value="P010500" <%if InStr(1, rs1("PopLIst"),"P010500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">新闻列表</font></td>
            <td height="20"><input  name="PopList" type="checkbox" value="P010501" <%if InStr(1, rs1("PopLIst"),"P010501" ,1)<>0 then response.Write("checked") %>>
               
							添加内容</td>
            <td><input  name="PopList" type="checkbox" value="P010502" <%if InStr(1, rs1("PopLIst"),"P010502" ,1)<>0 then response.Write("checked") %>>
               
							修改内容</td>
            <td><input  name="PopList" type="checkbox" value="P010503" <%if InStr(1, rs1("PopLIst"),"P010503" ,1)<>0 then response.Write("checked") %>>
               
							内容操作</td>
            <td><input  name="PopList" type="checkbox" value="P010504" <%if InStr(1, rs1("PopLIst"),"P010504" ,1)<>0 then response.Write("checked") %>>
               
							审核内容</td>
            <td><input  name="PopList" type="checkbox" value="P010505" <%if InStr(1, rs1("PopLIst"),"P010505" ,1)<>0 then response.Write("checked") %>>
               
							删除内容</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox" value="P010506" <%if InStr(1, rs1("PopLIst"),"P010506" ,1)<>0 then response.Write("checked") %>>
               
							内容搜索</td>
            <td><input  name="PopList" type="checkbox" value="P010507" <%if InStr(1, rs1("PopLIst"),"P010507" ,1)<>0 then response.Write("checked") %>>
               
							加入专题</td>
            <td><input  name="PopList" type="checkbox" value="P010508" <%if InStr(1, rs1("PopLIst"),"P010508" ,1)<>0 then response.Write("checked") %>>
               
							加入自由JS</td>
            <td><input  name="PopList" type="checkbox" value="P010509" <%if InStr(1, rs1("PopLIst"),"P010509" ,1)<>0 then response.Write("checked") %>>
               
							评论管理 </td>
            <td><input  name="PopList" type="checkbox" value="P010510" <%if InStr(1, rs1("PopLIst"),"P010510" ,1)<>0 then response.Write("checked") %>>
               
							生成</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox" value="P010511" <%if InStr(1, rs1("PopLIst"),"P010511" ,1)<>0 then response.Write("checked") %>>
               
							预览</td>
            <td><input  name="PopList" type="checkbox" value="P010512" <%if InStr(1, rs1("PopLIst"),"P010512" ,1)<>0 then response.Write("checked") %>>
               
							合并栏目</td>
            <td></td>
            <td>&nbsp;</td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2" valign="middle"><input  name="PopList" type="checkbox" value="P010700" <%if InStr(1, rs1("PopLIst"),"P010700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">下载列表</font></td>
            <td height="20"><input  name="PopList" type="checkbox" value="P010701" <%if InStr(1, rs1("PopLIst"),"P010701" ,1)<>0 then response.Write("checked") %>>
               
							添加内容</td>
            <td><input  name="PopList" type="checkbox" value="P010702" <%if InStr(1, rs1("PopLIst"),"P010702" ,1)<>0 then response.Write("checked") %>>
               
							修改内容</td>
            <td><input  name="PopList" type="checkbox" value="P010703" <%if InStr(1, rs1("PopLIst"),"P010703" ,1)<>0 then response.Write("checked") %>>
							审核内容</td>
            <td><input  name="PopList" type="checkbox" value="P010704" <%if InStr(1, rs1("PopLIst"),"P010704" ,1)<>0 then response.Write("checked") %>>
							删除内容</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox" value="P010705" <%if InStr(1, rs1("PopLIst"),"P010705" ,1)<>0 then response.Write("checked") %>>
               
							预览</td>
            <td><input  name="PopList" type="checkbox" value="P010706" <%if InStr(1, rs1("PopLIst"),"P010706" ,1)<>0 then response.Write("checked") %>>
               
							生成</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td></td>
          </tr>
           <%
			If HaveValueTF=1 then		
			%>
          <tr>
						
            <td rowspan="2" valign="middle"><input  name="PopList3" type="checkbox" value="P010800" <%if InStr(1, rs1("PopLIst"),"P010800" ,1)<>0 then response.Write("checked") %>>
                            <font color="#FF0000">商品列表</font></td>
            <td height="20"><input  name="PopList" type="checkbox" value="P010801" <%if InStr(1, rs1("PopLIst"),"P010801" ,1)<>0 then response.Write("checked") %>>
							添加内容</td>
            <td><input  name="PopList" type="checkbox" value="P010802" <%if InStr(1, rs1("PopLIst"),"P010802" ,1)<>0 then response.Write("checked") %>>
							修改内容</td>
            <td><input  name="PopList" type="checkbox" value="P010803" <%if InStr(1, rs1("PopLIst"),"P010803" ,1)<>0 then response.Write("checked") %>>
							删除内容</td>
            <td><input  name="PopList" type="checkbox" value="P010804" <%if InStr(1, rs1("PopLIst"),"P010804" ,1)<>0 then response.Write("checked") %>>
							加入专区</td>
            <td></td>
          </tr>
					
          <tr>
						
            <td height="20"><input  name="PopList" type="checkbox" value="P010805" <%if InStr(1, rs1("PopLIst"),"P010805" ,1)<>0 then response.Write("checked") %>>
							预览</td>
            <td><input  name="PopList" type="checkbox" value="P010806" <%if InStr(1, rs1("PopLIst"),"P010806" ,1)<>0 then response.Write("checked") %>>
							生成</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td></td>
          </tr>
		<%End If%>		
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox" value="P010600" <%if InStr(1, rs1("PopLIst"),"P010600" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">投稿管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P010601" <%if InStr(1, rs1("PopLIst"),"P010601" ,1)<>0 then response.Write("checked") %>>
               
							添加投稿</td>
            <td><input  name="PopList" type="checkbox" value="P010602" <%if InStr(1, rs1("PopLIst"),"P010602" ,1)<>0 then response.Write("checked") %>>
               
							修改投稿</td>
            <td><input  name="PopList" type="checkbox" value="P010603" <%if InStr(1, rs1("PopLIst"),"P010603" ,1)<>0 then response.Write("checked") %>>
               
							审核投稿</td>
            <td><input  name="PopList" type="checkbox" value="P010604" <%if InStr(1, rs1("PopLIst"),"P010604" ,1)<>0 then response.Write("checked") %>>
               
							删除投稿</td>
            <td></td>
          </tr>
           
					
          <tr> 
						 
						
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2" align="center" valign="middle"> <input name="PopList" type="checkbox" value="P020000" <%if InStr(1, rs1("PopLIst"),"P020000" ,1)<>0 then response.Write("checked") %>> 
							<strong>专题管理</strong></td>
            <td height="20"><input name="PopList" type="checkbox"  value="P020100" <%if InStr(1, rs1("PopLIst"),"P020100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">添加专题</font> </td>
            <td><input name="PopList" type="checkbox"  value="P020200" <%if InStr(1, rs1("PopLIst"),"P020200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">修改专题</font></td>
            <td><input name="PopList" type="checkbox"  value="P020300" <%if InStr(1, rs1("PopLIst"),"P020300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">删除专题</font></td>
            <td><input  name="PopList" type="checkbox" value="P020310" <%if InStr(1, rs1("PopLIst"),"P020310" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">专题初始化</font></td>
            <td><input  name="PopList" type="checkbox" value="P020320" <%if InStr(1, rs1("PopLIst"),"P020320" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">专题合并</font></td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input name="PopList" type="checkbox"  value="P020400" <%if InStr(1, rs1("PopLIst"),"P020400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">专题新闻管理</font> </td>
            <td><input name="PopList" type="checkbox"  value="P020401" <%if InStr(1, rs1("PopLIst"),"P020401" ,1)<>0 then response.Write("checked") %>>
               
							删除专题新闻</td>
            <td><input name="PopList" type="checkbox"  value="P020402" <%if InStr(1, rs1("PopLIst"),"P020402" ,1)<>0 then response.Write("checked") %>>
               
							专题新闻操作</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="<%=ShowRows2%>" align="center" valign="middle"><input name="PopList" type="checkbox"  value="P030000" <%if InStr(1, rs1("PopLIst"),"P030000" ,1)<>0 then response.Write("checked") %>> 
							<strong>站点管理</strong> </td>
            <td height="20"> <input name="PopList" type="checkbox"  value="P030100" <%if InStr(1, rs1("PopLIst"),"P030100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布首页</font></td>
            <td><input name="PopList" type="checkbox"  value="P030200" <%if InStr(1, rs1("PopLIst"),"P030200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布二级域名</font></td>
            <td><input name="PopList" type="checkbox"  value="P030300" <%if InStr(1, rs1("PopLIst"),"P030300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布栏目</font></td>
            <td><input name="PopList" type="checkbox"  value="P030400" <%if InStr(1, rs1("PopLIst"),"P030400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布新闻</font></td>
            <td><input  name="PopList" type="checkbox"  value="P030500" <%if InStr(1, rs1("PopLIst"),"P030500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布专题</font></td>
            <td><input  name="PopList" type="checkbox"  value="P030600" <%if InStr(1, rs1("PopLIst"),"P030600" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">发布下载</font></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input name="PopList" type="checkbox"  value="P030700" <%if InStr(1, rs1("PopLIst"),"P030700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">模板管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P030705" <%if InStr(1, rs1("PopLIst"),"P030705" ,1)<>0 then response.Write("checked") %>>
               
							编辑模板</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input name="PopList" type="checkbox"  value="P030800" <%if InStr(1, rs1("PopLIst"),"P030800" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">标签管理</font></td>
            <td><input name="PopList" type="checkbox"  value="P030801" <%if InStr(1, rs1("PopLIst"),"P030801" ,1)<>0 then response.Write("checked") %>>
               
							创建分类</td>
            <td><input  name="PopList" type="checkbox"  value="P030802" <%if InStr(1, rs1("PopLIst"),"P030802" ,1)<>0 then response.Write("checked") %>>
               
							新建标签</td>
            <td><input  name="PopList" type="checkbox"  value="P030803" <%if InStr(1, rs1("PopLIst"),"P030803" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P030804" <%if InStr(1, rs1("PopLIst"),"P030804" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P030805" <%if InStr(1, rs1("PopLIst"),"P030805" ,1)<>0 then response.Write("checked") %>>
               
							备份</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input name="PopList" type="checkbox"  value="P031300" <%if InStr(1, rs1("PopLIst"),"P031300" ,1)<>0 then response.Write("checked") %>>
							<font color="#FF0000">自由标签</font></td>
            <td><input  name="PopList" type="checkbox"  value="P031301" <%if InStr(1, rs1("PopLIst"),"P031301" ,1)<>0 then response.Write("checked") %>>
							新建标签</td>
            <td><input  name="PopList" type="checkbox"  value="P031302" <%if InStr(1, rs1("PopLIst"),"P031302" ,1)<>0 then response.Write("checked") %>>
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P031303" <%if InStr(1, rs1("PopLIst"),"P031303" ,1)<>0 then response.Write("checked") %>>
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P031304" <%if InStr(1, rs1("PopLIst"),"P031304" ,1)<>0 then response.Write("checked") %>>
							预览</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P030900" <%if InStr(1, rs1("PopLIst"),"P030900" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">备份管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P030901" <%if InStr(1, rs1("PopLIst"),"P030901" ,1)<>0 then response.Write("checked") %>>
               
							查看标签 </td>
            <td><input  name="PopList" type="checkbox"  value="P030902" <%if InStr(1, rs1("PopLIst"),"P030902" ,1)<>0 then response.Write("checked") %>>
               
							删除备份</td>
            <td><input  name="PopList" type="checkbox"  value="P030903" <%if InStr(1, rs1("PopLIst"),"P030903" ,1)<>0 then response.Write("checked") %>>
               
							还原标签</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P031000" <%if InStr(1, rs1("PopLIst"),"P031000" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">下载样式管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P031001" <%if InStr(1, rs1("PopLIst"),"P031001" ,1)<>0 then response.Write("checked") %>>
               
							新建样式</td>
            <td><input  name="PopList" type="checkbox"  value="P031002" <%if InStr(1, rs1("PopLIst"),"P031002" ,1)<>0 then response.Write("checked") %>>
               
							修改样式</td>
            <td><input  name="PopList" type="checkbox"  value="P031003" <%if InStr(1, rs1("PopLIst"),"P031003" ,1)<>0 then response.Write("checked") %>>
               
							查看样式 </td>
            <td><input  name="PopList" type="checkbox"  value="P031004" <%if InStr(1, rs1("PopLIst"),"P031004" ,1)<>0 then response.Write("checked") %>>
               
							删除样式</td>
            <td></td>
          </tr>
           
		          <%
		  If HaveValueTF=1 then
			  %>			
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P031200" <%if InStr(1, rs1("PopLIst"),"P031200" ,1)<>0 then response.Write("checked") %>>
               
							<font color="#FF0000">商城样式管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P031201" <%if InStr(1, rs1("PopLIst"),"P031201" ,1)<>0 then response.Write("checked") %>>
               
							新建样式</td>
            <td><input  name="PopList" type="checkbox"  value="P031202" <%if InStr(1, rs1("PopLIst"),"P031202" ,1)<>0 then response.Write("checked") %>>
               
							修改样式</td>
            <td><input  name="PopList" type="checkbox"  value="P031203" <%if InStr(1, rs1("PopLIst"),"P031203" ,1)<>0 then response.Write("checked") %>>
               
							查看样式 </td>
            <td><input  name="PopList" type="checkbox"  value="P031204" <%if InStr(1, rs1("PopLIst"),"P031204" ,1)<>0 then response.Write("checked") %>>
               
							删除样式</td>
            <td></td>
          </tr>
                    <%
		  End IF
			  %> 
					
          <tr> 
						
            <td><input  name="PopList" type="checkbox"  value="P031100" <%if InStr(1, rs1("PopLIst"),"P031100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">生成JS</font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="7" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P040000" <%if InStr(1, rs1("PopLIst"),"P040000" ,1)<>0 then response.Write("checked") %>> 
							<strong>系统管理</strong></td>
            <td height="20"> <input  name="PopList" type="checkbox"  value="P040100" <%if InStr(1, rs1("PopLIst"),"P040100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">管理员组管理 </font></td>
            <td><input  name="PopList" type="checkbox"  value="P040101" <%if InStr(1, rs1("PopLIst"),"P040101" ,1)<>0 then response.Write("checked") %>>
               
							新建管理员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040102" <%if InStr(1, rs1("PopLIst"),"P040102" ,1)<>0 then response.Write("checked") %>>
               
							修改管理员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040103" <%if InStr(1, rs1("PopLIst"),"P040103" ,1)<>0 then response.Write("checked") %>>
               
							删除管理员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040104" <%if InStr(1, rs1("PopLIst"),"P040104" ,1)<>0 then response.Write("checked") %>>
               
							设置权限</td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P040200" <%if InStr(1, rs1("PopLIst"),"P040200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">管理员管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P040201" <%if InStr(1, rs1("PopLIst"),"P040201" ,1)<>0 then response.Write("checked") %>>
               
							新建管理员</td>
            <td><input  name="PopList" type="checkbox"  value="P040202" <%if InStr(1, rs1("PopLIst"),"P040202" ,1)<>0 then response.Write("checked") %>>
               
							修改管理员</td>
            <td><input  name="PopList" type="checkbox"  value="P040203" <%if InStr(1, rs1("PopLIst"),"P040203" ,1)<>0 then response.Write("checked") %>>
               
							删除管理员</td>
            <td><input  name="PopList" type="checkbox"  value="P040204" <%if InStr(1, rs1("PopLIst"),"P040204" ,1)<>0 then response.Write("checked") %>>
               
							设置密码</td>
            <td><input  name="PopList" type="checkbox"  value="P040206" <%if InStr(1, rs1("PopLIst"),"P040206" ,1)<>0 then response.Write("checked") %>>
               
							锁定管理员</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P040300" <%if InStr(1, rs1("PopLIst"),"P040300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">会员组管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P040301" <%if InStr(1, rs1("PopLIst"),"P040301" ,1)<>0 then response.Write("checked") %>>
               
							新建会员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040302" <%if InStr(1, rs1("PopLIst"),"P040302" ,1)<>0 then response.Write("checked") %>>
               
							修改会员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040303" <%if InStr(1, rs1("PopLIst"),"P040303" ,1)<>0 then response.Write("checked") %>>
               
							删除会员组</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input name="PopList" type="checkbox"  value="P040400" <%if InStr(1, rs1("PopLIst"),"P040400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000"> 会员管理</font></td>
            <td><input name="PopList" type="checkbox"  value="P040401" <%if InStr(1, rs1("PopLIst"),"P040401" ,1)<>0 then response.Write("checked") %>>
               
							新建会员</td>
            <td><input name="PopList" type="checkbox"  value="P040402" <%if InStr(1, rs1("PopLIst"),"P040402" ,1)<>0 then response.Write("checked") %>>
               
							修改会员</td>
            <td><input  name="PopList" type="checkbox"  value="P040403" <%if InStr(1, rs1("PopLIst"),"P040403" ,1)<>0 then response.Write("checked") %>>
               
							删除会员</td>
            <td><input  name="PopList" type="checkbox"  value="P040404" <%if InStr(1, rs1("PopLIst"),"P040404" ,1)<>0 then response.Write("checked") %>>
               
							设置会员组</td>
            <td><input  name="PopList" type="checkbox"  value="P040405" <%if InStr(1, rs1("PopLIst"),"P040405" ,1)<>0 then response.Write("checked") %>>
               
							锁定会员</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P040500" <%if InStr(1, rs1("PopLIst"),"P040500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">系统参数</font></td>
            <td><input  name="PopList" type="checkbox"  value="P040501" <%if InStr(1, rs1("PopLIst"),"P040501" ,1)<>0 then response.Write("checked") %>>
               
							新闻系统参数</td>
            <td><input  name="PopList" type="checkbox"  value="P040502" <%if InStr(1, rs1("PopLIst"),"P040502" ,1)<>0 then response.Write("checked") %>>
               
							下载系统参数</td>
            <td><input  name="PopList" type="checkbox"  value="P040503" <%if InStr(1, rs1("PopLIst"),"P040503" ,1)<>0 then response.Write("checked") %>>
               
							系统常量设置</td> <%
			If HaveValueTF=1 then		
			%>
            <td><input  name="PopList" type="checkbox"  value="P040504" <%if InStr(1, rs1("PopLIst"),"P040504" ,1)<>0 then response.Write("checked") %>>
               				商参数设置</td>
							
			<%End If%>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2"><input  name="PopList" type="checkbox"  value="P040600" <%if InStr(1, rs1("PopLIst"),"P040600" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">数据库管理</font> </td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P0406001" <%if InStr(1, rs1("PopLIst"),"P0406001" ,1)<>0 then response.Write("checked") %>>
               
							数据统计</td>
            <td><input  name="PopList" type="checkbox"  value="P040602" <%if InStr(1, rs1("PopLIst"),"P040602" ,1)<>0 then response.Write("checked") %>>
               
							空间占用</td>
            <td><input  name="PopList" type="checkbox"  value="P040603" <%if InStr(1, rs1("PopLIst"),"P040603" ,1)<>0 then response.Write("checked") %>>
               
							数据库备份</td>
            <td><input  name="PopList" type="checkbox"  value="P040604" <%if InStr(1, rs1("PopLIst"),"P040604" ,1)<>0 then response.Write("checked") %>>
               
							数据库压缩</td>
            <td><input  name="PopList" type="checkbox"  value="P040605" <%if InStr(1, rs1("PopLIst"),"P040605" ,1)<>0 then response.Write("checked") %>>
               
							执行SQl语句</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P040606" <%if InStr(1, rs1("PopLIst"),"P040606" ,1)<>0 then response.Write("checked") %>>
               
							删除日志</td>
            <td><input  name="PopList" type="checkbox"  value="P040607" <%if InStr(1, rs1("PopLIst"),"P040607" ,1)<>0 then response.Write("checked") %>>
               
							后台日志管理</td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P050000" <%if InStr(1, rs1("PopLIst"),"P050000" ,1)<>0 then response.Write("checked") %>> 
							<strong>虚拟目录</strong></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P050100" <%if InStr(1, rs1("PopLIst"),"P050100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">新建目录</font></td>
            <td><input  name="PopList" type="checkbox"  value="P050200" <%if InStr(1, rs1("PopLIst"),"P050200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">删除目录</font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P050300" <%if InStr(1, rs1("PopLIst"),"P050300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">文件管理 </font></td>
            <td><input  name="PopList" type="checkbox"  value="P050301" <%if InStr(1, rs1("PopLIst"),"P050301" ,1)<>0 then response.Write("checked") %>>
               
							导入文件</td>
            <td><input  name="PopList" type="checkbox"  value="P050302" <%if InStr(1, rs1("PopLIst"),"P050302" ,1)<>0 then response.Write("checked") %>>
               
							删除文件</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="5" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P060000" <%if InStr(1, rs1("PopLIst"),"P060000" ,1)<>0 then response.Write("checked") %>> 
							<strong>JS管理</strong></td>
            <td height="20"> <input  name="PopList" type="checkbox"  value="P060100" <%if InStr(1, rs1("PopLIst"),"P060100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">新建自由JS</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060200" <%if InStr(1, rs1("PopLIst"),"P060200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">修改自由JS</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060300" <%if InStr(1, rs1("PopLIst"),"P060300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">删除自由JS</font></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P060400" <%if InStr(1, rs1("PopLIst"),"P060400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">自由JS新闻管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060401" <%if InStr(1, rs1("PopLIst"),"P060401" ,1)<>0 then response.Write("checked") %>>
               
							删除JS新闻</td>
            <td>&nbsp;</td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P060500" <%if InStr(1, rs1("PopLIst"),"P060500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">栏目JS管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060501" <%if InStr(1, rs1("PopLIst"),"P060501" ,1)<>0 then response.Write("checked") %>>
               
							新建</td>
            <td><input  name="PopList" type="checkbox"  value="P060502" <%if InStr(1, rs1("PopLIst"),"P060502" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P060503" <%if InStr(1, rs1("PopLIst"),"P060503" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P060600" <%if InStr(1, rs1("PopLIst"),"P060600" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">系统JS管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060601" <%if InStr(1, rs1("PopLIst"),"P060601" ,1)<>0 then response.Write("checked") %>>
               
							新建</td>
            <td><input  name="PopList" type="checkbox"  value="P060602" <%if InStr(1, rs1("PopLIst"),"P060602" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P060603" <%if InStr(1, rs1("PopLIst"),"P060603" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P060700" <%if InStr(1, rs1("PopLIst"),"P060700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">JS代码调用</font></td>
            <td><input  name="PopList" type="checkbox"  value="P060701" <%if InStr(1, rs1("PopLIst"),"P060701" ,1)<>0 then response.Write("checked") %>>
               
							栏目JS</td>
            <td><input  name="PopList" type="checkbox"  value="P060702" <%if InStr(1, rs1("PopLIst"),"P060702" ,1)<>0 then response.Write("checked") %>>
               
							系统JS</td>
            <td><input  name="PopList" type="checkbox"  value="P060703" <%if InStr(1, rs1("PopLIst"),"P060703" ,1)<>0 then response.Write("checked") %>>
               
							自由JS</td>
            <td><input  name="PopList" type="checkbox"  value="P060704" <%if InStr(1, rs1("PopLIst"),"P060704" ,1)<>0 then response.Write("checked") %>>
               
							广告JS</td>
            <td><input  name="PopList" type="checkbox"  value="P060705" <%if InStr(1, rs1("PopLIst"),"P060705" ,1)<>0 then response.Write("checked") %>>
               
							JS属性</td>
          </tr>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="11" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P070000" <%if InStr(1, rs1("PopLIst"),"P070000" ,1)<>0 then response.Write("checked") %>> 
							<strong>辅助管理</strong></td>
            <td height="20"> <input  name="PopList" type="checkbox"  value="P070100" <%if InStr(1, rs1("PopLIst"),"P070100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">回收站管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P070101" <%if InStr(1, rs1("PopLIst"),"P070101" ,1)<>0 then response.Write("checked") %>>
               
							还原</td>
            <td><input  name="PopList" type="checkbox"  value="P070102" <%if InStr(1, rs1("PopLIst"),"P070102" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P070103" <%if InStr(1, rs1("PopLIst"),"P070103" ,1)<>0 then response.Write("checked") %>>
               
							属性</td>
            <td><input  name="PopList" type="checkbox"  value="P070104" <%if InStr(1, rs1("PopLIst"),"P070104" ,1)<>0 then response.Write("checked") %>>
               
							请空回收站</td>
            <td><input  name="PopList" type="checkbox"  value="P070105" <%if InStr(1, rs1("PopLIst"),"P070105" ,1)<>0 then response.Write("checked") %>>
               
							还原回收站</td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2"> <input  name="PopList" type="checkbox"  value="P070200" <%if InStr(1, rs1("PopLIst"),"P070200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">广告管理</font></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P070201" <%if InStr(1, rs1("PopLIst"),"P070201" ,1)<>0 then response.Write("checked") %>>
               
							广告添加</td>
            <td><input  name="PopList" type="checkbox"  value="P070202" <%if InStr(1, rs1("PopLIst"),"P070202" ,1)<>0 then response.Write("checked") %>>
               
							广告修改</td>
            <td><input  name="PopList" type="checkbox"  value="P070203" <%if InStr(1, rs1("PopLIst"),"P070203" ,1)<>0 then response.Write("checked") %>>
               
							广告删除</td>
            <td><input  name="PopList" type="checkbox"  value="P070204" <%if InStr(1, rs1("PopLIst"),"P070204" ,1)<>0 then response.Write("checked") %>>
               
							暂停广告</td>
            <td><input  name="PopList" type="checkbox"  value="P070205" <%if InStr(1, rs1("PopLIst"),"P070205" ,1)<>0 then response.Write("checked") %>>
               
							激活广告</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P070206" <%if InStr(1, rs1("PopLIst"),"P070206" ,1)<>0 then response.Write("checked") %>>
               
							调用代码</td>
            <td><input  name="PopList" type="checkbox"  value="P070207" <%if InStr(1, rs1("PopLIst"),"P070207" ,1)<>0 then response.Write("checked") %>>
               
							显示统计</td>
            <td><input  name="PopList" type="checkbox"  value="P070208" <%if InStr(1, rs1("PopLIst"),"P070208" ,1)<>0 then response.Write("checked") %>>
               
							点击统计</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2"> <input  name="PopList" type="checkbox"  value="P070300" <%if InStr(1, rs1("PopLIst"),"P070300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">投票管理</font></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P070301" <%if InStr(1, rs1("PopLIst"),"P070301" ,1)<>0 then response.Write("checked") %>>
               
							新建投票</td>
            <td><input  name="PopList" type="checkbox"  value="P070302" <%if InStr(1, rs1("PopLIst"),"P070302" ,1)<>0 then response.Write("checked") %>>
               
							修改投票 </td>
            <td><input  name="PopList" type="checkbox"  value="P070303" <%if InStr(1, rs1("PopLIst"),"P070303" ,1)<>0 then response.Write("checked") %>>
               
							删除投票</td>
            <td><input  name="PopList" type="checkbox"  value="P070304" <%if InStr(1, rs1("PopLIst"),"P070304" ,1)<>0 then response.Write("checked") %>>
               
							激活投票</td>
            <td><input  name="PopList" type="checkbox"  value="P070305" <%if InStr(1, rs1("PopLIst"),"P070305" ,1)<>0 then response.Write("checked") %>>
               
							暂停投票 </td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P070306" <%if InStr(1, rs1("PopLIst"),"P070306" ,1)<>0 then response.Write("checked") %>>
               
							查看投票 </td>
            <td><input  name="PopList" type="checkbox"  value="P070307" <%if InStr(1, rs1("PopLIst"),"P070307" ,1)<>0 then response.Write("checked") %>>
               
							获取代码</td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P070400" <%if InStr(1, rs1("PopLIst"),"P070400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">常规管理</font> </td>
            <td><input  name="PopList" type="checkbox"  value="P070401" <%if InStr(1, rs1("PopLIst"),"P070401" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P070402" <%if InStr(1, rs1("PopLIst"),"P070402" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P070403" <%if InStr(1, rs1("PopLIst"),"P070403" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P070500" <%if InStr(1, rs1("PopLIst"),"P070500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">友情链接管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P070501" <%if InStr(1, rs1("PopLIst"),"P070501" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P070502" <%if InStr(1, rs1("PopLIst"),"P070502" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P070503" <%if InStr(1, rs1("PopLIst"),"P070503" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P080600" <%if InStr(1, rs1("PopLIst"),"P080600" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">插件管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P080601" <%if InStr(1, rs1("PopLIst"),"P080601" ,1)<>0 then response.Write("checked") %>>
               
							新建插件</td>
            <td><input  name="PopList" type="checkbox"  value="P080602" <%if InStr(1, rs1("PopLIst"),"P080602" ,1)<>0 then response.Write("checked") %>>
               
							修改插件</td>
            <td><input  name="PopList" type="checkbox"  value="P080603" <%if InStr(1, rs1("PopLIst"),"P080603" ,1)<>0 then response.Write("checked") %>>
               
							删除插件</td>
            <td><input  name="PopList" type="checkbox"  value="P080604" <%if InStr(1, rs1("PopLIst"),"P080604" ,1)<>0 then response.Write("checked") %>>
               
							显示插件</td>
            <td><input  name="PopList" type="checkbox"  value="P080605" <%if InStr(1, rs1("PopLIst"),"P080605" ,1)<>0 then response.Write("checked") %>>
               
							隐藏插件</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P070606" <%if InStr(1, rs1("PopLIst"),"P070606" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">归档管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P070607" <%if InStr(1, rs1("PopLIst"),"P070607" ,1)<>0 then response.Write("checked") %>>
               
							生成</td>
            <td><input  name="PopList" type="checkbox"  value="P070608" <%if InStr(1, rs1("PopLIst"),"P070608" ,1)<>0 then response.Write("checked") %>>
               
							预览</td>
            <td><input  name="PopList" type="checkbox"  value="P070609" <%if InStr(1, rs1("PopLIst"),"P070609" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P070700" <%if InStr(1, rs1("PopLIst"),"P070700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">留言管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P070701" <%if InStr(1, rs1("PopLIst"),"P070701" ,1)<>0 then response.Write("checked") %>>
               
							留言添加</td>
            <td><input  name="PopList" type="checkbox"  value="P070702" <%if InStr(1, rs1("PopLIst"),"P070702" ,1)<>0 then response.Write("checked") %>>
               
							留言查看</td>
            <td><input  name="PopList" type="checkbox"  value="P070703" <%if InStr(1, rs1("PopLIst"),"P070703" ,1)<>0 then response.Write("checked") %>>
               
							留言修改</td>
            <td><input  name="PopList" type="checkbox"  value="P070704" <%if InStr(1, rs1("PopLIst"),"P070704" ,1)<>0 then response.Write("checked") %>>
               
							留言删除</td>
            <td><input  name="PopList" type="checkbox"  value="P070705" <%if InStr(1, rs1("PopLIst"),"P070705" ,1)<>0 then response.Write("checked") %>>
               
							留言回复</td>
          </tr>
           
										
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P070800" <%if InStr(1, rs1("PopLIst"),"P070800" ,1)<>0 then response.Write("checked") %>>
               
							<font color="#FF0000">帮助管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P070801" <%if InStr(1, rs1("PopLIst"),"P070801" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P070802" <%if InStr(1, rs1("PopLIst"),"P070802" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P070803" <%if InStr(1, rs1("PopLIst"),"P070803" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P070804" <%if InStr(1, rs1("PopLIst"),"P070804" ,1)<>0 then response.Write("checked") %>>
               
							检索</td>
            <td><input  name="PopList" type="checkbox"  value="P070805" <%if InStr(1, rs1("PopLIst"),"P070805" ,1)<>0 then response.Write("checked") %>>
               
							查看</td>
          </tr>
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="8" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P080000" <%if InStr(1, rs1("PopLIst"),"P080000" ,1)<>0 then response.Write("checked") %>> 
							<strong>辅助工具</strong></td>
            <td rowspan="2"> <input  name="PopList" type="checkbox"  value="P080100" <%if InStr(1, rs1("PopLIst"),"P080100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">采集站点</font></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P080101" <%if InStr(1, rs1("PopLIst"),"P080101" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P080102" <%if InStr(1, rs1("PopLIst"),"P080102" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P080103" <%if InStr(1, rs1("PopLIst"),"P080103" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P080104" <%if InStr(1, rs1("PopLIst"),"P080104" ,1)<>0 then response.Write("checked") %>>
               
							复制</td>
            <td><input  name="PopList" type="checkbox"  value="P080105" <%if InStr(1, rs1("PopLIst"),"P080105" ,1)<>0 then response.Write("checked") %>>
               
							向导</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P080106" <%if InStr(1, rs1("PopLIst"),"P080106" ,1)<>0 then response.Write("checked") %>>
               
							采集</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P080200" <%if InStr(1, rs1("PopLIst"),"P080200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">采集关键字</font></td>
            <td><input  name="PopList" type="checkbox"  value="P080201" <%if InStr(1, rs1("PopLIst"),"P080201" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P080202" <%if InStr(1, rs1("PopLIst"),"P080202" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P080203" <%if InStr(1, rs1("PopLIst"),"P080203" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P080300" <%if InStr(1, rs1("PopLIst"),"P080300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">采集新闻</font></td>
            <td><input  name="PopList" type="checkbox"  value="P080301" <%if InStr(1, rs1("PopLIst"),"P080301" ,1)<>0 then response.Write("checked") %>>
               
							修改 </td>
            <td><input  name="PopList" type="checkbox"  value="P080302" <%if InStr(1, rs1("PopLIst"),"P080302" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P080303" <%if InStr(1, rs1("PopLIst"),"P080303" ,1)<>0 then response.Write("checked") %>>
               
							入库</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P080400" <%if InStr(1, rs1("PopLIst"),"P080400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">采集历史</font></td>
            <td><input  name="PopList" type="checkbox"  value="P080401" <%if InStr(1, rs1("PopLIst"),"P080401" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P080402" <%if InStr(1, rs1("PopLIst"),"P080402" ,1)<>0 then response.Write("checked") %>>
               
							入库</td>
            <td><input  name="PopList" type="checkbox"  value="P080403" <%if InStr(1, rs1("PopLIst"),"P080403" ,1)<>0 then response.Write("checked") %>>
               
							删除全部</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2"> <input  name="PopList" type="checkbox"  value="P080500" <%if InStr(1, rs1("PopLIst"),"P080500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">流量统计</font></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P080501" <%if InStr(1, rs1("PopLIst"),"P080501" ,1)<>0 then response.Write("checked") %>>
               
							获取代码</td>
            <td><input  name="PopList" type="checkbox"  value="P080502" <%if InStr(1, rs1("PopLIst"),"P080502" ,1)<>0 then response.Write("checked") %>>
               
							网站维护</td>
            <td><input  name="PopList" type="checkbox"  value="P080503" <%if InStr(1, rs1("PopLIst"),"P080503" ,1)<>0 then response.Write("checked") %>>
               
							简要数据</td>
            <td><input  name="PopList" type="checkbox"  value="P080504" <%if InStr(1, rs1("PopLIst"),"P080504" ,1)<>0 then response.Write("checked") %>>
               
							24小时统计 </td>
            <td><input  name="PopList" type="checkbox"  value="P080505" <%if InStr(1, rs1("PopLIst"),"P080505" ,1)<>0 then response.Write("checked") %>>
               
							日统计</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P080506" <%if InStr(1, rs1("PopLIst"),"P080506" ,1)<>0 then response.Write("checked") %>>
               
							月统计 </td>
            <td><input  name="PopList" type="checkbox"  value="P080507" <%if InStr(1, rs1("PopLIst"),"P080507" ,1)<>0 then response.Write("checked") %>>
               
							系统/浏览器</td>
            <td><input  name="PopList" type="checkbox"  value="P080508" <%if InStr(1, rs1("PopLIst"),"P080508" ,1)<>0 then response.Write("checked") %>>
               
							地区统计</td>
            <td><input  name="PopList" type="checkbox"  value="P080509" <%if InStr(1, rs1("PopLIst"),"P080509" ,1)<>0 then response.Write("checked") %>>
               
							来源统计</td>
            <td><input  name="PopList" type="checkbox"  value="P080510" <%if InStr(1, rs1("PopLIst"),"P080510" ,1)<>0 then response.Write("checked") %>>
               
							来访者信息统计</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P080700" <%if InStr(1, rs1("PopLIst"),"P080700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">字段内容替换</font></td>
            <td><input  name="PopList" type="checkbox"  value="P080800" <%if InStr(1, rs1("PopLIst"),"P080800" ,1)<>0 then response.Write("checked") %>>
               
							<font color="#FF0000">DW插件辅助</font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <%
		  If HaveValueTF=1 then
			  %>
           
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="7" align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P090000" <%if InStr(1, rs1("PopLIst"),"P090000" ,1)<>0 then response.Write("checked") %>> 
							<strong>B2C商城</strong></td>
            <td> <input  name="PopList" type="checkbox"  value="P090100" <%if InStr(1, rs1("PopLIst"),"P090100" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">专区管理</font></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P090101" <%if InStr(1, rs1("PopLIst"),"P090101" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P090102" <%if InStr(1, rs1("PopLIst"),"P090102" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P090103" <%if InStr(1, rs1("PopLIst"),"P090103" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P090104" <%if InStr(1, rs1("PopLIst"),"P090104" ,1)<>0 then response.Write("checked") %>>
               
							查看商品</td>
            <td>&nbsp;</td>
          </tr>
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P090200" <%if InStr(1, rs1("PopLIst"),"P090200" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">厂家管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P090201" <%if InStr(1, rs1("PopLIst"),"P090201" ,1)<>0 then response.Write("checked") %>>
               
							添加</td>
            <td><input  name="PopList" type="checkbox"  value="P090202" <%if InStr(1, rs1("PopLIst"),"P090202" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P090203" <%if InStr(1, rs1("PopLIst"),"P090203" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P090300" <%if InStr(1, rs1("PopLIst"),"P090300" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">定单管理</font></td>
            <td><input  name="PopList" type="checkbox"  value="P090301" <%if InStr(1, rs1("PopLIst"),"P090301" ,1)<>0 then response.Write("checked") %>>
               
							查看 </td>
            <td><input  name="PopList" type="checkbox"  value="P090302" <%if InStr(1, rs1("PopLIst"),"P090302" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td><input  name="PopList" type="checkbox"  value="P090303" <%if InStr(1, rs1("PopLIst"),"P090303" ,1)<>0 then response.Write("checked") %>>
               
							改变支付状态</td>
            <td><input  name="PopList" type="checkbox"  value="P090304" <%if InStr(1, rs1("PopLIst"),"P090304" ,1)<>0 then response.Write("checked") %>>
               
							改变出货状态</td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P090400" <%if InStr(1, rs1("PopLIst"),"P090400" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">积分/金币</font></td>
            <td><input  name="PopList" type="checkbox"  value="P090401" <%if InStr(1, rs1("PopLIst"),"P090401" ,1)<>0 then response.Write("checked") %>>
               
							查看</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td rowspan="2"> <input  name="PopList" type="checkbox"  value="P090500" <%if InStr(1, rs1("PopLIst"),"P090500" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">综合统计</font></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20"><input  name="PopList" type="checkbox"  value="P090600" <%if InStr(1, rs1("PopLIst"),"P090600" ,1)<>0 then response.Write("checked") %>>
               
							<font color="#FF0000">在线支付设置</font></td>
            <td><input  name="PopList" type="checkbox"  value="P090800" <%if InStr(1, rs1("PopLIst"),"P090800" ,1)<>0 then response.Write("checked") %>>
               
							<font color="#FF0000">配送须知</font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
           
					
          <tr> 
						
            <td height="20"> <input  name="PopList" type="checkbox"  value="P090700" <%if InStr(1, rs1("PopLIst"),"P090700" ,1)<>0 then response.Write("checked") %>> 
							<font color="#FF0000">邮寄资料</font></td>
            <td><input  name="PopList" type="checkbox"  value="P090701" <%if InStr(1, rs1("PopLIst"),"P090701" ,1)<>0 then response.Write("checked") %>>
               
							修改</td>
            <td><input  name="PopList" type="checkbox"  value="P090702" <%if InStr(1, rs1("PopLIst"),"P090702" ,1)<>0 then response.Write("checked") %>>
               
							删除</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
                    <%
			  Set PConfigObj = Nothing
		  End if
		  %>
            
					
          <tr> 
						
            <td colspan="7" align="center" valign="top"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td align="center" valign="middle"><input  name="PopList" type="checkbox"  value="P999999" <%if InStr(1, rs1("PopLIst"),"P999999" ,1)<>0 then response.Write("checked") %>> 
							<strong>公共权限</strong></td>
            <td height="20"><input  name="PopList" type="checkbox"  value="P990100" <%if InStr(1, rs1("PopLIst"),"P990100" ,1)<>0 then response.Write("checked") %>>
               
							新建文件</td>
            <td><input  name="PopList" type="checkbox"  value="P990200" <%if InStr(1, rs1("PopLIst"),"P990200" ,1)<>0 then response.Write("checked") %>>
               
							新建目录</td>
            <td><input  name="PopList" type="checkbox"  value="P990300" <%if InStr(1, rs1("PopLIst"),"P990300" ,1)<>0 then response.Write("checked") %>>
               
							导入文件</td>
            <td><input  name="PopList" type="checkbox"  value="P990400" <%if InStr(1, rs1("PopLIst"),"P990400" ,1)<>0 then response.Write("checked") %>>
               
							删除目录文件</td>
            <td></td>
            <td></td>
          </tr>
           
					
          <tr> 
						
            <td height="20" colspan="7"><hr style="width:95%;"></td>
          </tr>
           
					
          <tr> 
						
            <td><input name="PopList" type="checkbox" value="P000000" <%if InStr(1, rs1("PopLIst"),"P000000" ,1)<>0 then response.Write("checked") %>> 
							<strong>新闻权限</strong></td>
            <td height="20"></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
           
			</table></td> 
		</tr> 
	</table> 
	<table border="0" cellspacing="0" cellpadding="0"> 
		<tr> 
			<td> </td> 
		</tr> 
		<%
do while Not RsClassObj.Eof
	Dim iScheck1
	if InStr(1,rs1("PopLIst"),RsClassObj("Classid"),1)<>0 then iScheck1=" checked"
%> 
		<tr> 
			<td><table border="0" cellspacing="0" cellpadding="0"> 
					<tr align="left" class="TempletItem"> 
						<td><img src="../../Images/Folder/folderclosed.gif"></td> 
						<td><input name="PopList" type="checkbox" value="<% = RsClassObj("Classid") %>"<%=iScheck1%>> 
							<% = RsClassObj("ClassCName") %></td> 
					</tr> 
				</table></td> 
		</tr> 
		<%
	iScheck1 = ""
	Response.Write(GetChildClassList(RsClassObj("ClassID"),""))
	RsClassObj.MoveNext
loop
%> 
	</table> 
</form> 
<%
rs1.close
set rs1=nothing
%> 
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function Modify(){if(confirm("您确定要修改?")){document.PopForm.submit();}return false;}
function ChangeSelectState(Obj)
{
	var CheckAllChild=false;
	if (Obj.checked==true) CheckAllChild=true;
	var PopListObj=document.body.getElementsByTagName('INPUT');
	for (var i=0;i<PopListObj.length;i++)
	{
		CurrObj=PopListObj(i);
		if (CurrObj.ParentID==Obj.ClassID)
		{
			CurrObj.checked=CheckAllChild;
		}
	}
}
function CheckParent(ParentID)
{
	var PopListObj=document.body.getElementsByTagName('INPUT');
	for (var i=0;i<PopListObj.length;i++)
	{
		CurrObj=PopListObj(i);
		if (CurrObj.ClassID==ParentID)
		{
			CurrObj.checked=GetParentObjCheckedTF(ParentID);
			return true;
		}
	}
}
function GetParentObjCheckedTF(ParentID)
{
	var CurrObj=null;
	var PopListObj=document.body.getElementsByTagName('INPUT');
	for (var i=0;i<PopListObj.length;i++)
	{
		CurrObj=PopListObj(i);
		if (CurrObj.ParentID==ParentID)
		{
			if (CurrObj.checked==true) return true;
		}
	}
	return false;
}
</script>
