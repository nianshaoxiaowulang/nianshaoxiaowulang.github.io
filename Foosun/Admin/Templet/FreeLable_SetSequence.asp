<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/checkPopedom.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../Inc/FieldConst.asp" -->
<% 
Dim  DBC,Conn,TempClassListStr,TempListStr
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
'==============================================================================
'产品目录：风讯产品N系列
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System V1.0.0
'最新更新：2004.8
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071 
'技术支持QQ：66252421 
'程序开发：风讯开发组 & 风讯插件开发组
'Email:service@cooin.com
'论坛支持：风讯在线论坛(http://bbs.cooin.com   http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：www.cooin.com    开发者园地：www.aspsun.cn
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'风讯在线保留此程序的法律追究权利
'==============================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
'权限判限
if Not JudgePopedomTF(Session("Name"),"P030802") and Not JudgePopedomTF(Session("Name"),"P030803")  then
 	Call ReturnError1()
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../../CSS/ModeWindow.css">
<title>设置查询排序</title>
</head>
<body topmargin="0">
<%
Dim SqlStr,QueryNum,OrderArray,OrderCNameArray,FieldsArray,FieldsCNameArray,TempArray,FieldObj,TempRs,TempFieldCName,TableCName
Dim NewsSql,NewsRs,NewsClassSql,NewsClassRs,DownloadSql,DownloadRs,flag,i,j,indexOfField
'News、Download、NewsClass表的字段中文名数组
SqlStr = Request("SqlStr")
TempArray = split(SqlStr,"*")
'解析参数，分别保存到查询数量、字段、字段名和排序项数组
QueryNum = TempArray(0)
OrderArray = split(TempArray(1),",")
FieldsArray = split(TempArray(2),",")
FieldsCNameArray = split(TempArray(3),",")
OrderCNameArray = OrderArray
NewsSql = "select * from FS_news where 1=0"
set NewsRs = conn.Execute(NewsSql)
NewsClassSql = "select * from FS_newsclass where 1=0"
set NewsClassRs = conn.Execute(NewsClassSql)
DownloadSql = "select * from FS_download where 1=0"
set DownloadRs = conn.Execute(DownloadSql)
'把排序项中的字段英文名翻译成中文名
for i = 0 to UBound(OrderArray)
'response.write(OrderArray(1))
'response.end
	flag = false
	if inStr(OrderArray(i),"FS_News.") <> 0 Then
		j = 0
		for each FieldObj in NewsRs.Fields
			if Trim(Mid(Replace(OrderArray(i)," Desc",""),9)) = Trim(FieldObj.Name) Then
				indexOfField = GetIndexOfField(FieldObj.Name,NewsFieldEName)
				flag = true
				OrderCNameArray(i) = "新闻."&NewsFieldName(indexOfField)
				Exit for
			end if
			j = j + 1
		Next
	elseif inStr(OrderArray(i),"FS_NewsClass.") <> 0 Then
		j = 0
		for each FieldObj in NewsClassRs.Fields
			if Trim(Mid(Replace(OrderArray(i)," Desc",""),14)) = Trim(FieldObj.Name) Then
				indexOfField = GetIndexOfField(FieldObj.Name,NewsClassFieldEName)
				flag = true
				OrderCNameArray(i) = "栏目."&NewsClassFieldName(indexOfField)
				Exit for
			end if
			j = j + 1
		Next
	elseif inStr(OrderArray(i),"FS_Download.") <> 0 Then
		j = 0
		for each FieldObj in DownloadRs.Fields
			if Trim(Mid(Replace(OrderArray(i)," Desc",""),13)) = Trim(FieldObj.Name) Then
				indexOfField = GetIndexOfField(FieldObj.Name,DownloadFieldEName)
				flag = true
				OrderCNameArray(i) = "下载."&DownloadFieldName(indexOfField)
				Exit for
			end if
			j = j + 1
		Next
	End if
	if flag = false then
		OrderArray(i) = ""
		OrderCNameArray(i) = ""
	end if
next
set NewsRs = nothing
set NewsClassRs = nothing
set DownloadRs = nothing
%>
<table width="100%" height="322" border="0" cellpadding="0" cellspacing="0">
  <form action="" method="post" name="FreeLableForm">
  <tr> 
    <td height="285" colspan="3"> 
      <table width="100%" border="0" cellpadding="4" cellspacing="1">
          <tr> 
           <td width="39%" align="center">字段次序</td>
           <td width="6%">&nbsp;</td>
           <td align="center">排序次序</td>
           <td width="13%">&nbsp;</td>
          </tr>
          <tr> 
		   <td align="center">
		   		 <select size="18" name="FieldsList" id="FieldsList" style="width:100%">
                <%
				for i = 0 to UBound(FieldsArray)
					if FieldsArray(i) <> "" then
				%>
                <option value="<%=FieldsArray(i)%>"><%=FieldsCNameArray(i)%></option>
                <%
					end if
				next
				%>
              </select> 
		   </td>
           <td width="6%">
		   <table>
                <tr> 
                  <td><input name="" onclick="MoveList('FieldsList','Up');" id="operation3" type="button" value="↑上移"></td>
                </tr>
                <tr> 
                  <td><input name="" onclick="MoveList('FieldsList','Down');" id="operation4" type="button" value="↓下移"></td>
                </tr>
              </table>
			</td>
            <td width="42%"> <select size="18" name="OrderList" id="OrderList" style="width:100%">
                <%
				for i = 0 to UBound(OrderArray)
					if OrderArray(i) <> "" then
				%>
                <option value="<%=OrderArray(i)%>"> 
                <%If InStr(OrderArray(i),"Desc") = 0 then Response.Write("↑ "&OrderCNameArray(i)) else Response.Write("↓ "&OrderCNameArray(i)) end if%>
                </option>
                <%
					end if
				next
				%>
              </select> </td>
            <td> <table>
                <tr> 
                  <td><input name="" onclick="MoveList('OrderList','Up');" id="operation3" type="button" value="↑上移"></td>
                </tr>
                <tr> 
                  <td><input name="" onclick="MoveList('OrderList','Down');" id="operation4" type="button" value="↓下移"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
  	<td width="74%" height="35">&nbsp;&nbsp;查询数量: 
      <input name="QueryNum" size ="12" type="text" value="<%=QueryNum%>"></td>
	<td width="14%">
		<input name="" type="button" onclick="DoSubmit();" value="确定">
	</td>
	<td width="12%">
		<input name="" type="button" onclick="window.close();" value="取消">
	</td>
  </tr>
        </form>
</table>
</body>
</html>
<script>
window.returnValue = "";
//对指定列表的选中项根据指定方向进行移动
function MoveList(ListName,dir)
{
	var valueStr,textStr,index;
	var ListObj=document.all(ListName);
	index = ListObj.selectedIndex;
	
	if(index == -1) return;
	
	switch(dir)
	{
		case "Up":
			if(ListObj.selectedIndex != 0)
			{
				valueStr = ListObj.options(index-1).value;
				textStr = ListObj.options(index-1).text;
				ListObj.options(index-1).value = ListObj.options(index).value;
				ListObj.options(index-1).text = ListObj.options(index).text;
				valueStr = ListObj.options(index).value = valueStr;
				textStr = ListObj.options(index).text = textStr;
				ListObj.options(index-1).selected = true;
			}
			break;
		case "Down":
			if(ListObj.selectedIndex != ListObj.length-1)
			{
				valueStr = ListObj.options(index+1).value;
				textStr = ListObj.options(index+1).text;
				ListObj.options(index+1).value = ListObj.options(index).value;
				ListObj.options(index+1).text = ListObj.options(index).text;
				valueStr = ListObj.options(index).value = valueStr;
				textStr = ListObj.options(index).text = textStr;
				ListObj.options(index+1).selected = true;
			}
			break;
	}
}
//从列表和查询数量输入框中生成结果返回
function DoSubmit()
{
	var OrderStr = "",FieldsStr="",i;
	var OrderListObj=document.FreeLableForm.OrderList;
	var FieldsListObj=document.FreeLableForm.FieldsList;
	if(IsNumeric(document.FreeLableForm.QueryNum.value) == false)
	{
		alert("查询数量输入错误!");
		document.FreeLableForm.QueryNum.focus();
		return;
	}
	for(i=0;i<OrderListObj.options.length;i++)
		if(OrderStr == "")
			OrderStr = OrderListObj.options(i).value;
		else
			OrderStr = OrderStr + "," + OrderListObj.options(i).value;
	for(i=0;i<FieldsListObj.options.length;i++)
		if(FieldsStr == "")
		{
			FieldsStr = FieldsListObj.options(i).value;
			FieldsNameStr = FieldsListObj.options(i).text;
		}
		else
		{
			FieldsStr = FieldsStr + "," + FieldsListObj.options(i).value;
			FieldsNameStr = FieldsNameStr + "," + FieldsListObj.options(i).text;
		}

	window.returnValue = document.FreeLableForm.QueryNum.value+"*"+OrderStr+"*"+FieldsStr+"*"+FieldsNameStr;
	window.close();
}
//判断字符串是否为空或正整数
function IsNumeric(Str)
{
	var i,NumericStr="0123456789";
	if(Str=="") return true;
	if(Str.charAt(0) =="0") return false;
	for(i=0;i<Str.length;i++)
		if(NumericStr.indexOf(Str.substr(i,1)) == -1)
			return false;
	return true;
}
</script>