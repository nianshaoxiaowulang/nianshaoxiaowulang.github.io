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
'��ƷĿ¼����Ѷ��ƷNϵ��
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System V1.0.0
'���¸��£�2004.8
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071 
'����֧��QQ��66252421 
'���򿪷�����Ѷ������ & ��Ѷ���������
'Email:service@cooin.com
'��̳֧�֣���Ѷ������̳(http://bbs.cooin.com   http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺www.cooin.com    ������԰�أ�www.aspsun.cn
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ���߱����˳���ķ���׷��Ȩ��
'==============================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
'Ȩ������
if Not JudgePopedomTF(Session("Name"),"P030802") and Not JudgePopedomTF(Session("Name"),"P030803")  then
 	Call ReturnError1()
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../../CSS/ModeWindow.css">
<title>���ò�ѯ����</title>
</head>
<body topmargin="0">
<%
Dim SqlStr,QueryNum,OrderArray,OrderCNameArray,FieldsArray,FieldsCNameArray,TempArray,FieldObj,TempRs,TempFieldCName,TableCName
Dim NewsSql,NewsRs,NewsClassSql,NewsClassRs,DownloadSql,DownloadRs,flag,i,j,indexOfField
'News��Download��NewsClass����ֶ�����������
SqlStr = Request("SqlStr")
TempArray = split(SqlStr,"*")
'�����������ֱ𱣴浽��ѯ�������ֶΡ��ֶ���������������
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
'���������е��ֶ�Ӣ���������������
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
				OrderCNameArray(i) = "����."&NewsFieldName(indexOfField)
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
				OrderCNameArray(i) = "��Ŀ."&NewsClassFieldName(indexOfField)
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
				OrderCNameArray(i) = "����."&DownloadFieldName(indexOfField)
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
           <td width="39%" align="center">�ֶδ���</td>
           <td width="6%">&nbsp;</td>
           <td align="center">�������</td>
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
                  <td><input name="" onclick="MoveList('FieldsList','Up');" id="operation3" type="button" value="������"></td>
                </tr>
                <tr> 
                  <td><input name="" onclick="MoveList('FieldsList','Down');" id="operation4" type="button" value="������"></td>
                </tr>
              </table>
			</td>
            <td width="42%"> <select size="18" name="OrderList" id="OrderList" style="width:100%">
                <%
				for i = 0 to UBound(OrderArray)
					if OrderArray(i) <> "" then
				%>
                <option value="<%=OrderArray(i)%>"> 
                <%If InStr(OrderArray(i),"Desc") = 0 then Response.Write("�� "&OrderCNameArray(i)) else Response.Write("�� "&OrderCNameArray(i)) end if%>
                </option>
                <%
					end if
				next
				%>
              </select> </td>
            <td> <table>
                <tr> 
                  <td><input name="" onclick="MoveList('OrderList','Up');" id="operation3" type="button" value="������"></td>
                </tr>
                <tr> 
                  <td><input name="" onclick="MoveList('OrderList','Down');" id="operation4" type="button" value="������"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
  	<td width="74%" height="35">&nbsp;&nbsp;��ѯ����: 
      <input name="QueryNum" size ="12" type="text" value="<%=QueryNum%>"></td>
	<td width="14%">
		<input name="" type="button" onclick="DoSubmit();" value="ȷ��">
	</td>
	<td width="12%">
		<input name="" type="button" onclick="window.close();" value="ȡ��">
	</td>
  </tr>
        </form>
</table>
</body>
</html>
<script>
window.returnValue = "";
//��ָ���б��ѡ�������ָ����������ƶ�
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
//���б�Ͳ�ѯ��������������ɽ������
function DoSubmit()
{
	var OrderStr = "",FieldsStr="",i;
	var OrderListObj=document.FreeLableForm.OrderList;
	var FieldsListObj=document.FreeLableForm.FieldsList;
	if(IsNumeric(document.FreeLableForm.QueryNum.value) == false)
	{
		alert("��ѯ�����������!");
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
//�ж��ַ����Ƿ�Ϊ�ջ�������
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