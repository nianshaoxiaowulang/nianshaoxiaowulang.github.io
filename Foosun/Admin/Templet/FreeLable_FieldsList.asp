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
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'============================================================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<%
'Ȩ������
if Not JudgePopedomTF(Session("Name"),"P030802") and Not JudgePopedomTF(Session("Name"),"P030803") then
 	Call ReturnError1()
end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.selectedTab {
	background-color: #0000FF;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" onselectstart="return false;">
<%
Dim ITableName,TableName,TableCName,FieldObj,i,TableField,TempCNameArray,TempENameArray,TempTypeArray,RsObj,indexOfField
ITableName = Request("TableName")
'���ݲ���ѡ��ͬ�����ݱ����Ӧ��Ϣ
Select Case(Lcase(ITableName))
	Case "fs_news"
		TableName = "FS_News"
		TableCName = "����"
		TempCNameArray = NewsFieldName
		TempENameArray = NewsFieldEName
		TempTypeArray = NewsFieldType
		TableField = "News_Fields"
		Set RsObj = Conn.Execute("Select * from FS_News where 1=0")
	Case "fs_newsclass"
		TableName = "FS_NewsClass"
		TableCName = "��Ŀ"
		TempCNameArray = NewsClassFieldName
		TempENameArray = NewsClassFieldEName
		TempTypeArray = NewsClassFieldType
		TableField = "NewsClass_Fields"
		Set RsObj = Conn.Execute("Select * from FS_NewsClass where 1=0")
	Case "fs_download"
		TableName = "FS_Download"
		TableCName = "����"
		TempCNameArray = DownloadFieldName
		TempENameArray = DownloadFieldEName
		TempTypeArray = DownloadFieldType
		TableField = "Download_Fields"
		Set RsObj = Conn.Execute("Select * from FS_Download where 1=0")
	Case Else
		Response.End	
End Select
%>
<table width=98%>
  <%
	i = 0
	for Each FieldObj In RsObj.Fields
		indexOfField = GetIndexOfField(FieldObj.name,TempENameArray)
%>
	<tr>
	<td height=23 width="34%"><span ondblclick="SelectTab(<%=i%>)" onclick="ClickTab(<%=i%>)" value="<%=TableName&"."&FieldObj.name%>" alt="<%=TempTypeArray(indexOfField)%>" id="<%=TableField&i%>" class="TempletItem"><% = TempCNameArray(indexOfField)%></span></td>
	<td width="11%"><% = GetFieldType(TempTypeArray(indexOfField))%></td>
	<td width="18%"><div id ="Direction_<%=i%>" align=center></div></td>
	<td width="17%"><div id ="Operation_<%=i%>" align=right></div></td>
	<td width="20%"><div id ="Value_<%=i%>"></div></td>
	</tr>
<%
	i = i + 1
	Next
	Set RsObj = Nothing
%>
</table>
</body>
</html>
<script language="JavaScript">
var IsSqlDataBase = <%=IsSqlDataBase%>;
var SelectIndex = -1;
var TableName = "<%=TableName%>";
var	TableCName = "<%=TableCName%>";
var TableField = "<%=TableField%>";
var FieldsNum = <%=i%>;
var DirectionStrArray=new Array(FieldsNum);
var OperationStrArray=new Array(3*FieldsNum);
var ValueStrArray=new Array(3*FieldsNum);
var OpNumOfFieldArray= new Array(FieldsNum);
var i;
//��ʼ���������Ͳ���ֵ����
for(i=0;i<FieldsNum*3;i++)
{
	if(i<FieldsNum)
	{
		DirectionStrArray[i] = "";
		OpNumOfFieldArray[i] = 1;
	}
	OperationStrArray[i] = "";
	ValueStrArray[i] = "";
}
//��ʼ��ѡ���ֶΡ����ʽ��������������л�����������ʾ״̬���ɸ������ڳ�ʼ��ʱ����
function Initial(FieldArray,ExpArray,OrderArray)
{
	var OpArray = new Array("<>",">=","<=","=",">","<"," In");
	var i,j,FieldName,FieldAlt,ExpStr,OpStr,OperationStr,ValueStr,indexSearch;
	for(i=0;i<FieldsNum;i++)
	{
		FieldName = document.all(TableField+i).value;
		FieldAlt = document.all(TableField+i).alt;
		
		if(SearchStrInArray(FieldArray,FieldName,0) != -1)
			document.all(TableField+i).className = "TempletSelectItem";
		indexSearch = SearchStrInArray(OrderArray,FieldName,0);
		if(indexSearch != -1)
		{
			if(OrderArray[indexSearch].indexOf("Desc") != -1)
			{
				DirectionStrArray[i] = "Desc";
				document.all("Direction_"+i).innerHTML = "����";
			}
			else
			{
				DirectionStrArray[i] = "Asc";
				document.all("Direction_"+i).innerHTML = "����";
			}
		}
		indexSearch = SearchStrInArray(ExpArray,FieldName,0);
		while(indexSearch != -1)
		{
			ExpStr = ExpArray[indexSearch].replace(FieldName,"");
			for(j=0;j<OpArray.length;j++)
				if(ExpStr.indexOf(OpArray[j]) == 0)
				{
					OpStr = OpArray[j];
					ExpStr = ExpStr.replace(OpStr,"");
				}
			if(OpNumOfFieldArray[i] < 3)
			{
				if(OperationArray(i,OpNumOfFieldArray[i],"Get","") == "")
				{
					OperationArray(i,OpNumOfFieldArray[i],"Set",OpStr.replace(" ",""))
				}
				else
				{
					OpNumOfFieldArray[i]++;
					OperationArray(i,OpNumOfFieldArray[i],"Set",OpStr.replace(" ",""))
				}
				
				if(FieldAlt == "116" || FieldAlt == "16")
				{
					ValueArray(i,OpNumOfFieldArray[i],"Set",ExpStr.replace("(","").replace(")",""));
				}
				else
				{
					if(FieldAlt == "7")
					{
						ExpStr = ExpStr.replace(/\#/g,"").replace("(","").replace(")","");
						ValueArray(i,OpNumOfFieldArray[i],"Set",ExpStr);
					}
					else
					{
						if(OpStr == " In")
						{
							ExpStr = ExpStr.replace(/\'/g,"").replace("(","");
							ExpStr = ExpStr.substr(0,ExpStr.length-1);
							ValueArray(i,OpNumOfFieldArray[i],"Set",ExpStr);
						}
						else
							ValueArray(i,OpNumOfFieldArray[i],"Set",ExpStr.replace(/\'/g,""));
					}
				}
	
			}
			indexSearch = SearchStrInArray(ExpArray,FieldName,indexSearch+1);
		}
		OperationStr = "";
		ValueStr = "";
		for(j=1;j<=OpNumOfFieldArray[i];j++)
		{
			OpStr = OperationArray(i,j,"Get","");
			if(OpStr != "")
			{
				Tempstr = ValueArray(i,j,"Get","");
				if(Tempstr.length > 10)
					Tempstr = Tempstr.substring(0,7)+"...";
				if(OperationStr == "")
				{
					OperationStr = OpStr.replace("<","&lt;");
					ValueStr = Tempstr;
				}
				else
				{
					OperationStr = OperationStr+"<br>"+OpStr.replace("<","&lt;");
					ValueStr = ValueStr+"<br>"+Tempstr;
				}
			}
		}
		document.all("Operation_"+i).innerHTML = OperationStr;
		document.all("Value_"+i).innerHTML = ValueStr;
	}
}
//�ڸ�����������м����ַ�����StartIndexΪ��ʼ���������
function SearchStrInArray(ArrayObj,Str,StartIndex)
{
	var ReturnVal=-1,i=0;
	if(StartIndex<0) StartIndex = 0;
	for(i=StartIndex;i<ArrayObj.length;i++)
		if(ArrayObj[i].indexOf(Str) != -1)
		{
			ReturnVal = i;
			break;
		}
	return ReturnVal;
} 
//��OperationStrArray����Ϊ����ģ��һ����ά���飬��ȡ�򱣴������
function OperationArray(index,col,operation,value)
{
	switch(operation.toLowerCase())
	{
		case "get":
			return OperationStrArray[(col-1)*FieldsNum+index];
		case "set":
			OperationStrArray[(col-1)*FieldsNum+index] = value;
	}
}
//��ValueStrArray����Ϊ����ģ��һ����ά���飬��ȡ�򱣴����ֵ
function ValueArray(index,col,operation,value)
{
	switch(operation.toLowerCase())
	{
		case "get":
			return ValueStrArray[(col-1)*FieldsNum+index];
		case "set":
			ValueStrArray[(col-1)*FieldsNum+index] = value;
	}
}
//ѡ�л�ȡ��ѡ��ĳ���ֶΣ����ø����ڵ���ӻ�ɾ��������������ʾ״̬
function SelectTab(index)
{
	if(document.all(TableField+index).className != "TempletSelectItem")
	{
		document.all(TableField+index).className = "TempletSelectItem";
		parent.AddField(document.all(TableField+index).value,TableCName+"."+document.all(TableField+index).innerText);
	}
	else
	{
		document.all(TableField+index).className = "TempletItem";
		parent.RemoveField(document.all(TableField+index).value);
	}
}
//��������ѡ���ֶε�ѡ��״̬���ɸ����ڵ���
function CleanSelected()
{
	var i;
	for(i=0;i<FieldsNum;i++)
	{
		document.all(TableField+i).className = "TempletItem";
	}
}
//�ֶε��������仯����Ӧ���ڸ�����������������
function DirectionChange(index)
{
	DirectionStrArray[index] = document.all("Dir_"+index).value;
	if(document.all("Dir_"+index).value == "")
		parent.SetOrderToArray(document.all(TableField+index).value);
	else
		parent.SetOrderToArray(document.all(TableField+index).value+" "+document.all("Dir_"+index).value);
}
//�����������仯����Ӧ���ڸ����������ò������ʽ
function OperationChange(OperationNum,index)
{
	var i,j,OperationListObj;
	OperationListObj = document.all("Op"+OperationNum+"_"+index);
	if(OperationListObj.value != "")
	{
		for(i=1;i<=3;i++)
		{
			if(OperationNum == i) continue;
			if(OperationListObj.value == OperationArray(index,i,"Get",""))
			{
				alert("ͬһ�ֶ��в����������ظ�");
				for(j=0;j<OperationListObj.options.length;j++)
					if(OperationListObj.options(j).value == OperationArray(index,OperationNum,"Get",""))
						OperationListObj.options(j).selected = true;
				return;
			}
		}
	}
	else
	{
		document.all("Val"+OperationNum+"_"+index).value = "";
	}
	OperationArray(index,OperationNum,"Set",OperationListObj.value);
	CheckValue(OperationNum,index);
	if(document.all("Val"+OperationNum+"_"+index).value != "")
		SetExpressionToArray();
}
//����û��ڲ���ֵ���������Ч��
function CheckValue(OperationNum,index)
{
	var keyCode = event.keyCode;
	if(keyCode == 37 || keyCode == 39 || keyCode == 8) return;
	var AltStr = document.all(TableField+index).alt;
	var OpStr = document.all("Op"+OperationNum+"_"+index).value;
	var ValObj = document.all("Val"+OperationNum+"_"+index);
	if(keyCode==13) SetExpressionToArray();
	if(	AltStr == "116" || AltStr == "16")
	{
		if(OpStr != "In")
		{
			if(keyCode<48 || keyCode>57)
				ValObj.value = CleanStrExcept(ValObj.value,"0123456789");
		}
		else
		{
			if(keyCode<48 || keyCode>57 || keyCode!=188)
				ValObj.value = CleanStrExcept(ValObj.value,"0123456789,");
		}
	}
	else
	{
		if(AltStr == "7")
		{
			if(OpStr != "In")
			{
				if(keyCode<48 || keyCode>57 || keyCode!=189 || keyCode!=109 || keyCode!=191 || keyCode!=111)
					ValObj.value = CleanStrExcept(ValObj.value,"0123456789-/");
			}
			else
			{
				if(keyCode<48 || keyCode>57 || keyCode!=189 || keyCode!=109 || keyCode!=191 || keyCode!=111|| keyCode!=188)
					ValObj.value = CleanStrExcept(ValObj.value,"0123456789-/,");
			}
		}
		else
		{
			if(OpStr != "In")
			{
				if(keyCode==188)
					ValObj.value = ValObj.value.replace(/\,/g,"");
			}
			if(keyCode==222)
				ValObj.value = ValObj.value.replace(/\'/g,"");
		}
	}
	ValueArray(index,OperationNum,"Set",ValObj.value);
}
//����������ʽ���Ը����ڵĲ������ʽ���и���
function SetExpressionToArray()
{
	var ExpArray = new Array();
	var i,j,k,OpStr,FieldStr,AltStr,ValStr,DateArray;
	for(i=0;i<FieldsNum;i++)
	{
		FieldStr = document.all(TableField+i).value;
		AltStr = document.all(TableField+i).alt;
		for(j=1;j<=3;j++)
		{
			if(	OperationArray(i,j,"Get","") != "")
			{
				OpStr =	OperationArray(i,j,"Get","");
				ValStr = ValueArray(i,j,"Get","");
				if(AltStr == "7")
				{
					DateArray = ValStr.split(",");
					for(k=0;k<DateArray.length;k++)
						if(IsValidDate(DateArray[k]) == false)
						{
							alert("�ֶ�["+document.all(TableField+i).innerText+"]"+OpStr+" "+"��������Ч");
							return;
						}
				}
				ExpArray[ExpArray.length] = CreateExpression(FieldStr,AltStr,OpStr,ValStr);
			}
		}
	}
	parent.SetExpressionToArray(ExpArray,TableName);
}
//���ɲ������ʽ��
function CreateExpression(FieldStr,AltStr,OpStr,ValStr)
{
	var ExpressionStr = "";
	if(AltStr == "116" || AltStr == "16" )
	{
		if(ValStr == "")
			ValStr = "0";
		if(OpStr == "In")
			ExpressionStr = FieldStr+" In("+ValStr.replace(/\'/g,"")+")";
		else
			ExpressionStr = FieldStr+OpStr+ValStr;
	}
	else
	{
		if(AltStr == "7")
		{
			if(OpStr == "In")
			{
				if(IsSqlDataBase == 1)
				{
					ExpressionStr = FieldStr+" In('"+ValStr.replace(/\'/g,"").replace(/\,/g,"','")+"')";
				}
				else
				{
					ExpressionStr = FieldStr+" In(#"+ValStr.replace(/\'/g,"").replace(/\,/g,"#,#")+"#)";
				}
			}
			else
			{
				if(IsSqlDataBase == 1)
				{
					ExpressionStr = FieldStr+OpStr+"'"+ValStr+"'";
				}
				else
				{
					ExpressionStr = FieldStr+OpStr+"#"+ValStr+"#";
				}
			}
		}
		else
		{
			if(OpStr == "In")
				ExpressionStr = FieldStr+" In('"+ValStr.replace(/\'/g,"").replace(/\,/g,"','")+"')";
			else
				ExpressionStr = FieldStr+OpStr+"'"+ValStr+"'";
		}
	}
	return ExpressionStr;
}
//�ж��ַ����Ƿ�Ϊ��Ч����
function IsValidDate(DateStr)
{
	DateStr = DateStr.replace(/\-/g,",").replace(/\//g,",")
	var DeteStrArray = DateStr.split(",");
	var DayNumOfMonth = new Array(0,31,28,31,30,31,30,31,31,30,31,30,31);
	if(DeteStrArray.length == 1 || DeteStrArray.length > 3) return false;
	if(DeteStrArray[0] == "" || DeteStrArray[0].length > 4) return false;
	if(DeteStrArray[1] == "" || eval(DeteStrArray[1]+"> 12") == true) return false;
	if(DeteStrArray.length == 3)
	{
		if(eval("("+DeteStrArray[0]+" % 4 == 0 && "+DeteStrArray[0]+" % 100 != 0) || "+DeteStrArray[0]+" % 400 == 0") == true)
			DayNumOfMonth[2] = 29;
		if(eval(DeteStrArray[2]+" > DayNumOfMonth["+DeteStrArray[1]+"]") == true || DeteStrArray[2] == "")
			return false;
	}
	return true;
}
//����ַ������ָ���ַ�����ַ�
function CleanStrExcept(Str,ExceptStr)
{
	if(ExceptStr == "") return "";
	var i=0;
	while(i<Str.length)
	{
		if(ExceptStr.indexOf(Str.charAt(i)) == -1)			
			Str = Str.substr(0,i)+Str.substr(i+1);
		else
			i = i+1;
	}
	return Str;
}
//�û�����ĳ���ֶΣ�����ǰһ�������ֶεĲ�����������ֵ������ǰһ���ͱ��ε����ֶε���ʾ״̬�͸����ڵ���Ӳ������ʽ������
function ClickTab(index)
{
	var i,DirectionStr,OperationStr,ValueStr,Tempstr;
	DirectionStr = "";
	OperationStr = "";
	ValueStr = "";
	Tempstr = "";
	if(SelectIndex != -1)
	{
		if(SelectIndex == index) return;
		for(i=1;i<=OpNumOfFieldArray[SelectIndex];i++)
		{
			if(OperationArray(SelectIndex,i,"Get","") != "")
			{
				if(document.all(TableField+SelectIndex).alt == "116" && document.all("Val"+i+"_"+SelectIndex).value =="")
					document.all("Val"+i+"_"+SelectIndex).value = "0";
				ValueArray(SelectIndex,i,"Set",document.all("Val"+i+"_"+SelectIndex).value);
				Tempstr = ValueArray(SelectIndex,i,"Get","");
				if(Tempstr.length > 10)
					Tempstr = Tempstr.substring(0,7)+"...";
				if(OperationStr == "")
				{
					OperationStr = OperationArray(SelectIndex,i,"Get","").replace("<","&lt;");
					ValueStr = Tempstr;
				}
				else
				{
					OperationStr = OperationStr+"<br>"+OperationArray(SelectIndex,i,"Get","").replace("<","&lt;");
					ValueStr = ValueStr+"<br>"+Tempstr;
				}
			}
		}
		if(DirectionStrArray[SelectIndex] == "Asc")
			DirectionStr = "����";
			else
				if(DirectionStrArray[SelectIndex] == "Desc")
					DirectionStr = "����";
			
		document.all("Direction_"+SelectIndex).innerHTML = DirectionStr;
		document.all("Operation_"+SelectIndex).innerHTML = OperationStr
		document.all("Value_"+SelectIndex).innerHTML = ValueStr;
		SetExpressionToArray();
	}
	SelectIndex = index;
	if(SelectIndex != -1)
	{
		parent.SetAddExpContainerHTML(TableName,"���� <a href='#' onclick=AddExpression('"+TableName+"',"+index+")>+</a>");
		DirectionStr = "<select id='Dir_"+SelectIndex+"' onchange='DirectionChange("+SelectIndex+")' style='width:100%;'><option value=''><option value='Asc'>����<option value='Desc'>����</select>";
		DirectionStr = DirectionStr.replace("value='"+DirectionStrArray[SelectIndex]+"'","value='"+DirectionStrArray[SelectIndex]+"' selected");
		document.all("Direction_"+SelectIndex).innerHTML = DirectionStr;
		document.all("Operation_"+SelectIndex).innerHTML = CreateHMTL("Operation",SelectIndex);
		document.all("Value_"+SelectIndex).innerHTML = CreateHMTL("Value",SelectIndex);
	}
	else
		parent.SetAddExpContainerHTML(TableField,"����");
}
//��ĳ���ֶ���Ӳ������ʽ��������ͬһ�ֶ��������3���������ʽ
function AddExpression(index)
{
	var OperationStr,ValueStr,i;
	if(OpNumOfFieldArray[index]<3)
	{
		OperationStr = CreateHMTL("Operation",SelectIndex);
		ValueStr = CreateHMTL("Value",SelectIndex);
		
		i = OpNumOfFieldArray[index] +1;
		OperationStr = OperationStr+"<br><select id='Op"+i+"_"+index+"' onchange='OperationChange("+i+","+index+")' style='width:100%;'><option value=''><option value='='>=<option value='<>'>&lt;&gt;<option value='>'>&gt;<option value='>='>&gt;=<option value='<'>&lt;<option value='<='>&lt;=<option value='In'>In</select>";
		ValueStr = ValueStr+"<br><input id='Val"+i+"_"+index+"' value='' onKeyUp='CheckValue("+i+","+index+")' size=8>";

		OpNumOfFieldArray[index] = i;

		document.all("Operation_"+SelectIndex).innerHTML = OperationStr;
		document.all("Value_"+SelectIndex).innerHTML = ValueStr;
	}
}
//������������ĳ���ֶεĲ�����ѡ���б�����ֵ������HTML����
function CreateHMTL(Name,index)
{
	var i,j,TempOprStr,TempValueStr,OperationStr,ValueStr;
	j=0;
	OperationStr ="";
	for(i=1;i<=3;i++)
	{
		if(OperationArray(index,i,"Get","") !="")
		{
			j++;			
			TempOprStr = "<select id='Op"+j+"_"+index+"' onchange='OperationChange("+j+","+index+")' style='width:100%;'><option value=''><option value='='>=<option value='<>'>&lt;&gt;<option value='>'>&gt;<option value='>='>&gt;=<option value='<'>&lt;<option value='<='>&lt;=<option value='In'>In</select>";
			TempOprStr = TempOprStr.replace("value='"+OperationArray(index,i,"Get","")+"'","value='"+OperationArray(index,i,"Get","")+"' selected")
			TempValueStr = "<input id='Val"+j+"_"+index+"' value='"+ValueArray(index,i,"Get","")+"' onKeyUp='CheckValue("+j+","+index+")' size=8>";
			if(OperationStr == "")
			{
				OperationStr = TempOprStr;
				ValueStr = TempValueStr;
			}						
			else
			{
				OperationStr = OperationStr+"<br>"+TempOprStr;
				ValueStr = ValueStr+"<br>"+TempValueStr;
			}
			continue;
		}
	}		
	if(OperationStr == "")
	{
		OperationStr = "<select id='Op"+1+"_"+index+"' onchange='OperationChange("+1+","+index+")' style='width:100%;'><option value=''><option value='='>=<option value='<>'>&lt;&gt;<option value='>'>&gt;<option value='>='>&gt;=<option value='<'>&lt;<option value='<='>&lt;=<option value='In'>In</select>";
		ValueStr = "<input id='Val"+1+"_"+index+"' value='' onKeyUp='CheckValue("+1+","+index+")' size=8>";
	}
	if(j==0)
		OpNumOfFieldArray[index] = 1;
	else
		OpNumOfFieldArray[index] = j;
		
	for(i=1;i<=3;i++)
	{
		if(OperationArray(index,i,"Get","") == "")
		{
			for(j=i+1;j<=3;j++)
				if(	OperationArray(index,j,"Get","") != "")
				{
					OperationArray(index,i,"Set",OperationArray(index,j,"Get",""))
					OperationArray(index,j,"Set","")
					ValueArray(index,i,"Set",ValueArray(index,j,"Get",""))
					ValueArray(index,j,"Set","")
				}	
		}
	}
	switch(Name)
	{
		case "Operation":
			return OperationStr;
		case "Value":
			return ValueStr;

	}
}
</script>
<%
Set Conn = Nothing

Function GetFieldType(FieldType)
	Select Case FieldType
		Case 0
			GetFieldType = "Empty"
		Case 16
			GetFieldType = "�Զ�"
		Case 100
			GetFieldType = "�ı�"
		Case 116
			GetFieldType = "����"
		Case 230
			GetFieldType = "��ע"
		Case 2
			GetFieldType = "SmallInt"
		Case 3
			GetFieldType = "Integer"
		Case 20
			GetFieldType = "BigInt"
		Case 17
			GetFieldType = "UnsignedTinyInt"
		Case 18
			GetFieldType = "UnsignedSmallInt"
		Case 19
			GetFieldType = "UnsignedInt"
		Case 21
			GetFieldType = "UnsignedBigInt"
		Case 4
			GetFieldType = "Single"
		Case 5
			GetFieldType = "Double"
		Case 6
			GetFieldType = "Currency"
		Case 14
			GetFieldType = "Decimal"
		Case 131
			GetFieldType = "Numeric"
		Case 11
			GetFieldType = "Boolean"
		Case 10
			GetFieldType = "Error"
		Case 132
			GetFieldType = "UserDefined"
		Case 12
			GetFieldType = "Variant"
		Case 9
			GetFieldType = "IDispatch"
		Case 13
			GetFieldType = "IUnknown"
		Case 72
			GetFieldType = "GUID"
		Case 7
			GetFieldType = "����"
		Case 133
			GetFieldType = "DBDate"
		Case 134
			GetFieldType = "DBTime"
		Case 135
			GetFieldType = "DBTimeStamp"
		Case 8
			GetFieldType = "BSTR"
		Case 129
			GetFieldType = "Char"
		Case 200
			GetFieldType = "VarChar"
		Case 201
			GetFieldType = "LongVarChar"
		Case 130
			GetFieldType = "WChar"
		Case 202
			GetFieldType = "VarWChar"
		Case 203
			GetFieldType = "LongVarWChar"
		Case 128
			GetFieldType = "Binary"
		Case 204
			GetFieldType = "VarBinary"
		Case 205
			GetFieldType = "LongVarBinary"
		Case 136
			GetFieldType = "Chapter"
		Case 64
			GetFieldType = "FileTime"
		Case 138
			GetFieldType = "PropVariant"
		Case 139
			GetFieldType = "VarNumeric"
		Case &H2000
			GetFieldType = "Array"
	End Select
End Function
%>