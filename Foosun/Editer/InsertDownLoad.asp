<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim TempClassListStr
TempClassListStr = ClassList
Function ClassList()
	Dim Rs
	Set Rs = Conn.Execute("select ClassID,ClassCName,ClassEName from FS_newsclass where ParentID = '0' and DelFlag=0 order by AddTime desc")
	do while Not Rs.Eof
		ClassList = ClassList & "<option value="&Rs("ClassID")&"" & ">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum,ClassEName from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>插入下载</title>
</head>
<link rel="stylesheet" type="text/css" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body topmargin="0" leftmargin="0" scroll=no>
<div align="center">
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"> <div align="left">栏目列表 
          <select onChange="ChangeDownLoadList(this.value);" style="width:75%;" name="select">
            <option value="" selected>栏目选择</option>
            <% =TempClassListStr %>
          </select>
        </div></td>
      <td height="30"> <div align="right">下载列表 
          <select onChange="SetDownAddress(this.value);" style="width:75%;" name="SelectDownLoad">
            <option value="">选择下载</option>
            <%
		  Dim DownLoadObj
		  Set DownLoadObj = Conn.Execute("Select * from FS_DownLoad")
		  do while Not DownLoadObj.Eof
		  %>
            <option Name="<% = DownLoadObj("Name") %>" value="<% = DownLoadObj("DownLoadID") %>" ClassID="<% = DownLoadObj("ClassID") %>"> 
            <% = DownLoadObj("Name") %>
            </option>
            <%
		  	DownLoadObj.MoveNext
		  Loop
		  Set DownLoadObj = Nothing
		  %>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="52"> <div align="left">地址列表</div></td>
              <td> <select onclick="SetAddress(this.value);" name="SelectDownAddress" style="width:480;" size="5">
                </select> </td>
            </tr>
          </table>
        </div></td>
    </tr>
    <tr> 
      <td height="36" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="52"> <div align="left">文件地址</div></td>
            <td><input type="text" name="FilePath" style="width:80%;"> <input type="button" onClick="OpenWindowAndSetValue('../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.all.FilePath);" name="Submit3" value="选择文件"> 
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="30"><div align="center"> 
          <input type="button" onClick="GetReturnValue();" name="Submit" value=" 确 定 ">
        </div></td>
      <td height="30"><div align="center"> 
          <input type="button" name="Submit2" onClick="window.close();" value=" 取 消 ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Dim DownAddressObj,IDStr,DownLoadStr,UrlStr,NameStr
Set DownAddressObj = Conn.Execute("Select * from FS_DownLoadAddress")
do while Not DownAddressObj.Eof
	if IDStr = "" then
		IDStr = DownAddressObj("ID")
	else
		IDStr = IDStr & "$$$" & DownAddressObj("ID")
	end if
	if DownLoadStr = "" then
		DownLoadStr = DownAddressObj("DownLoadID")
	else
		DownLoadStr = DownLoadStr & "$$$" & DownAddressObj("DownLoadID")
	end if
	if UrlStr = "" then
		UrlStr = DownAddressObj("Url")
	else
		UrlStr = UrlStr & "$$$" & DownAddressObj("Url")
	end if
	if NameStr = "" then
		NameStr = DownAddressObj("AddressName")
	else
		NameStr = NameStr & "$$$" & DownAddressObj("AddressName")
	end if
	DownAddressObj.MoveNext
Loop
Set DownAddressObj = Nothing
%>
<script language="JavaScript">
var IDStr='<% = IDStr %>';
var DownLoadIDStr='<% = DownLoadStr %>';
var UrlStr='<% = UrlStr %>';
var NameStr='<% = NameStr %>';
var DownLoadOptionArray=new Array();
var DownAddressObjArray=new Array();
function DownAddressObj(ID,DownLoadID,Url,Name)
{
	this.ID=ID;
	this.DownLoadID=DownLoadID;
	this.Url=Url;
	this.Name=Name;
}
setTimeout('InitailObjArray()',100);
function InitailObjArray()
{
	var IDArray=IDStr.split('$$$'),DownLoadIDArray=DownLoadIDStr.split('$$$'),UrlArray=UrlStr.split('$$$'),NameArray=NameStr.split('$$$');
	for (var i=0;i<IDArray.length;i++)
	{
		DownAddressObjArray[DownAddressObjArray.length]=new DownAddressObj(IDArray[i],DownLoadIDArray[i],UrlArray[i],NameArray[i]);
	}
	for (var i=0;i<document.all.SelectDownLoad.length;i++)
	{
		if (document.all.SelectDownLoad.options(i).ClassID!=null)	DownLoadOptionArray[DownLoadOptionArray.length]=document.all.SelectDownLoad.options(i);
	}
}
function SetDownAddress(DownLoadID)
{
	document.all.FilePath.value='';
	DeleteAllOption(document.all.SelectDownAddress);
	for (var i=0;i<DownAddressObjArray.length;i++)
	{
		if (DownAddressObjArray[i].DownLoadID==DownLoadID)
		{
			AddFolderList(document.all.SelectDownAddress,DownAddressObjArray[i].Name+'||'+DownAddressObjArray[i].Url,DownAddressObjArray[i].Url);
		}
	}
}
function AddFolderList(SelectObj,Lable,LableContent)
{
	var i=0,AddOption;
	if (!SearchOptionExists(SelectObj,Lable))
	{
		AddOption = document.createElement("OPTION");
		AddOption.text=Lable;
		AddOption.value=LableContent;
		SelectObj.add(AddOption);
		//SelectObj.options(SelectObj.length-1).selected=true;
	}
}
function SearchOptionExists(Obj,SearchText)
{
	var i;
	for(i=0;i<Obj.length;i++)
	{
		if (Obj.options(i).text==SearchText)
		{
			return true;
		}
	}
	return false;
}
function DeleteAllOption(Obj)
{
	var OptionLength=Obj.length;
	for (var i=0;i<OptionLength;i++)
	{
		Obj.options.remove(Obj.length-1);
	}
}
function GetReturnValue()
{
	if (document.all.FilePath.value!='')
	{
		window.returnValue=document.all.FilePath.value;
		window.close();
	}
	else alert('请选择地址或者填写地址');
}
function ChangeDownLoadList(ClassID)
{
	DeleteAllOption(document.all.SelectDownAddress);
	var i=0,OptionLength=document.all.SelectDownLoad.length,Index=0;
	for (i=0;i<OptionLength;i++)
	{
		Index=document.all.SelectDownLoad.length-1;
		if (document.all.SelectDownLoad.options(Index).ClassID!=null) document.all.SelectDownLoad.options.remove(Index);
	}
	document.all.FilePath.value='';
	if (ClassID!='')
	{
		for (i=0;i<DownLoadOptionArray.length;i++)
		{
			if (DownLoadOptionArray[i].ClassID==ClassID)
			{
				//alert(DownLoadOptionArray[i].ClassID);
				//return;
				document.all.SelectDownLoad.add(DownLoadOptionArray[i]);
			}
		}
	}
	else
	{
		for (i=0;i<DownLoadOptionArray.length;i++)
		{
			document.all.SelectDownLoad.add(DownLoadOptionArray[i]);
		}
	}
	document.all.SelectDownLoad.options(0).selected=true;
}
function SetAddress(FileUrl)
{
	if (FileUrl!='')
	{
		document.all.FilePath.value=FileUrl;
	}
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
<%
Set Conn = Nothing
%>