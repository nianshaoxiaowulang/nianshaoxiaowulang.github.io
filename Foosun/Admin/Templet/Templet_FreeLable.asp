<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
	Dim DBC,Conn
	On Error Resume Next
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P031300") then Call ReturnError1()
Dim LableSql,RsLableObj,RsLblTypeObj,LableID

LableSql = "Select name,freelableid,addtime,stylecontent,description from FS_FreeLable"
Set RsLableObj = Server.CreateObject(G_FS_RS)
RSLableObj.open LableSql,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签列表</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
function AddFreeLable()
{
	location.href = 'FreeLable_Edit.asp';
}
function FreeLableObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFreeLable();",'新建','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFreeLable();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFreeLable();",'删除','disabled');
}
function ContentMenuShowEvent()
{
	ChangeLableMenuStatus();
}
function ChangeLableMenuStatus()
{
	var EventObjInArray=false,SelectFolder='',SelectFile='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.TypeID;
				else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.TypeID
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectFile=='') SelectFile=ListObjArray[i].Obj.LableID;
				else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.LableID
			}
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Obj.TypeID!=null)
					{
						if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.TypeID;
						else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.TypeID
					}
					if (ListObjArray[i].Obj.LableID!=null)
					{
						if (SelectFile=='') SelectFile=ListObjArray[i].Obj.LableID;
						else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.LableID
					}
				}
			}
		}
	}
	if ((SelectFolder=='')&&(SelectFile=='')) DisabledContentMenuStr=',修改,删除,备份,';
	else
	{
		if ((SelectFile!='')&&(SelectFolder==''))
		{
			if (SelectFile.indexOf('***')!=-1) DisabledContentMenuStr=',修改,';
			else DisabledContentMenuStr='';
		}
		if ((SelectFolder!='')&&(SelectFile!='')) DisabledContentMenuStr=DisabledContentMenuStr+',修改,备份,';
		if ((SelectFolder!='')&&(SelectFile==''))
		{
			if (SelectFolder.indexOf('***')!=-1) DisabledContentMenuStr=',修改,备份,';
			else DisabledContentMenuStr=',备份,';
		}
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if ((CurrObj.LableID!=null))
		{
			ListObjArray[ListObjArray.length]=new FreeLableObj(CurrObj,j,false);
			j++;
		}
	}
}
function EditFreeLable()
{
	var SelectedLable='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if(SelectedLable!='')
	{
		if (SelectedLable.indexOf('***')==-1) location='FreeLable_Edit.asp?FreeLableID='+SelectedLable;
		else alert('一次只能够编辑一个自由标签');
	}
	else alert('请选择要编辑的自由标签');
}
function DelFreeLable()
{
	var SelectedLable='',SelectedType='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if ((SelectedLable!='')||(SelectedType!=''))
	OpenWindow('Frame.asp?PageTitle=删除自由标签&FileName=DelFreeLable.asp&DelLable='+SelectedLable,200,120,window);
	else alert('没有选择要删除的自由标签');
}
function BackUpLable()
{
	var SelectedLable='',SelectedType='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.TypeID!=null)
			{
				if (SelectedType=='') SelectedType=ListObjArray[i].Obj.TypeID;
				else  SelectedType=SelectedType+'***'+ListObjArray[i].Obj.TypeID;
			}
			if (ListObjArray[i].Obj.LableID!=null)
			{
				if (SelectedLable=='') SelectedLable=ListObjArray[i].Obj.LableID;
				else  SelectedLable=SelectedLable+'***'+ListObjArray[i].Obj.LableID;
			}
		}
	}
	if (SelectedLable!='') OpenWindow('Frame.asp?PageTitle=备份标签&BackUpLable='+SelectedLable+'&FileName=BackUpLable.asp',200,120,window);
	else alert('请选择要备份的标签')
}
</script>
<body topmargin="2" leftmargin="2" onClick="SelectFreeLable();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="新建自由标签" onClick="AddFreeLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="修改自由标签" onClick="EditFreeLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除自由标签" onClick="DelFreeLable();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
	 <td height="93" valign="top">
  		<table width="100%" border="0" cellpadding="0" cellspacing="0" dwcopytype="CopyTableRow">
		 <tr> 
       		<td height="26" align="center" class="ButtonListLeft" colspan="2">名称</td>
    		<td width="17%" align="center" class="ButtonList">建立日期</td>
    		<td width="10%" align="center" class="ButtonList">大小</td>
    		<td width="50%" align="center" class="ButtonList">描述</td>
  		</tr>
<%
Dim i
i=0
While not RsLableObj.eof
%>
        <tr style="background:white;cursor:default;"> 
		  <td width="3%" align="center"><img src="../../Images/FreeLableIcon.gif"></td>
          <td width="20%"><span id="Freelable<%=i%>" LableID="<%=RsLableObj("freelableid")%>"><%=RsLableObj("name")%></span></td>
          <td align="center"><%=RsLableObj("addtime")%></td>
          <td align="right"><%=len(Replace(Trim(RsLableObj("stylecontent")),"*|*","'"))%>字节</td>
          <td align="center"><%if len(trim(RsLableObj("description"))) > 30 then Response.write(Left(RsLableObj("description"),30)&"...") else Response.write(RsLableObj("description")) end if%></td>
        </tr>
        <%
	i = i + 1
	RsLableObj.MoveNext
Wend
RsLableObj.Close
%>
	</table>
</td>
</tr>
</table>
</body>
</html>
<%
Set RsLableObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
function SelectFreeLable()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
			{
				if (MaxIndex>0)
					for (i=MaxIndex-1;i<EndIndex;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
				else
				{
					ListObjArray[ElIndex].Obj.className='TempletSelectItem';
					ListObjArray[ElIndex].Selected=true;
				}
			}
			else
			{
				if (ObjInArray)
				{
					for (i=EndIndex;i<MaxIndex-1;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					if (ElIndex>=0)
					{
						ListObjArray[ElIndex].Obj.className='TempletSelectItem';
						ListObjArray[ElIndex].Selected=true;
					}
				}
			}
		}
	}
	else
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	}
}
</script>