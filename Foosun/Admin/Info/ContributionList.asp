<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
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
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================

Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010600") then Call ReturnError1()
Dim NewsSql,RsNewsObj,RsClassObj,ClassID,ClassCName,RsChildClassObj,AllowContributionTF,DisableStr
ClassID = Request("ClassID")
ClassID = Replace(Replace(Replace(Replace(Replace(ClassID,"'",""),"and",""),"select",""),"or",""),"union","")
if ClassID = "0" or ClassID = "" then
	NewsSql = "Select * from FS_Contribution order by AddTime desc"
	AllowContributionTF = True
Else	
	NewsSql = "Select * from FS_Contribution where ClassID='" & ClassID & "' order by AddTime desc"
	Set RsClassObj = Conn.Execute("Select Contribution from FS_NewsClass where ClassID='" & ClassID & "'")
	if Not RsClassObj.Eof then
		if RsClassObj("Contribution") = 1 then
			AllowContributionTF = True
		else
			AllowContributionTF = False
			DisableStr = "disabled"
		end if
	else
		AllowContributionTF = False
		DisableStr = "disabled"
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectContr();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="�½�" onClick="CreateNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>�½�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�޸�" onClick="EditNews();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>�޸�</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="ɾ��" onClick="DelNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="���" onClick="Audit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>���</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="37%" height="26" class="ButtonListLeft">
<div align="center">�������</div></td>
          <td width="18%" height="26" class="ButtonList">
<div align="center">���ʱ��</div></td>
          <td width="18%" height="26" class="ButtonList">
<div align="center">������Ŀ</div></td>
          <td width="16%" height="26" class="ButtonList">
<div align="center">����</div></td>
          <td width="11%" height="26" class="ButtonList">
<div align="center">��С</div></td>
        </tr>
<%
if AllowContributionTF = True then
	Set RsNewsObj = Conn.Execute(NewsSql)
	do while Not RsNewsObj.Eof
	ClassCName=conn.execute("select ClassCName from FS_NewsClass where ClassID='" & RsNewsObj("Classid") & "'")(0)
%>
        <tr> 
          <td height="20">
		  <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="../../Images/Info/WordNews.gif"></td>
                <td><span class="TempletItem" NewsID="<% = RsNewsObj("ContID") %>" align="center"> 
                  <% = GotTopic(RsNewsObj("Title"),30) %>
                  </span> </td>
              </tr>
          </table>
		  </td>
          <td height="20"><div align="center" class="TempletItem"><% = RsNewsObj("AddTime") %></div></td>
		  <td height="20"><div align="center" class="TempletItem"><% = ClassCName %></div></td>
          <td height="20"><div align="center" class="TempletItem"><% = RsNewsObj("Author") %></div></td>
          <td height="20"><div align="center" class="TempletItem"><% = Len(RsNewsObj("Content")) %>
              b</div></td>
        </tr>
        <%
		RsNewsObj.MoveNext
	loop
%>
<%
else
%>
  <tr> 
    <td colspan="5" height="26"><div align="center">����Ŀ������Ͷ�� </div>
      <div align="center"></div></td>
    </tr>
<%
end if
%>
      </table>
	</td>
  </tr>
</table>
</body>
</html>
<%
Set RsChildClassObj = Nothing
Set RsNewsObj = Nothing
Set RsClassObj = Nothing
Set Conn = Nothing
%>
<script language="javascript"> 
var ClassID = '<% = ClassID %>';
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditNews();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelNews();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit();','���','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContribution='',DisabledContentMenuStr='';
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
			if (SelectContribution=='') SelectContribution=ListObjArray[i].Obj.NewsID;
			else SelectContribution=SelectContribution+'***'+ListObjArray[i].Obj.NewsID;
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
					if (SelectContribution=='') SelectContribution=ListObjArray[i].Obj.NewsID;
					else SelectContribution=SelectContribution+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContribution=='') DisabledContentMenuStr=',�޸�,ɾ��,���,';
	else
	{
		if (SelectContribution.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,'
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.NewsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectContr()
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
function CreateNews()
{
	location='ContributionAdd.asp?ClassID='+ClassID;
}
function EditNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) location='ContributionModify.asp?ClassID='+ClassID+'&NewsID='+SelectedNews;
		else alert('һ��ֻ�ܹ��޸�һ������');
	}
	else alert('��ѡ��Ҫ�޸ĵ�����');
}
function DelNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=ContributionDell.asp&PageTitle=���ɾ��&NewsID='+SelectedNews,220,110,window);
	else alert('��ѡ��Ҫɾ����Ͷ��');
}
function Audit()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) location='ContributionCheck.asp?NewsID='+SelectedNews+'&ClassID='+ClassID;
		else alert('һ��ֻ�ܹ����һ������');
	}
	else alert('��ѡ��Ҫ��˵�Ͷ��');
}
function CutOperation()
{
	parent.MoveTF=true;
	if (NewsID!='')
	{  
	     parent.MoveOrCopySourceClass=BigClassID;
		 parent.MoveOrCopySourceNews=NewsID;
	}
}

function CopyOperation()
{
	parent.MoveTF=false;
	if (NewsID!='')
	{
	     parent.MoveOrCopySourceClass=BigClassID;
		 parent.MoveOrCopySourceNews=NewsID;
	}
}

function PasteOperation()
{
	var MoveOrCopyClassPara='MoveTF:'+parent.MoveTF+',SourceClass:'+parent.MoveOrCopySourceClass+',SourceNews:'+parent.MoveOrCopySourceNews+',ObjectClass:'+parent.MoveOrCopyObjectClass+',';
	OpenWindow('ContTip.asp?FileName=MoveOrCopyCont.asp&Titles=����ƶ�����&MoveOrCopyClassPara='+MoveOrCopyClassPara,310,95,window);
}
</script>
