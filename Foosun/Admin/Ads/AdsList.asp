<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P070200") then Call ReturnError1()
Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
Dim TempState,RsAdsObj,AdsSql,AdsFlag,TempAddTime,TempType,TempAdsState,Location,AdsIDTemp
AdsSql = "Select * from FS_Ads order by Location asc"
AdsFlag = "�������"
TempState = Request("State")
if TempState <> "" then
     if  Cstr(TempState)="InGear" then
		 AdsSql = "Select * from FS_Ads where State=1 order by Location asc"
		 AdsFlag = "�������"
	 elseif Cstr(TempState)="ClickMax" then
		 AdsSql = "Select * from FS_Ads order by ClickNum desc,Location asc"
		 AdsFlag = "������"
	 elseif Cstr(TempState)="ClickMin" then
		 AdsSql = "Select * from FS_Ads order by ShowNum desc,Location asc"
		 AdsFlag = "��ʾ���"
	 elseif Cstr(TempState)="Abate" then
		 AdsSql = "Select * from FS_Ads where State=0 order by Location asc"
		 AdsFlag = "ʧЧ���"
	 elseif Cstr(TempState)="ShowAds" then
		 AdsSql = "Select * from FS_Ads where Type=1 order by Location asc"
		 AdsFlag = "��ʾ���"
	 elseif Cstr(TempState)="Stop" then
		 AdsSql = "Select * from FS_Ads where State=2 order by Location asc"
		 AdsFlag = "��ͣ���"
	 elseif Cstr(TempState)="NewWindow" then
		 AdsSql = "Select * from FS_Ads where Type=2 order by Location asc"
		 AdsFlag = "�����´���"
	 elseif Cstr(TempState)="OpenWindow" then
		 AdsSql = "Select * from FS_Ads where Type=3 order by Location asc"
		 AdsFlag = "���´���"
	 elseif Cstr(TempState)="FilterAway" then
		 AdsSql = "Select * from FS_Ads where Type=4 order by Location asc"
		 AdsFlag = "������ʧ"
	 elseif Cstr(TempState)="DialogBox" then
		 AdsSql = "Select * from FS_Ads where Type=5 order by Location asc"
		 AdsFlag = "��ҳ�Ի���"
	 elseif Cstr(TempState)="ClarityBox" then
		 AdsSql = "Select * from FS_Ads where Type=6 order by Location asc"
		 AdsFlag = "͸���Ի���"
	 elseif Cstr(TempState)="DriftBox" then
		 AdsSql = "Select * from FS_Ads where Type=8 order by Location asc"
		 AdsFlag = "��������"
	 elseif Cstr(TempState)="LeftBottom" then
		 AdsSql = "Select * from FS_Ads where Type=9 order by Location asc"
		 AdsFlag = "���µ׶�"
	 elseif Cstr(TempState)="RightBottom" then
		 AdsSql = "Select * from FS_Ads where Type=7 order by Location asc"
		 AdsFlag = "���µ׶�"
	 elseif Cstr(TempState)="Couplet" then
		 AdsSql = "Select * from FS_Ads where Type=10 order by Location asc"
		 AdsFlag = "�������"
	 elseif Cstr(TempState)="Cycle" then
		 AdsSql = "Select * from FS_Ads where Type=11 or CycleTF=1 order by Location asc"
		 AdsFlag = "ѭ�����"
	 end if
end if
Set RsAdsObj = Server.CreateObject(G_FS_RS)
RsAdsObj.Open AdsSql,Conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����б�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectAds();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="��ӹ��" onClick="AddAds();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="�޸Ĺ��" onClick="EditAds();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�޸�</td>
		  <td  width=2 class="Gray">|</td>
          <td width=35 align="center" alt="ɾ�����" onClick="DelAds();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td  width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="��ͣ" onClick="StopAds();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��ͣ</td>
		  <td  width=2 class="Gray">|</td>
          <td width=35 align="center" alt="����" onClick="StartAds();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td  width=2 class="Gray">|</td>
          <td width=55 align="center" alt="���ô���" onClick="GetCode();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���ô���</td>
		  <td  width=2 class="Gray">|</td>
          <td width=55 align="center" alt="����ͳ��" onClick="ShowStat();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����ͳ��</td>
		  <td  width=2 class="Gray">|</td>
          <td width=55 align="center" alt="���ͳ��" onClick="ClickStat();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���ͳ��</td>
 		  <td  width=2 class="Gray">|</td>
          <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;</td>
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
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="10%" height="26" class="ButtonListLeft"> 
            <div align="center">���λ</div></td>
          <td width="23%" height="26" class="ButtonList"> 
            <div align="center">�������</div></td>
          <td width="22%" height="26" class="ButtonList"> 
            <div align="center">���ʱ��</div></td>
          <td width="16%" height="26" class="ButtonList"> 
            <div align="center">�������</div></td>
          <td width="16%" height="26" class="ButtonList"> 
            <div align="center">��ʾ����</div></td>
          <td width="13%" height="26" class="ButtonList"> 
            <div align="center">״̬</div></td>
        </tr>
        <%
if Not RsAdsObj.Eof then
	Dim page_size,page_no,page_total,record_all
	page_size=20
	page_no=request.querystring("page_no")
	if page_no <= 1 or page_no = "" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no = 1
	end if
	RsAdsObj.PageSize = page_size
	page_total = RsAdsObj.PageCount
	if (cint(page_no) > page_total) then page_no=page_total
	RsAdsObj.AbsolutePage=page_no
	record_all=RsAdsObj.RecordCount
	dim i
	for i=1 to RsAdsObj.PageSize
	if RsAdsObj.eof then exit for
	select  case RsAdsObj("Type")
	    case "1"  TempType = "��ʾ���"
	    case "2"  TempType = "�����´���"
	    case "3"  TempType = "���´���"
	    case "4"  TempType = "������ʧ"
	    case "5"  TempType = "��ҳ�Ի���"
	    case "6"  TempType = "͸���Ի���"
	    case "7"  TempType = "���µ׶�"
	    case "8"  TempType = "��������"
	    case "9"  TempType = "���µ׶�"
	    case "10"  TempType = "�������"
	    case "11"  TempType = "ѭ�����"
     end select
	 TempAddTime = year(RsAdsObj("AddTime"))&"-"&month(RsAdsObj("AddTime"))&"-"&day(RsAdsObj("AddTime"))
      select case RsAdsObj("State")
	       case "0" TempAdsState="ʧЧ"
		   case "1" TempAdsState="����"
		   case "2" TempAdsState="��ͣ"
		  end select
     AdsIDTemp = RsAdsObj("ID")
%>
        <tr height="20"> 
          <td><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Ads.gif" width="24" height="22"></td>
                <td><span class="TempletItem" AdsID="<% = RsAdsObj("ID") %>" State="<%=RsAdsObj("State")%>" Location="<%=RsAdsObj("Location")%>">��<%=RsAdsObj("Location")%>λ</span></td>
              </tr>
            </table></td>
          <td height="25"> <div align="center"><%=TempType%></div></td>
          <td><div align="center"><%=TempAddTime%></div></td>
          <td><div align="center"><%=RsAdsObj("ClickNum")%></div></td>
          <td><div align="center"><%=RsAdsObj("ShowNum")%></div></td>
          <td><div align="center"><%=TempAdsState%></div></td>
        </tr>
        <%
		RsAdsObj.MoveNext
	Next
end if
RsAdsObj.Close
Set RsAdsObj = Nothing
%>
        <%if page_total>1 then%>
        <tr> 
          <td colspan="6" height="18"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td  valign="middle" height="10"> 
                <td width="51%" height="25" valign="middle"> <div align="right"> 
                    <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
                    <% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
                    <%
				if Page_Total=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
				else
					if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
					else
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/endpage.gif border=0 alt=βҳ></img>&nbsp;"
					end if
				end if
				%>
                  </div></td>
                <td width="7%" valign="middle"><select onChange="ChangePage(this.value);" style="width:50;" name="select">
                    <% for i=1 to Page_Total %>
                    <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>"> 
                    <% = i %>
                    </option>
                    <% next %>
                  </select></td>
              </tr>
            </table></td>
        </tr>
        <% end if %>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script>
var TempStates = '<% = TempState %>';
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditAds();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelAds();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StopAds();','��ͣ','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.StartAds();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.GetCode();','���ô���','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
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
	var EventObjInArray=false,SelectAds='',DisabledContentMenuStr='';
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.AdsID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.AdsID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.AdsID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.AdsID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',�޸�,ɾ��,��ͣ,����,���ô���,';
	else
	{
		if (SelectAds.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,���ô���,'
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
		if (CurrObj.AdsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectAds()
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
function AddAds()
{
	location='AdsAdd.asp?Typess='+TempStates;
}
function EditAds()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
	{
		if (SelectedAds.indexOf('***')==-1) location='AdsModify.asp?Location='+SelectedAds;
		else alert('һ��ֻ�ܹ��޸�һ�����');
	}
	else alert('��ѡ��Ҫ�޸ĵĹ��');
}
function StopAds()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
		OpenWindow('Frame.asp?FileName=AdsTip.asp&PageTitle=��ͣ���&Types=Stop&Location='+SelectedAds,220,105,window);
	else alert('��ѡ����');
}
function StartAds()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
		OpenWindow('Frame.asp?FileName=AdsTip.asp&Types=Star&PageTitle=������&Location='+SelectedAds,220,105,window);
	else alert('��ѡ����');
}
function DelAds()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
		OpenWindow('Frame.asp?FileName=AdsTip.asp&Types=Dell&PageTitle=ɾ�����&Location='+SelectedAds,220,105,window);
	else alert('��ѡ����');
}
function GetCode()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
	{
		if (SelectedAds.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=Code.asp&PageTitle=���ô���&Location='+SelectedAds,360,160,window);
		else alert('��ѡ��һ�����');
	}
	else alert('��ѡ����');
}
function ShowStat()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
	{
		if (SelectedAds.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=VisitList.asp&PageTitle=����ͳ��&Types=Shows&Location='+SelectedAds,360,200,window);
		else alert('��ѡ��һ�����');
	}
	else alert('��ѡ����');
}
function ClickStat()
{
	var SelectedAds='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Location!=null)
			{
				if (SelectedAds=='') SelectedAds=ListObjArray[i].Obj.Location;
				else  SelectedAds=SelectedAds+'***'+ListObjArray[i].Obj.Location;
			}
		}
	}
	if (SelectedAds!='')
	{
		if (SelectedAds.indexOf('***')==-1) OpenWindow('Frame.asp?FileName=VisitList.asp&PageTitle=���ͳ��&Types=Clicks&Location='+SelectedAds,360,200,window);
		else alert('��ѡ��һ�����');
	}
	else alert('��ѡ����');
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>