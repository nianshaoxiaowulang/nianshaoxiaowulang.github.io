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
Dim Conn,DownLoadConfig,DBC,SQLStr,IPList,Lock,IPType
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040502") then Call ReturnError1()
SQLStr="Select * from FS_DownLoadConfig"
Set DownLoadConfig = Server.CreateObject(G_FS_RS)
DownLoadConfig.Open SQLStr,Conn,1,3
if Not DownLoadConfig.Eof then
	IPList = DownLoadConfig("IPList")
	Lock = DownLoadConfig("Lock")
	IPType = DownLoadConfig("IPType")
else
	IPList = ""
	Lock = ""
	IPType = ""
end if
if Request.Form("Operation") = "Modify" then
	On Error Resume Next
	IPList = Request.Form("IPList")
	Lock = Request.Form("Lock")
	IPType = Request.Form("IPType")
	DownLoadConfig("IPList")=Replace(Replace(IPList,"'",""),"""","")
	if Lock = "1" then
		DownLoadConfig("Lock") = 1
	else
		DownLoadConfig("Lock") = 0
	end if
	DownLoadConfig("IPType")=Replace(Replace(IPType,"'",""),"""","")
	DownLoadConfig.update
	if Err.Mumber = 0 then
		%>
		<script language="javascript">
		alert('�޸ĳɹ�');window.location='DownLoadParameter.asp';
		</script>
		<%
	else
		%>
		<script language="javascript">
		alert('�д���������ˢ�º�����');window.location='DownLoadParameter.asp';
		</script>
		<%
		Response.Redirect("DownLoadParameter.asp")  
	end if
	Response.End
end if 
%>
<html>
<title>���ز���������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2" scroll=yes  oncontextmenu="return false;">
<form name="Form" method=post action="" >
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="SetIPList();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp;<input type=hidden name=operation value=Modify></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1"  bordercolor="e6e6e6" bgcolor="#dddddd">
    <tr valign="middle"> 
      <td width="120" height="30" bgcolor="#FFFFFF"> 
        <div align="left"> �Ƿ�ӷ�����</div></td>
          
      <td width="469" height="30" bgcolor="#FFFFFF"> 
        <input <% if Lock=1 then Response.Write("Checked") %>  name="Lock" type="checkbox" id="Lock" value="1">
            <input name="IPList" type="hidden" id="IPList"></td>
        </tr>
        
    <tr valign="middle" > 
      <td width="120" rowspan="2" bgcolor="#FFFFFF"> 
        <div align="left"> 
              <p> 
                <input name="IPType" <% if IPType=1 then Response.Write("Checked") %> type="radio" value="1" checked>
                ����IP��</p>
              <p> 
                <input type="radio" <% if IPType=0 then Response.Write("Checked") %>  name="IPType" value="0">
                ���IP��</p>
            </div></td>
          
      <td height="30" bgcolor="#FFFFFF"> 
        <select name="IPSelectList" size="10" multiple id="IPSelectList" style="width:100%;">
              <%
		  Dim TempArray,i
		  if Not IsNull(IPList) then
			  TempArray = Split(IPList,"$")
			  for i=LBound(TempArray) to UBound(TempArray)
			  %>
			  <option value="<% = TempArray(i) %>"><% = TempArray(i) %></option>
			  <%
			  Next
		  end if
		  %>
           </select>
		  </td>
        </tr>
        <tr valign="middle" > 
          
      <td width="469" height="30" bgcolor="#FFFFFF"> 
        <input name="BeginIP" type="text" id="BeginIP">
            --- 
            <input name="EndIP" type="text" id="EndIP"> <input type="button" onClick="AddIPList();" name="Submit3" value=" �� �� "> 
            <input type="button" onClick="DelIPList();" name="Submit4" value=" ɾ �� ">
      </td>
        </tr>
    </table>
</form>
</body>
</html>
<%
DownLoadConfig.close
Set DownLoadConfig =nothing
Set Conn=nothing
%>
<script language="JavaScript">
function AddIPList()
{
	var BeginIPStr=document.Form.BeginIP.value,EndIPStr=document.Form.EndIP.value;
	if (CheckIP(BeginIPStr))
	{
		if (CheckIP(EndIPStr))
		{
			if (CheckBeginAndEndIP(BeginIPStr,EndIPStr))
			{
				var TempStr=BeginIPStr+'-'+EndIPStr;
				AddList(document.Form.IPSelectList,TempStr,TempStr);
				document.Form.BeginIP.value='';
				document.Form.EndIP.value='';
			}
		}
		else
		{
			alert('����IP��ַ����');
			document.Form.EndIP.focus();
			document.Form.EndIP.select();
		}
	}
	else
	{
		alert('��ʼIP��ַ����');
		document.Form.BeginIP.focus();
		document.Form.BeginIP.select();
	}
}
function DelIPList()
{
	DelList(document.Form.IPSelectList);
}
function SetIPList()
{
	var TempStr='',Obj=document.Form.IPSelectList;
	for(var i=0;i<Obj.length;i++)
	{
		if (TempStr=='') TempStr=Obj.options(i).value;
		else TempStr=TempStr+'$'+Obj.options(i).value;
	}
	document.Form.IPList.value=TempStr;
	document.Form.submit();
}
function CheckBeginAndEndIP(BeginIPStr,EndIPStr)
{
	return true;
}
function CheckIP(IPAddress)
{
	var TempArray=null,TempInt=0;
	TempArray=IPAddress.split('.');
	if (TempArray.length!=4) return false;
	for (var i=0;i<TempArray.length;i++)
	{
		if (TempArray[i]!='')
		{
			TempInt=parseInt(TempArray[i]);
			if ((TempInt<0)||(TempInt>255)) return false;
		}
		else return false;
	}
	return true;
}
function AddList(SelectObj,Lable,LableContent)
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
			Obj.options(i).selected=true;
			return true;
		}
	}
	return false;
}
function DelList(SelectObj)
{
	var OptionLength=SelectObj.length;
	for(var i=0;i<OptionLength;i++)
	{
		if (SelectObj.options(SelectObj.length-1).selected==true) SelectObj.options.remove(SelectObj.length-1);
		//OptionLength=SelectObj.length;
	}
}
</script>