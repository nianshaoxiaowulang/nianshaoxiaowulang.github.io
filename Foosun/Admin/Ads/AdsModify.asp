<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="Cls_Ads.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P070202") then Call ReturnError1()
Conn.Execute("Update FS_Ads set State=0 where State<>0 and (ClickNum>=MaxClick or ShowNum>=MaxShow or ( EndTime<>null and EndTime<="&StrSqlDate&"))")
Dim AdsModObj,AMLocation,CycleLocationRs,CycleLocationSql
AMLocation = Clng(request("Location"))
Set AdsModObj = Conn.Execute("select * from FS_Ads where Location="&AMLocation&"")
if AdsModObj.eof then
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	Response.End
else
%>

<html>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ��</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="AdsForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="document.AdsForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
            <td width="35" align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input type="hidden" name="action" value="mod"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF"> 
      <td width="14%" align="left" valign="middle"> 
        <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �� λ</div></td>
      <td width="37%" align="left" valign="middle" bgcolor="#FFFFFF"> 
        <input name="Location" type="text" id="Location2" style="width:90%" value="<%=AdsModObj("Location")%>" disabled> 
        <input name="Location" type="hidden" id="Location" value="<%=AdsModObj("Location")%>"> 
      </td>
      <td width="11%" valign="middle" bgcolor="#FFFFFF">�������</td>
      <td width="38%" valign="middle" bgcolor="#FFFFFF"> 
        <select name="Type" style="width:90%"  onchange="ChooseType(this);">
          <option value="1" <% if AdsModObj("Type")="1" then response.write("selected") end if%>>��ͨ��ʾ���</option>
          <option value="2" <% if AdsModObj("Type")="2" then response.write("selected") end if%>>�����´���</option>
          <option value="3" <% if AdsModObj("Type")="3" then response.write("selected") end if%>>���´���</option>
          <option value="4" <% if AdsModObj("Type")="4" then response.write("selected") end if%>>������ʧ</option>
          <option value="5" <% if AdsModObj("Type")="5" then response.write("selected") end if%>>��ҳ�Ի���</option>
          <option value="6" <% if AdsModObj("Type")="6" then response.write("selected") end if%>>͸���Ի���</option>
          <option value="8" <% if AdsModObj("Type")="8" then response.write("selected") end if%>>��������</option>
          <option value="9" <% if AdsModObj("Type")="9" then response.write("selected") end if%>>���µ׶�</option>
          <option value="7" <% if AdsModObj("Type")="7" then response.write("selected") end if%>>���µ׶�</option>
          <option value="10" <% if AdsModObj("Type")="10" then response.write("selected") end if%>>�������</option>
          <option value="11" <% if AdsModObj("Type")="11" then response.write("selected") end if%>>ѭ�����</option>
        </select> <font color="#0000FF">&nbsp; </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle"> 
        <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ѭ����� 
        </div></td>
      <td valign="middle"> 
        <input name="CycleTF" type="checkbox" id="CycleTF2" value="1"<% if AdsModObj("CycleTF")="1" then response.write("checked") end if%>  onclick="ChooseCycleDis();">
        ѭ�����λ 
        <select name="CycleLocation" style="width:51%">
          <option value="0"></option>
          <%
			set CycleLocationRs = server.createobject(G_FS_RS)
			CycleLocationSql = "select * from FS_Ads where Type=11"
			CycleLocationRs.open CycleLocationSql,conn,1,1
			while not CycleLocationRs.eof
		%>
          <option value="<%=CycleLocationRs("Location")%>" <%if clng(AdsModObj("CycleLocation"))=clng(CycleLocationRs("Location")) then response.write("selected") end if%>><%=CycleLocationRs("Location")%></option>
          <%
			CycleLocationRs.movenext
			wend
			CycleLocationRs.close
			set CycleLocationRs=nothing
		 %>
        </select> <%if AdsModObj("CycleLocation")<>"0" then %> <input name="TempLocation" type="hidden" id="TempLocation3" value="<%=AdsModObj("CycleLocation")%>"> 
        <%else%> <input name="TempLocation" type="hidden" id="TempLocation3" value="0"> 
        <%end if%> </td>
      <td valign="middle">ѭ������</td>
      <td valign="middle"> 
        <select name="CycleDirection" id="select"  style="width:38%">
          <option value="up" <%if AdsModObj("CycleDirection")="up" then response.write("selected") end if%>>����</option>
          <option value="down" <%if AdsModObj("CycleDirection")="down" then response.write("selected") end if%>>����</option>
          <option value="left" <%if AdsModObj("CycleDirection")="left" then response.write("selected") end if%>>����</option>
          <option value="right" <%if AdsModObj("CycleDirection")="right" then response.write("selected") end if%>>����</option>
        </select> &nbsp;ѭ���ٶ� 
        <input name="CycleSpeed" type="text" value="<%=AdsModObj("CycleSpeed")%>" style="width:25%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle"> 
        <div align="left"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ��ַ</div></td>
      <td valign="middle"> 
        <input name="LeftPicPath" type="text" value="<%=AdsModObj("LeftPicPath")%>" size="17"> 
        <input type="button" name="Submit" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.AdsForm.LeftPicPath);"> 
      </td>
      <td valign="middle">��ͼ��ַ</td>
      <td valign="middle"> 
        <input name="RightPicPath" type="text" size="18" disabled value="<%=AdsModObj("RightPicPath")%>" title="����������Ϊ������棬��ѡ�����������Ͳ���ѡ��"> 
        <input type="button" name="PPPCCC" value="ѡ��ͼƬ" disabled onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.AdsForm.RightPicPath);"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle"> 
        <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ�߶�</div></td>
      <td valign="middle"> 
        <input name="PicHeight" type="text" id="PicHeight2" value="<%=AdsModObj("PicHeight")%>" style="width:90%"> 
        &nbsp; </td>
      <td valign="middle">ͼƬ���</td>
      <td valign="middle"> 
        <input name="PicWidth" type="text" id="PicWidth" value="<%=AdsModObj("PicWidth")%>" style="width:90%"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle"> 
        <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ӵ�ַ</div></td>
      <td valign="middle"> 
        <input name="Url" type="text" id="Url" value="<%=AdsModObj("Url")%>" style="width:90%"> 
      </td>
      <td valign="middle">˵������</td>
      <td valign="middle"> 
        <input name="Explain" type="text" id="Explain" value="<%=AdsModObj("Explain")%>" style="width:90%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ѭ������</td>
      <td valign="middle"> 
        <input id="Cycle1" name="Cycle" type="radio" value="0" <%if AdsModObj("Cycle")="0" then response.write("checked") end if%>  onclick="ChooseCycle();">
        ������ѭ��&nbsp;&nbsp; <input type="radio" name="Cycle" value="1" <%if AdsModObj("Cycle")="1" then response.write("checked") end if%> onclick="ChooseCycle();">
        ������ѭ��</td>
      <td valign="middle">��ֹ����</td>
      <td valign="middle"> 
        <input name="EndTime" type="text" disabled readonly size="18" value="<%if AdsModObj("Cycle")="0" then response.Write("") else Response.Write(AdsModObj("EndTime")) end if%>"> 
        <input type="button" name="EEETTT" disabled value="ѡ������"  onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.AdsForm.EndTime);document.AdsForm.EndTime.focus();"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle"> 
        <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������</div></td>
      <td valign="middle"> 
        <input name="MaxClick" type="text" id="MaxClick2" value="<% if AdsModObj("MaxClick")="2147483647" then response.write("") else response.write(AdsModObj("MaxClick")) end if%>" style="width:90%"> 
        &nbsp;</td>
      <td valign="middle">��ʾ����</td>
      <td valign="middle"> 
        <input name="MaxShow" type="text" id="MaxShow" disabled style="width:90%" value="<% if AdsModObj("MaxShow")="2147483647" then response.write("") else response.write(AdsModObj("MaxShow")) end if%>" title="���ù��������ʾ����,��潫����ʾ�����ﵽ��������ʧЧ����������ô�����ÿ�"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="left" valign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��汸ע</td>
      <td colspan="4" valign="middle"> 
        <textarea name="Remark" rows="6" id="Remark2" style="width:96%"><%=AdsModObj("Remark")%></textarea> 
      </td>
    </tr>
</table>
</form>
</body>
</html>
<script>
function ChooseType(Obj)
{
  if (Obj.value=='10')
   {
    document.AdsForm.RightPicPath.disabled=false;
	document.AdsForm.PPPCCC.disabled=false;
	}
   else
    {
    document.AdsForm.RightPicPath.disabled=true;
	document.AdsForm.PPPCCC.disabled=true;
	 }
   if (Obj.value=='11') 
      {
	   document.AdsForm.CycleDirection.disabled=false;
	   document.AdsForm.CycleSpeed.disabled=false;
	   document.AdsForm.CycleTF.disabled=true;
	   document.AdsForm.CycleLocation.disabled=true;
	  }
   else 
     {
	  document.AdsForm.CycleTF.checked=false;
	  document.AdsForm.CycleTF.disabled=false;
	  document.AdsForm.CycleLocation.disabled=true;
	   document.AdsForm.CycleDirection.disabled=true;
	   document.AdsForm.CycleSpeed.disabled=true;
	  }
 }
 
function  ChooseCycle()
{
 if (document.AdsForm.Cycle1.checked==false)
    {
	document.AdsForm.MaxClick.disabled=false;
	document.AdsForm.EndTime.disabled=false;
	document.AdsForm.EEETTT.disabled=false;
	document.AdsForm.MaxShow.disabled=false;
	 }
  else
    {
	document.AdsForm.MaxClick.disabled=true;
	document.AdsForm.EndTime.disabled=true;
	document.AdsForm.EEETTT.disabled=true;
	document.AdsForm.MaxShow.disabled=true;
	document.AdsForm.EndTime.value='';
	 }
 }
 function ChooseCycleDis()
 {
  if (document.AdsForm.CycleTF.checked==true)
      {
	   document.AdsForm.CycleLocation.disabled=false;
	   document.AdsForm.CycleDirection.disabled=false;
	   document.AdsForm.CycleSpeed.disabled=false;
	   }
   else
      {
	   document.AdsForm.CycleLocation.disabled=true;
	   document.AdsForm.CycleDirection.disabled=true;
	   document.AdsForm.CycleSpeed.disabled=true;
	   }
  }
  
function TempFun()
{
  if (document.AdsForm.Type.value=='10')
   {
    document.AdsForm.RightPicPath.disabled=false;
	document.AdsForm.PPPCCC.disabled=false;
	}
   else
    {
    document.AdsForm.RightPicPath.disabled=true;
	document.AdsForm.PPPCCC.disabled=true;
	 }
   if (document.AdsForm.Type.value=='11') 
      {
	   document.AdsForm.CycleDirection.disabled=false;
	   document.AdsForm.CycleSpeed.disabled=false;
	   document.AdsForm.CycleTF.disabled=true;
	   document.AdsForm.CycleLocation.disabled=true;
	  }
   else 
     {
	  document.AdsForm.CycleTF.checked=false;
	  document.AdsForm.CycleTF.disabled=false;
	  document.AdsForm.CycleLocation.disabled=true;
	   document.AdsForm.CycleDirection.disabled=true;
	   document.AdsForm.CycleSpeed.disabled=true;
	  }
 }
 ChooseCycle();
 ChooseCycleDis();
 TempFun();
</script>
<%
  end if
  
  if request.form("action")="mod" then
  dim RsAdsObj,AdsChooseObj,ACycleSpeed,AdsSql,ALocation,ACycleDirection,ACycleTF,ACycleLocation,AType,ALeftPicPath,APicWidth,ARightPicPath,APicHeight,AUrl,AExplain,ACycle,AMaxShow,AMaxClick,AEndTime,ARemark
	   if request.form("Location")="" then
			response.write("<script>alert(""�������ݴ���"");location=""javascript:window.close()"";</script>")
       		response.end
		else
	      ALocation = Clng(request.form("Location"))
		end if
	   if isnumeric(request.form("Type"))=false then
			response.write("<script>alert(""������ʹ���"");location=""javascript:history.back(-1)"";</script>")
       		response.end
	    else
			AType = request.form("Type")
		end if
		if request.form("LeftPicPath")<>"" then
			ALeftPicPath = replace(replace(request.form("LeftPicPath"),"'",""),"""","")
		 else
		    response.write("<script>alert(""��ѡ����ͼƬ"");location=""javascript:history.back(-1)"";</script>")
		    response.end
		end if
		if AType="10" then
			if request.form("RightPicPath")="" then
				response.write("<script>alert(""��ѡ����������ͼƬ"");location=""javascript:history.back(-1)"";</script>")
				response.end
			end if
		end if
		if isnumeric(request.form("PicWidth"))=false or isnumeric(request.form("PicHeight"))=false then
           response.write("<script>alert(""���ͼƬ������Ϊ������"");location=""javascript:history.back(-1)"";</script>")
	       response.end
	    else
           APicWidth = request.form("PicWidth")
		   APicHeight = request.form("PicHeight")
		end if
		if request.form("Url")<>"" then
		   AUrl = replace(replace(request.form("Url"),"'",""),"""","")
		else
		   response.write("<script>alert(""�����������ӵ�ַ"");location=""javascript:history.back(-1)"";</script>")
		   response.end
		end if
		if request.form("Explain")<>"" then
		   AExplain = replace(replace(request.form("Explain"),"'",""),"""","")
		else
		   AExplain = ""
		end if
		if request.form("Cycle")="0" then
		   ACycle="0"
		   AMaxShow = "2147483647"
		   AMaxClick = "2147483647"
		   AEndTime = ""
		else
		   ACycle="1"
		   if request.form("MaxShow")<>"" and isnumeric(request.form("MaxShow"))=false then
		      response.write("<script>alert(""��������ʾ������Ϊ������"");location=""javascript:history.back(-1)"";</script>")
		      response.end
		   else
			   AMaxShow = request.form("MaxShow")
			end if
		   if request.form("MaxClick")<>"" and isnumeric(request.form("MaxClick"))=false then
		      response.write("<script>alert(""��������ʾ������Ϊ������"");location=""javascript:history.back(-1)"";</script>")
		      response.end
		   else
			   AMaxClick = request.form("MaxClick")
			end if
			if request.form("EndTime")="" or isnull(request.form("EndTime")) then
			   AEndTime=""
			else
			   AEndTime=formatdatetime(request.form("EndTime"))
			end if
		end if
		if request.form("Remark")<>"" then
		   ARemark = replace(replace(request.form("Remark"),"'",""),"""","")
		 else
		   ARemark = ""
		 end if
		 if request.form("CycleTF")="1" and AType <> "11"  then
		    if request.form("CycleLocation")="0" or request.form("CycleLocation")=""  then
			  response.write("<script>alert(""��ѡ��ѭ�����λ"");location=""javascript:history.back(-1)"";</script>")
			  response.end
			end if
		 end if
		  if AType="11" or request.form("CycleTF")="1" then
		     if isnumeric(request.form("CycleSpeed"))=false then
				  response.write("<script>alert(""���ѭ���ٶȱ���Ϊ������"");location=""javascript:history.back(-1)"";</script>")
				  response.end
			 else
				 ACycleSpeed=request.form("CycleSpeed")
			 end if
		  else
		  	 ACycleSpeed="2"
		  end if
	  set RsAdsObj=server.createobject(G_FS_RS)
	  AdsSql="select * from FS_Ads where Location="&ALocation&""
	  RsAdsObj.open AdsSql,Conn,1,3
	  RsAdsObj("Type")=cint(AType)
	  RsAdsObj("LeftPicPath")=cstr(ALeftPicPath)
	  if AType="10" then
		  RsAdsObj("RightPicPath")=Replace(Replace(request.form("RightPicPath"),"'",""),"""","")
	  End If
	  RsAdsObj("PicWidth")= cint(APicWidth)
	  RsAdsObj("PicHeight")= cint(APicHeight)
	  RsAdsObj("Url")= cstr(AUrl)
	  RsAdsObj("Explain")= cstr(AExplain)
	  RsAdsObj("CycleSpeed")= cint(ACycleSpeed)
	  if AMaxShow<>"" then
		  RsAdsObj("MaxShow")= AMaxShow
	  else
		  RsAdsObj("MaxShow")= "2147483647"
	  end if
	  if AMaxClick<>"" then
		  RsAdsObj("MaxClick")= AMaxClick
	   else
		  RsAdsObj("MaxClick")= "2147483647"
	   end if
	  if AEndTime<>"" then
		  RsAdsObj("EndTime")= AEndTime
	  end if
	  RsAdsObj("Remark")= ARemark
	  if  RsAdsObj("State")<>"2" then
		  RsAdsObj("State") = "1"
	  end if
	  RsAdsObj("LastTime")= now()
	  RsAdsObj("Cycle")= cint(ACycle)
	  if request.form("CycleDirection")<>"" and isnull(request.form("CycleDirection"))=false then
		  RsAdsObj("CycleDirection") = Cstr(request.form("CycleDirection"))
	  else
		  RsAdsObj("CycleDirection") = "up"
	  end if
	  if AType="11" then
		  RsAdsObj("CycleTF")="1"
		  ACycleTF = "1"
	   else
		  if request.form("CycleTF")="1" then
			  RsAdsObj("CycleTF")="1"
			  ACycleTF = "1"
		  else
			  RsAdsObj("CycleTF")="0"
			  ACycleTF = "0"
		  end if
	  end if
	  if ACycleTF="1" and AType<>"11" then
		  RsAdsObj("CycleLocation")=request.form("CycleLocation")
	   else
		  RsAdsObj("CycleLocation")="0"
	  end if
	  RsAdsObj.update
	  RsAdsObj.close
	  set RsAdsObj = nothing
		  select case AType
		    case "1" call ShowAds(ALocation)
			case "2" call NewWindow(ALocation)
			case "3" call OpenWindow(ALocation)
			case "4" call FilterAway(ALocation)
			case "5" call DialogBox(ALocation)
			case "6" call ClarityBox(ALocation)
			case "7" call RightBottom(ALocation)
			case "8" call DriftBox(ALocation)
			case "9" call LeftBottom(ALocation)
			case "10" call Couplet(ALocation)
		  end select
		    if ACycleTF<>"1" and request.form("TempLocation")="0" then

			else
				call Cycle(ALocation,request.form("TempLocation"))
			end if
			Response.Redirect("AdsList.asp")
			Response.End
	end if
set Conn=nothing
%>
