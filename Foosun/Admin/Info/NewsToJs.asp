<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<!--#include file="../../../Inc/ThumbnailFunction.asp" -->
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
If Not JudgePopedomTF(Session("Name"),"P010508") then Call ReturnError()
Dim TempSysRootDir
If SysRootDir = "" then
	TempSysRootDir = ""
Else
	TempSysRootDir = "/" & SysRootDir
End if

Dim Types,NewsID,RsNewsObj
If Request("NewsID")<>"" and Request("Types")<>"" then
   NewsID = Cstr(Request("NewsID"))
   Types = Cstr(Request("Types"))
Else
	Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
	Response.End
End if
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加新闻到自由JS</title>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <form action="" name="ToPicJsForm" method="post" >
    <tr> 
      <td width="7%" height="5">&nbsp;</td>
      <td width="16%" height="5">&nbsp;</td>
      <td width="77%" height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>JS名称</td>
      <td><select name="JSName" id="JSName" style="width:90%" onChange=ChooseJsName(this.options[this.selectedIndex].value)>
          <option value="" <%If Request("JSEName")="" then Response.Write("selected")%>> 
          </option>
      <%
	    Dim PicJsObj
		Set PicJsObj = Conn.Execute("Select EName,CName,Manner from FS_FreeJS order by AddTime desc")
	    Do While Not PicJsObj.eof 
	  %>
          <option value="<%=PicJsObj("EName")&"***"&PicJsObj("Manner")%>" <%If Cstr(Request("JSEName")) = Cstr(PicJsObj("EName")) then Response.Write("selected")%>><%=PicJsObj("CName")%></option>
     <%
			PicJsObj.MoveNext
		Loop
	    PicJsObj.Close
		Set PicJsObj = Nothing
	  %>
        </select> <input name="Manner" type="hidden" id="Manner"> <input name="JSEName" type="hidden" id="JSEName"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>图片 地址</td>
      <td><input name="PicPath" type="text" id="PicPath" size="28" value="<%=Request("PicPath")%>"> 
        <input type="button" name="PicChooseButton" value="选择图片" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.ToPicJsForm.PicPath);"></td>
    </tr>
    <tr> 
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
      <td height="5">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"> 
          <input type="button" name="Submit2" value=" 确 定 " onClick="ChoosePicPath();">
          <input name="action" type="hidden" id="action" value="trues">
          <input type="button" name="Submit3" value=" 取 消 " onClick="window.close();">
        </div></td>
    </tr>
  </form>
</table>
</body>
</html>
<script>
function ChoosePicPath()
{
	var Value=parseInt(document.ToPicJsForm.Manner.value);
	if (Value>=6)
    {
	  if (document.ToPicJsForm.PicPath.value=='')
		 {
		  alert('图片地址不能为空');
		  return;
		  }
	 }
  document.ToPicJsForm.submit();
 }
 
function ChooseWordJsName(TempString)
{
   var TempArr=TempString.split("***");
   document.ToWordJsForm.Manner.value=TempArr[1];
   document.ToWordJsForm.JSEName.value=TempArr[0];
 }
 
function ChooseJsName(TempStr)
{
	var TempArray=TempStr.split("***");
	document.ToPicJsForm.Manner.value=TempArray[1];
	document.ToPicJsForm.JSEName.value=TempArray[0];
	var Value=parseInt(TempArray[1]);
	if (Value<6)
	{
		document.ToPicJsForm.PicPath.disabled=true;
		document.ToPicJsForm.PicChooseButton.disabled=true;
	}
	else
	{
		document.ToPicJsForm.PicPath.disabled=false;
		document.ToPicJsForm.PicChooseButton.disabled=false;
	}
}
</script>
<%
If Request.Form("action")="trues" then
  Dim JsFileObj,JsFileSql,TFFlagObj,NewsIDArray,Rt_i,RsNewsTFObj
  If Request.Form("JSEName")="" or isnull(Request.Form("JSEName")) then
	  Response.Write("<script>alert(""请选择自由JS"");</script>")
	  Response.End
  End if
  '======================================
  '如果系统设置了生成缩略图功能 则生成缩略图
  Dim JSPicWidth,JSPicHeight,FreeJSRs,OpenCreateThumbnail,CreateSmallPicOK
  Dim sRootDir,PicFileName
  CreateSmallPicOK=False
  OpenCreateThumbnail=Conn.Execute("Select ThumbnailComponent from FS_Config")(0)
  If Request.Form("PicPath")<>"" and OpenCreateThumbnail=1 then
 	PicFileName=mid(Request.Form("PicPath"),InStrRev(Request.Form("PicPath"),"/")+1)
	set FreeJsRs=Conn.execute("Select PicHeight,PicWidth,Manner From FS_FreeJS Where EName='"&Request.Form("JSEName")&"'")
	If Not FreeJSRs.EOF then
		If FreeJsRs("Manner")<> "12" and FreeJsRs("Manner")<> "16" then
			If SysRootDir<>"" then 
				sRootDir="/"&SysRootDir & left(Request.Form("PicPath"),instrrev(Request.Form("PicPath"),"/"))
			Else
				sRootDir=left(Request.Form("PicPath"),InStrRev(Request.Form("PicPath"),"/"))
			End IF	
			JSPicWidth=FreeJsRs("PicWidth")
			JSPicHeight=FreeJsRs("PicHeight")
			CreateSmallPicOK=CreateThumbnail(sRootDir&PicFileName,JSPicWidth,JSPicHeight,"0",sRootDir&"s_"&PicFileName)'由原图片生成指定宽度和高度的缩略图,如果成功返回True,失败返回False
		End If
	End If
 End If
 '=======================================
  NewsIDArray = Array("")
  NewsIDArray = Split(NewsID,"***")
  For Rt_i = 0 to UBound(NewsIDArray)
  Set RsNewsTFObj = Conn.Execute("Select FileName from FS_FreeJsFile where JSName='"&Request.Form("JSEName")&"' and FileName=(Select FileName from FS_News where NewsID='"&NewsIDArray(Rt_i)&"')")
	  If RsNewsTFObj.eof then
	  Set RsNewsObj = Conn.Execute("Select Title,FileName,ClassID,AddDate from FS_News where HeadNewsTF=0 and DelTF=0 and AuditTF=1 and NewsID='"&NewsIDArray(Rt_i)&"'")
		 If Not RsNewsObj.eof Then
			
			  Set JsFileObj = Server.Createobject(G_FS_RS)
			  JsFileSql="select * from FS_FreeJsFile where 1=0"
			  JsFileObj.open JsFileSql,Conn,3,3
			  JsFileObj.AddNew
			  JsFileObj("Title") = RsNewsObj("Title")
			  JsFileObj("JSName") = Request.Form("JSEName")
			  JsFileObj("FileName") = RsNewsObj("FileName")
			  If Request.Form("PicPath")<>"" then
				  If CreateSmallPicOK=True then 
				  	JsFileObj("PicPath") =left(Request.Form("PicPath"),InStrRev(Request.Form("PicPath"),"/"))&"s_"&PicFileName
			 	  Else
				  	JsFileObj("PicPath") =Request.Form("PicPath")
				  End If	
			  End if
			  JsFileObj("ClassID") = RsNewsObj("ClassID")
			  JsFileObj("NewsTime") = RsNewsObj("AddDate")
			  JsFileObj("ToJsTime") = Now()
			  JsFileObj.Update
			  JsFileObj.Close
			  Set JsFileObj = Nothing
		 End if
		 RsNewsObj.Close
		 Set RsNewsObj = Nothing
	 End If
	 RsNewsTFObj.Close
	 Set RsNewsTFObj = Nothing
	 Next
  
  '----------------生成JS文件-------------
  	Dim JSClassObj,ReturnValue
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = TempSysRootDir
  Select case Request.Form("Manner")
     case "1"   ReturnValue = JSClassObj.WCssA(Request.Form("JSEName"),True)
     case "2"   ReturnValue = JSClassObj.WCssB(Request.Form("JSEName"),True)
     case "3"   ReturnValue = JSClassObj.WCssC(Request.Form("JSEName"),True)
     case "4"   ReturnValue = JSClassObj.WCssD(Request.Form("JSEName"),True)
     case "5"   ReturnValue = JSClassObj.WCssE(Request.Form("JSEName"),True)
     case "6"   ReturnValue = JSClassObj.PCssA(Request.Form("JSEName"),True)
     case "7"   ReturnValue = JSClassObj.PCssB(Request.Form("JSEName"),True)
     case "8"   ReturnValue = JSClassObj.PCssC(Request.Form("JSEName"),True)
     case "9"   ReturnValue = JSClassObj.PCssD(Request.Form("JSEName"),True)
     case "10"  ReturnValue = JSClassObj.PCssE(Request.Form("JSEName"),True)
     case "11"  ReturnValue = JSClassObj.PCssF(Request.Form("JSEName"),True)
     case "12"  ReturnValue = JSClassObj.PCssG(Request.Form("JSEName"),True)
     case "13"  ReturnValue = JSClassObj.PCssH(Request.Form("JSEName"),True)
     case "14"  ReturnValue = JSClassObj.PCssI(Request.Form("JSEName"),True)
     case "15"  ReturnValue = JSClassObj.PCssJ(Request.Form("JSEName"),True)
     case "16"  ReturnValue = JSClassObj.PCssK(Request.Form("JSEName"),True)
     case "17"  ReturnValue = JSClassObj.PCssL(Request.Form("JSEName"),True)
   End Select
   Set JSClassObj = Nothing
  '----------------   Over   -------------
	if ReturnValue <> "" then
		Response.Write("<script>alert('" & ReturnValue & "');window.close();</script>")
	else
	  Response.Write("<script>window.close();</script>")
	end if
end if
%>