<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System v3.1 
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071,66252421
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060200") then Call ReturnError1()
Dim TempSysRootDir
if SysRootDir = "" then
	TempSysRootDir = ""
else
	TempSysRootDir = "/" & SysRootDir
end if

dim JSID,JSObj,TempManner,TempDateStr
if request("JSID")<>"" then
	JSID = Clng(request("JSID"))
Set JSObj = Conn.Execute("select * from FS_FreeJS where ID = "&JSID&"")
if JSObj.eof then
	 Response.Write("<script>alert(""参数传递错误"");window.close();</script>")
	 response.end
end if
TempManner = JSObj("Manner")
TempDateStr = JSObj("ShowTimeTF")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自由JS修改</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body leftmargin="2" topmargin="2" >
<form action="" method="post" name="JSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.JSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="action" type="hidden" id="action3" value="mod"> 
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E7E7E7">
    <tr bgcolor="#FFFFFF" Title="JS的中文名称，便于后台查阅和管理，请不要超过25个字符！"> 
      <td width="10%"> 
        <div align="center">名&nbsp;&nbsp;&nbsp;&nbsp;称</div></td>
      <td colspan="3"> 
        <input name="CName" type="text" id="CName" value="<%=JSObj("CName")%>" style="width:100%"> 
        <div align="center"></div></td>
      <td rowspan="12" align="center" id="PreviewArea"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">英文名称</div></td>
      <td colspan="3"> 
        <input name="EName" type="text" id="EName" value="<%=JSObj("EName")%>" disabled style="width:100%" Title="JS的英文名称，用于前台调用，请不要超过50个字符且不能与已经存在的JS重名！" > 
        <div align="center"></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">类&nbsp;&nbsp;&nbsp;&nbsp;型</div></td>
      <td width="20%"> 
        <input id="TypeW" name="Type" type="radio" value="0" <%if JSObj("Type")="0" then response.write("checked") end if%> onclick="TypeChoose();ChoosePic();" Title="JS类型（文字）选择！" >
        文字 
        <input id="TypeP" type="radio" name="Type" value="1" <%if JSObj("Type")="1" then response.write("checked") end if%> onclick="TypeChoose();ChoosePic();" Title="JS类型（图片）选择！" >
        图片</td>
      <td width="10%"> 
        <div align="center">新闻条数</div></td>
      <td width="20%"> 
        <input name="NewsNum" type="text" id="NewsNum2" value="<%=JSObj("NewsNum")%>" Title="此JS允许调用的新闻条数！"   style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">文字样式</div></td>
      <td> 
        <select name="Manner" id="MannerW" style="width:100% " Title="文字JS样式选择，上面有此样式的预览！" onChange="ChoosePic();">
          <option value="1" <%if JSObj("Manner")="1" then response.write("selected") end if%>>样式A</option>
          <option value="2" <%if JSObj("Manner")="2" then response.write("selected") end if%>>样式B</option>
          <option value="3" <%if JSObj("Manner")="3" then response.write("selected") end if%>>样式C</option>
          <option value="4" <%if JSObj("Manner")="4" then response.write("selected") end if%>>样式D</option>
          <option value="5" <%if JSObj("Manner")="5" then response.write("selected") end if%>>样式E</option>
        </select> </td>
      <td> 
        <div align="center">并排条数</div></td>
      <td> 
        <input name="RowNum" type="text" id="RowNum3" Title="此项设置JS在每行内显示的新闻条数，请务必不要置为‘0’！"  value="<%=JSObj("RowNum")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">图片样式</div></td>
      <td> 
        <select name="MannerP" id="MannerP" style="width:100% " disabled Title="图片JS样式选择，上面有此样式的预览！" onChange="ChoosePic();">
          <option value="6" <%if JSObj("Manner")="6" then response.write("selected") end if%>>样式A</option>
          <option value="7" <%if JSObj("Manner")="7" then response.write("selected") end if%>>样式B</option>
          <option value="8" <%if JSObj("Manner")="8" then response.write("selected") end if%>>样式C</option>
          <option value="9" <%if JSObj("Manner")="9" then response.write("selected") end if%>>样式D</option>
          <option value="10" <%if JSObj("Manner")="10" then response.write("selected") end if%>>样式E</option>
          <option value="11" <%if JSObj("Manner")="11" then response.write("selected") end if%>>样式F</option>
          <option value="12" <%if JSObj("Manner")="12" then response.write("selected") end if%>>样式G</option>
          <option value="13" <%if JSObj("Manner")="13" then response.write("selected") end if%>>样式H</option>
          <option value="14" <%if JSObj("Manner")="14" then response.write("selected") end if%>>样式I</option>
          <option value="15" <%if JSObj("Manner")="15" then response.write("selected") end if%>>样式J</option>
          <option value="16" <%if JSObj("Manner")="16" then response.write("selected") end if%>>样式K</option>
        </select></td>
      <td> 
        <div align="center">新闻行距</div></td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace3" value="<%=JSObj("RowSpace")%>"  style="width:100%;" Title="此项设置上下两条新闻之间的行距，请注意输入数值！" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">标题样式</div></td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" Title="新闻标题的CSS样式表。请直接输入样式名称。如果不选用此项设置，请置空！"  value="<%=JSObj("TitleCSS")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">新开窗口</div></td>
      <td> 
        <select name="OpenMode" id="select5" style="width:100%">
          <option value="1" <%If JSObj("OpenMode")=1 then Response.Write("selected")%>>是</option>
          <option value="0" <%If JSObj("OpenMode")=0 then Response.Write("selected")%>>否</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">标题字数</div></td>
      <td> 
        <input name="NewsTitleNum" type="text" id="NewsTitleNum2" value="<%=JSObj("NewsTitleNum")%>" Title="每条新闻的标题显示字数！"   style="width:100%;"></td>
      <td> 
        <div align="center">新闻日期</div></td>
      <td> 
        <select name="ShowTimeTF" id="select6" style="width:100%" onChange="ChooseDate(this.value);" Title="此项设置在新闻标题后面是否显示本条新闻的更新时间！" >
          <option value="1" <%If JSObj("ShowTimeTF")=1 then Response.Write("selected")%>>调用</option>
          <option value="0" <%If JSObj("ShowTimeTF")=0 then Response.Write("selected")%>>不调用</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">内容样式</div></td>
      <td> 
        <input name="ContentCSS" type="text" id="ContentCSS" Title="新闻内容的CSS样式表。请直接输入样式名称。如果不选用此项设置，请置空！"  value="<%=JSObj("ContentCSS")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">日期样式</div></td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" value="<%=JSObj("DateCSS")%>"  style="width:100%;" Title="日期字体的CSS样式。直接输入样式名称即可！" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">内容字数</div></td>
      <td> 
        <input name="ContentNum" type="text" id="ContentNum2" Title="为需要显示新闻内容的样式设置每条新闻的内容显示字数！"  value="<%=JSObj("ContentNum")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">背景样式</div></td>
      <td> 
        <input name="BackCSS" type="text" id="BackCSS2" value="<%=JSObj("BackCSS")%>"  style="width:100%;" Title="整体JS的背景样式（表格样式），请直接输入样式名称即可。如果不选用此项设置，请置空！" ></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">更多链接</div></td>
      <td> 
        <select name="MoreContent" id="select" style="width:100% " Title="此项为有新闻内容的样式在其右下角加一链接到该新闻页的链接，如果不显示此链接，请选择“不显示”！" >
          <option value="1" <%If JSObj("MoreContent")=1 then Response.Write("selected") %>>显示</option>
          <option value="0" <%If JSObj("MoreContent")=0 then Response.Write("selected") %>>不显示</option>
        </select></td>
      <td> 
        <div align="center">日期样式</div></td>
      <td> 
        <select name="DateType" id="select7" style="width:100%" Title="日期调用样式,默认为X月X日！" >
          <option value="1" <%if JSObj("DateType") = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
          <option value="2" <%if JSObj("DateType") = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
          <option value="3" <%if JSObj("DateType") = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
          <option value="4" <%if JSObj("DateType") = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
          <option value="5" <%if JSObj("DateType") = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
          <option value="6" <%if JSObj("DateType") = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
          <option value="7" <%if JSObj("DateType") = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
          <option value="8" <%if JSObj("DateType") = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
          <option value="9" <%if JSObj("DateType") = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
          <option value="10" <%if JSObj("DateType") = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
          <option value="11" <%if JSObj("DateType") = "11" then Response.Write("selected") end if%>><%=Month(Now)&"月"&Day(Now)&"日"%></option>
          <option value="12" <%if JSObj("DateType") = "12" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"时"%></option>
          <option value="13" <%if JSObj("DateType") = "13" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"点"%></option>
          <option value="14" <%if JSObj("DateType") = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"时"&Minute(Now)&"分"%></option>
          <option value="15" <%if JSObj("DateType") = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
          <option value="16" <%if JSObj("DateType") = "16" then Response.Write("selected") end if%>><%=Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日"%></option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">链接字样</div></td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" Title="为需要显示新闻链接的样式设置链接字样，可以是图片地址，如果是图片地址，请用<br>‘＜img src=../img/1.gif border=0＞’样式，其中‘src=’后为图片路径，‘border=0’为图片无边框！"  value="<%=JSObj("LinkWord")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">链接样式</div></td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" Title="为链接字样选择CSS样式，直接输入CSS样式名称即可！"  value="<%=JSObj("LinkCSS")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">图片宽度</div></td>
      <td> 
        <input name="PicWidth" type="text" disabled id="PicWidth3" Title="此项是为图片类型的JS设置图片的宽度参数！"  value="<%=JSObj("PicWidth")%>"  style="width:100%;"></td>
      <td> 
        <div align="center">图片高度</div></td>
      <td> 
        <input name="PicHeight" type="text" disabled id="PicHeight3" title="此项是为图片类型的JS设置图片的高度参数！"  value="<%=JSObj("PicHeight")%>"  style="width:100%;"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">导航图片</div></td>
      <td colspan="4"> 
        <input name="NaviPic" type="text" id="NaviPic" Title="新闻标题前面的导航图标，可以是“·”等字符，也可以是图片地址，如果是图片地址，请用<br>‘＜img src=../img/1.gif border=0＞’样式，其中‘src=’后为图片路径，‘border=0’为图片无边框！"  value="<%=JSObj("NaviPic")%>" style="width:100%"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">行间图片</div></td>
      <td colspan="4"> 
        <input name="RowBettween" type="text" id="RowBettween" style="width:80%;" size="26" value="<%=JSObj("RowBettween")%>" Title="此项设置上下两条新闻之间的间隔图片，请点击“选择图片”按钮进行设置，亦可为空！" > 
        <input id="RowBettweenButton" type="button" name="Submit34" value="选择图片" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.RowBettween);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">图片地址</div></td>
      <td colspan="4"> 
        <input name="PicPath" type="text" id="PicPath" value="<%=JSObj("PicPath")%>" style="width:80%;" size="26" disabled Title="为仅需一张图片的样式设置图片，请点击‘选择图片’按钮选择图片！" > 
        <input id="PicChooseButton" type="button" name="Submit34" value="选择图片" disabled onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.JSForm.PicPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="center">备&nbsp;&nbsp;&nbsp;&nbsp;注</div></td>
      <td colspan="4"> 
        <textarea name="Info" rows="6" id="Info" style="width:100%" Title="备注，用于代码调用时方便查看属性！" ><%=JSObj("Info")%></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<script> 
var TempDateScr = '<% = TempDateStr%>';
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	DocumentReadyTF=true;
	TypeChoose();
	ChoosePic();
	ChooseDate(TempDateScr);
}
function TypeChoose()
{
	if (document.JSForm.TypeW.checked==true)
	{ 
		document.JSForm.MannerW.disabled=false;
		document.JSForm.MannerP.disabled=true;
		document.JSForm.PicPath.disabled=true;
		document.JSForm.PicChooseButton.disabled=true;
		document.JSForm.PicWidth.disabled=true;
		document.JSForm.PicHeight.disabled=true;
	}
	else
	{
		document.JSForm.MannerW.disabled=true;
		document.JSForm.MannerP.disabled=false;
		document.JSForm.PicPath.disabled=false;
		document.JSForm.PicChooseButton.disabled=false;
		document.JSForm.PicWidth.disabled=false;
		document.JSForm.PicHeight.disabled=false;
	}
}
  
function ShowTitle(TempStr)
{
	document.all.TempTip.innerHTML='<font color=red>提示：</font><br><br>&nbsp;&nbsp;&nbsp;&nbsp;<font color=blue>'+TempStr+'</font>';
}
   
function ChooseDate(DateStr)
{ 
	if (DateStr==1)
	{
		document.JSForm.DateType.disabled=false;
		document.JSForm.DateCSS.disabled=false;
	}
	else
	{
		document.JSForm.DateType.disabled=true;
		document.JSForm.DateCSS.disabled=true;
	}
}
 
function ChoosePic()
{
	if (document.JSForm.MannerW.disabled==false) 
		 document.all.PreviewArea.innerHTML='<img src="Img/Css'+document.JSForm.MannerW.value+'.gif" border="0">';
	else 
		 document.all.PreviewArea.innerHTML='<img src="Img/Css'+document.JSForm.MannerP.value+'.gif" border="0">';
}
</script>
<%
  if Request.Form("action")="mod" then
     dim CNameWordNum,CNameStr,ENameWordNum,ENameStr,JSAddObj,JSNewsNum,JSNewsTitleNum,JSRowNum,JSContentNum,RsJSObj,RsJSSql
	 if NoCSSHackAdmin(request.form("CName"),"名称")<>"" then
	    CNameStr = Replace(Replace(request.form("CName"),"""",""),"'","")
		CNameWordNum = Cint(Len(CNameStr))
		if CNameWordNum>25 then
			 response.Write("<script>alert(""中文名称不能多于25个字符"");history.back();</script>")
			 response.end
		end if
	  else
		 response.Write("<script>alert(""请输入中文名称"");history.back();</script>")
		 response.end
	 end if
  	 if isnumeric(request.form("NewsNum")) = false then
		 response.Write("<script>alert(""新闻调用数必须为数字型"");history.back();</script>")
		 response.end
	 else
		 JSNewsNum = Cint(request.form("NewsNum"))
	 end if
  	 if isnumeric(request.form("NewsTitleNum")) = false then
		 response.Write("<script>alert(""新闻标题字数必须为数字型"");history.back();</script>")
		 response.end
	 else
		 JSNewsTitleNum = Cint(request.form("NewsTitleNum"))
	 end if
  	 if isnumeric(request.form("RowNum")) = false or request.form("RowNum")="0" then
		 response.Write("<script>alert(""新闻并排条数必须为数字型且不能为0"");history.back();</script>")
		 response.end
	 else
		 JSRowNum = Cint(request.form("RowNum"))
	 end if
  	 if isnumeric(request.form("ContentNum")) = false then
		 response.Write("<script>alert(""新闻内容字数必须为数字型"");history.back();</script>")
		 response.end
	 else
		 JSContentNum = Cint(request.form("ContentNum"))
	 end if
	  Set RsJSObj=server.createobject(G_FS_RS)
	  RsJSSql="select * from FS_FreeJS where ID = "&JSID&""
	  RsJSObj.open RsJSSql,Conn,1,3
	  Dim TempEName
	  TempEName = Cstr(RsJSObj("EName"))
	  RsJSObj("CName") = Cstr(CNameStr)
	  RsJSObj("Type") = Cint(Replace(Replace(Request.Form("Type"),"""",""),"'",""))
	  if Request.Form("Type") = "0" then
		  RsJSObj("Manner") = Cint(Replace(Replace(Request.Form("Manner"),"""",""),"'",""))
	  else
		  RsJSObj("Manner") = Cint(Replace(Replace(Request.Form("MannerP"),"""",""),"'",""))
	  end if
	  if Request.Form("PicWidth")<>"" and isnull(Request.Form("PicWidth"))=false then
	     if isnumeric(Request.Form("PicWidth"))=true then
			  RsJSObj("PicWidth") = Cint(Request.Form("PicWidth"))
	      else
			 response.Write("<script>alert(""图片宽度必须为数字型"");history.back();</script>")
			 response.end
		  end if
	  end if
	  if Request.Form("PicHeight")<>"" and isnull(Request.Form("PicHeight"))=false then
	     if isnumeric(Request.Form("PicHeight"))=true then
			  RsJSObj("PicHeight") = Cint(Request.Form("PicHeight"))
	      else
			 response.Write("<script>alert(""图片高度必须为数字型"");history.back();</script>")
			 response.end
		  end if
	  end if
	  RsJSObj("NewsNum") = Cint(JSNewsNum)
	  RsJSObj("NewsTitleNum") = Cint(JSNewsTitleNum)
	  RsJSObj("TitleCSS") = Cstr(Request.Form("TitleCSS"))
	  RsJSObj("ContentCSS") = Cstr(Request.Form("ContentCSS"))
	  RsJSObj("BackCSS") = Cstr(Request.Form("BackCSS"))
	  RsJSObj("RowNum") = Cint(JSRowNum)
	  if Request.Form("MannerP")="12" or Request.Form("MannerP")="16" then
		  if Replace(Replace(Request.Form("PicPath"),"""",""),"'","")<>"" then
			  RsJSObj("PicPath") = Cstr(Request.Form("PicPath"))
		  else
			 response.Write("<script>alert(""请选择图片地址"");history.back();</script>")
			 response.end
		  end if
	  end if
	  if Request.Form("ShowTimeTF")="1" then
		  RsJSObj("ShowTimeTF") = Cint(Request.Form("ShowTimeTF"))
	   else
		  RsJSObj("ShowTimeTF") = "0"
	  end if
	  RsJSObj("ContentNum") = Cint(JSContentNum)
	  RsJSObj("NaviPic") = Cstr(Request.Form("NaviPic"))
	  if Request.Form("DateType")="" or isnull(Request.Form("DateType")) or isnumeric(Request.Form("DateType"))=false then
		  RsJSObj("DateType") = "11"
	  else
		  RsJSObj("DateType") = Cint(Request.Form("DateType"))
	  end if
	  RsJSObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	  RsJSObj("Info") = Request.Form("Info")
	  RsJSObj("MoreContent") = Request.Form("MoreContent")
	  if Request.Form("MoreContent")=1 then
		  If Request.Form("LinkWord")<>"" and isnull(Request.Form("LinkWord"))=false then
			  RsJSObj("LinkWord") = Request.Form("LinkWord")
		  Else
		  	Response.Write("<script>alert(""请输入链接文字或图片"");</script>")
			Response.End
		  End If
		  If Request.Form("LinkCSS")<>"" or isnull(Request.Form("LinkCSS"))=false then
			  RsJSObj("LinkCSS") = Request.Form("LinkCSS")
		  End If
	  End If
	  If isnumeric(Request.Form("RowSpace")) then
		  RsJSObj("RowSpace") = Cint(Request.Form("RowSpace"))
	  Else
		  RsJSObj("RowSpace") = 2
	  End If
	  RsJSObj("RowBettween") = Request.Form("RowBettween")
	  RsJSObj("OpenMode") = Request.Form("OpenMode")
	  RsJSObj.update
	  RsJSObj.close
	  Set RsJSObj = Nothing
  '--------------------需要重新生成JS文件---------------------------------
	Dim JSClassObj,ReturnValue
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = TempSysRootDir
	Dim RefreshManner
	if Request.Form("Type") = "0" then
	  RefreshManner = Cint(Replace(Replace(Request.Form("Manner"),"""",""),"'",""))
	else
	  RefreshManner = Cint(Replace(Replace(Request.Form("MannerP"),"""",""),"'",""))
	end if
	Select case RefreshManner
		case "1"   ReturnValue = JSClassObj.WCssA(TempEName,True)
		case "2"   ReturnValue = JSClassObj.WCssB(TempEName,True)
		case "3"   ReturnValue = JSClassObj.WCssC(TempEName,True)
		case "4"   ReturnValue = JSClassObj.WCssD(TempEName,True)
		case "5"   ReturnValue = JSClassObj.WCssE(TempEName,True)
		case "6"   ReturnValue = JSClassObj.PCssA(TempEName,True)
		case "7"   ReturnValue = JSClassObj.PCssB(TempEName,True)
		case "8"   ReturnValue = JSClassObj.PCssC(TempEName,True)
		case "9"   ReturnValue = JSClassObj.PCssD(TempEName,True)
		case "10"   ReturnValue = JSClassObj.PCssE(TempEName,True)
		case "11"   ReturnValue = JSClassObj.PCssF(TempEName,True)
		case "12"  ReturnValue = JSClassObj.PCssG(TempEName,True)
		case "13"   ReturnValue = JSClassObj.PCssH(TempEName,True)
		case "14"   ReturnValue = JSClassObj.PCssI(TempEName,True)
		case "15"   ReturnValue = JSClassObj.PCssJ(TempEName,True)
		case "16"   ReturnValue = JSClassObj.PCssK(TempEName,True)
		case "17"   ReturnValue = JSClassObj.PCssL(TempEName,True)
	End Select
	Set JSClassObj = Nothing
  '--------------------需要重新生成JS文件---------------------------------
	if ReturnValue <> "" then
		Response.Write("<script>alert('" & ReturnValue & "');location='FreeJSList.asp'</script>")
	else
		Response.Redirect("FreeJSList.asp")
	end if
  end if

end if
Set Conn = Nothing
%>
