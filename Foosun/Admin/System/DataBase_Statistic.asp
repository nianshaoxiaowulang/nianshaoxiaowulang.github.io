<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P0406001") then Call ReturnError1()
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>数据统计</title>
<head></head>
<body leftmargin="2" topmargin="2"  ondragstart="return false;" onselectstart="return false;" oncontextmenu="return false;">
<form action="" method="post" name="DBSForm">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="28" class="ButtonListLeft">
<div align="center"><strong>数据库统计</strong></div></td>
  </tr>
</table>
  <table width="100%"  border="0" cellpadding="0" cellspacing="0">
    <tr height="10"> 
      <td colspan="6" align="center"> 
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
			  
            <td width="100" height="20" class="LableUP" id="Administrator" onClick="ChooseFile(this);" Types="Administrator"> 
              <div align="center">管理员统计</div></td>
			  
            <td width="100" id="Members" class="LableDown" Types="Members" onClick="ChooseFile(this);">
<div align="center">会员统计</div></td>
			  
            <td width="100" id="NewsClass" class="LableDown" Types="NewsClass" onClick="ChooseFile(this);">
<div align="center">栏目统计</div></td>
			  
            <td width="100" id="News_Month" class="LableDown" Types="News_Month" onClick="ChooseFile(this);">
<div align="center">新闻统计(月份)</div></td>
			  
            <td width="100" id="News_Class" class="LableDown" Types="News_Class" onClick="ChooseFile(this);">
<div align="center">新闻统计(栏目)</div></td>
			  
            <td width="100" id="Contribution" class="LableDown" Types="Contribution" onClick="ChooseFile(this);">
<div align="center">稿件统计</div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="60" colspan="6"><div align="center">[统计人]:&nbsp;&nbsp;<font color="red"><%=Session("Name")%></font>&nbsp;&nbsp;&nbsp;[统计时间]:&nbsp;&nbsp;<font color="red"><%=Now()%></font></div></td>
    </tr>
    <tr valign="top" id="ChooseYearClass" style="display:none"> 
      <td height="60" colspan="6"><div align="center">统计年份:&nbsp;&nbsp;&nbsp; 
          <select name="SelYearClass" id="SelYearClass" style="width:10% ">
            <%
	  Dim TempYearObj,MinYear,MaxYear,i
	  Set TempYearObj = Conn.Execute("Select Min(AddTime),Max(AddTime) from FS_NewsClass")
	  If isnull(TempYearObj(0))=false and isnull(TempYearObj(1))=false then
		  MinYear = Year(TempYearObj(0))
		  MaxYear = Year(TempYearObj(1))
		  For i = MinYear to MaxYear
		  %>
            <option value="<%=i%>" <%If Cint(i)=Year(Now()) then Response.Write("selected") end if%>><%=i%></option>
            <%
		  Next
	  End If
	  TempYearObj.Close
	  Set TempYearObj = Nothing
	  %>
          </select>
          &nbsp; 
          <input type="button" name="chooseY" value=" 确 定 " onClick="ChooseYearss();">
        </div></td>
    </tr>
    <tr valign="top" id="ChooseSpace" style="display:"> 
      <td height="60" colspan="6">&nbsp;</td>
    </tr>
    <tr valign="top" id="ChooseYear" style="display:none"> 
      <td height="60" colspan="6"><div align="center"> 统计年份:&nbsp;&nbsp;&nbsp; 
          <select name="SelYear" id="SelYear" style="width:10% ">
            <%
	  Set TempYearObj = Conn.Execute("Select Min(AddDate),Max(AddDate) from FS_News")
	  If isnull(TempYearObj(0))=false and isnull(TempYearObj(1))=false then
	  MinYear = Year(TempYearObj(0))
	  MaxYear = Year(TempYearObj(1))
	  For i=MinYear to MaxYear
	  %>
            <option value="<%=i%>" <%If Cint(i)=Year(Now()) then Response.Write("selected") end if%>><%=i%></option>
            <%
	  Next
	  End If
	  TempYearObj.Close
	  Set TempYearObj = Nothing
	  %>
          </select>
          &nbsp; 
          <input type="button" name="chooseY" value=" 确 定 " onClick="ChooseYears();">
        </div></td>
    </tr>
    <tr align="center" valign="middle"> 
      <td colspan="6" id="DataView"> <div align="center"> 
          <iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=Administrator" style="width:100%;height:400;"  frameborder=0></iframe>
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
<script>
function ChooseFile(Obj)
{
	switch (Obj.Types) 
	{
	 case 'Administrator':
	 ChooseYear.style.display = "none";
	 ChooseYearClass.style.display = "none";
	 ChooseSpace.style.display = '';
	 Administrator.className='LableUP';
	 Members.className='LableDown';
	 NewsClass.className='LableDown';
	 News_Month.className='LableDown';
	 News_Class.className='LableDown';
	 Contribution.className='LableDown';
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=Administrator" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
		
	 case 'Members':
	 ChooseYear.style.display = "none";
	 ChooseYearClass.style.display = "none";
	 ChooseSpace.style.display = '';
	 Administrator.className='LableDown';
	 Members.className='LableUP';
	 NewsClass.className='LableDown';
	 News_Month.className='LableDown';
	 News_Class.className='LableDown';
	 Contribution.className='LableDown';
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=Members" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
		
	 case 'NewsClass':
	 ChooseYear.style.display = "none";
	 ChooseYearClass.style.display = '';
	 ChooseSpace.style.display = "none";
	 NewsClass.className = 'LableUP';
	 Administrator.className='LableDown';
	 Members.className='LableDown';
	 News_Month.className='LableDown';
	 News_Class.className='LableDown';
	 Contribution.className='LableDown';
	 var NowDate,YearStr;
	 NowDate=new Date();
	 YearStr=NowDate.getYear();
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=NewsClass&ChooseYear='+YearStr+'" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
		
	 case 'News_Month':
	 ChooseYear.style.display = '';
	 ChooseYearClass.style.display = "none";
	 ChooseSpace.style.display = "none";
	 News_Month.className = 'LableUP';
	 NewsClass.className = 'LableDown';
	 Administrator.className='LableDown';
	 Members.className='LableDown';
	 News_Class.className='LableDown';
	 Contribution.className='LableDown';
	 var NowDate,YearStr;
	 NowDate=new Date();
	 YearStr=NowDate.getYear();
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=News_Month&ChooseYear='+YearStr+'" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
		
	 case 'News_Class':
	 ChooseYear.style.display = "none";
	 ChooseYearClass.style.display = "none";
	 ChooseSpace.style.display = '';
	 News_Class.className = 'LableUP';
	 NewsClass.className = 'LableDown';
	 Administrator.className='LableDown';
	 Members.className='LableDown';
	 News_Month.className='LableDown';
	 Contribution.className='LableDown';
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=News_Class" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
		
	 case 'Contribution':
	 ChooseYear.style.display = "none";
	 ChooseYearClass.style.display = "none";
	 ChooseSpace.style.display = '';
	 Contribution.className = 'LableUP';
	 News_Class.className = 'LableDown';
	 NewsClass.className = 'LableDown';
	 Administrator.className='LableDown';
	 Members.className='LableDown';
	 News_Month.className='LableDown';
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=Contribution" style="width:100%;height:400;"  frameborder=0></iframe>';
	 	break;
	 }
 }
 
 function ChooseYears()
 {
 	var ChYear = document.DBSForm.SelYear.value;
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=News_Month&ChooseYear='+ChYear+'" style="width:100%;height:400;"  frameborder=0></iframe>';
  }
  
  function ChooseYearss()
  {
 	var ChYear = document.DBSForm.SelYearClass.value;
	 document.all.DataView.innerHTML='<iframe id="MenuWindow" scrolling="no" src="DataBaseStatView.asp?Types=NewsClass&ChooseYear='+ChYear+'" style="width:100%;height:400;"  frameborder=0></iframe>';
   }
</script>