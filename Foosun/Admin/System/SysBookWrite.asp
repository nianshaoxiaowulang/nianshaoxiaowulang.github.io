<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'软件名称：FoosunShop System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim RsAdminConfigObj
Set RsAdminConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent,QPoint from FS_Config")
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070701") then Call ReturnError1()
If Request.Form("Action")="Submit" then
		If request.form("Title")="" Or request.form("Content")="" then
			Response.Write("<script>alert(""请填写完整"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Len(request.form("Content"))>RsAdminConfigObj("MaxContent") then
			Response.Write("<script>alert(""内容不能超过"&RsAdminConfigObj("MaxContent")&"字符"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set Rs = server.createobject(G_FS_RS)
	  Sql1 = "Select * from FS_GBook where 1=0"
	  Rs.open sql1,conn,1,3
	  Rs.AddNew
	  Rs("Title")=NoCSSHackInput(Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/",""))
	  Rs("Content")=NoCSSHackContent(Request.Form("Content"))
	  Rs("FaceNum")=NoCSSHackInput(Replace(request.form("FaceNum"),"'",""))
	  Rs("AddTime")=Now
	  'Rs("QTime")=""
	  Rs("isQ")=0
	  Rs("QID")=0
	  Rs("EditQ")=""
	  Rs("UserID")=0
	  If Request.Form("isAdmin")<>"" then
		  Rs("isAdmin")=1 
	  Else
		  Rs("isAdmin")=0 
	  End if
	  If Request.Form("Orders")<>"" then
		  Rs("Orders")=1 
	  Else
		  Rs("Orders")=2 
	  End if
	  If Request.Form("isLock")<>"" then
		  Rs("isLock")=1 
	  Else
		  Rs("isLock")=0 
	  End if
	  Rs.update
	  Response.Write("<script>alert(""添加成功！"&CopyRight&""");location=""SysBook.asp"";</script>")  
	  Response.End
	  Rs.close
	  Set Rs=nothing
End if
Dim NewsContent
NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>FoosunCMS Shop 1.0.0930</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body scroll=yes topmargin="2" leftmargin="2">
<form action="" method="POST" name="NewsForm">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="221" valign="top"> 
      <div align="left"> 
        <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
          <tr bgcolor="#EEEEEE"> 
            <td height="26" colspan="5" valign="middle"> <table width="289" height="22" border="0" cellpadding="0" cellspacing="0">
                        <tr>
          <td width=35 align="center" alt="保存" onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; <input name="Action" type="hidden" id="Action" value="Submit">
                    <input name="Content" type="hidden" id="Content" value="<% = NewsContent %>"> 
                  </td>
        </tr>

              </table></td>
          </tr>
        </table>
        <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%" height="159" valign="top"> <table width="75%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td height="3"></td>
                  </tr>
                </table>
                <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                    <tr bgcolor="#F2F2F2"> 
                      <td width="16%"> <div align="right">帖子标题：</div></td>
                      <td width="84%"> <input name="Title" type="text" id="Title" size="30"> 
                        <input name="isAdmin" type="checkbox" id="isAdmin" value="1">
                        管理员可见 
                        <input name="Orders" type="checkbox" id="Orders" value="1">
                        固顶 
                        <input name="isLock" type="checkbox" id="isLock" value="1">
                        锁定</td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td bgcolor="#F2F2F2"> <div align="right">当前表情：</div></td>
                      <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td> <input name="FaceNum" type="radio" value="1" checked> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face1.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="2"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face2.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="3"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face3.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="4"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face4.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="5"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face5.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="6"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face6.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="7"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face7.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="8"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face8.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="9"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face9.gif" width="22" height="22"></td>
                          </tr>
                          <tr> 
                            <td> <input type="radio" name="FaceNum" value="10"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face10.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="11"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face11.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="12"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face12.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="13"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face13.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="14"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face14.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="15"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face15.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="16"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face16.gif" width="22" height="22"></td>
                            <td> <input type="radio" name="FaceNum" value="17"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face17.gif" width="22" height="22"> 
                            </td>
                            <td> <input type="radio" name="FaceNum" value="18"> 
                              <img src="../../../<%=UserDir%>/GBook/Images/face18.gif" width="22" height="22"> 
                            </td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr bgcolor="#F2F2F2"> 
                      <td colspan="2"> <div align="right"></div>
                        <iframe id='NewsContent' src='../../../<% = UserDir %>/Editer/BookNewsEditer.asp' frameborder=0 scrolling=no width='100%' height='320'></iframe></td>
                    </tr>
                 
                </table></TD>
            </TR>
          </TBODY>
        </TABLE>
      </div></td>
  </tr>
</table></form>
</body>
</html>
<script>
function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.NewsForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.NewsForm.Content.value=='') document.NewsForm.Content.value=TempContentArray[i];
			else document.NewsForm.Content.value=document.NewsForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.NewsForm.submit();
}
</script>
