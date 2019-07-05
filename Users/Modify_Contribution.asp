<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
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
'==============================================================================
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop from FS_Config")
Set DBC = Nothing
%>
<!--#include file="Comm/User_Purview.Asp" -->
<% 
If request.Form("action")="add" then
		If Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/","")="" or request.form("Content")="" then
			Response.Write("<script>alert(""请填写稿件标题和内容"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
		If Cstr(Request.Form("ClassID")) = "" then
			Response.Write("<script>alert(""请选择投稿栏目"&CopyRight&""");location=""javascript:history.back(1)"";</script>")  
			Response.End
		End if
	  Dim Rs,Sql1
	  Set rs = server.createobject(G_FS_RS)
	  Sql1 = "select * from FS_Contribution where ContID='"&Replace(Replace(Request.Form("ContID"),"'",""),Chr(39),"")&"'"
	  Rs.open sql1,conn,1,3
	  Rs("Title")=NoCSSHackInput(Replace(Replace(Replace(request.form("Title"),"'",""),"\",""),"/",""))
	  Rs("SubTitle")=NoCSSHackInput(Replace(request.form("SubTitle"),"'",""))
	  Rs("Content")=NoCSSHackContent(request.Form("Content"))
	  Rs("AddTime")=now()
	  Rs("KeyWords")=NoCSSHackInput(Replace(request.form("KeyWords"),"'",""))
	  Rs("Author")=NoCSSHackInput(Replace(Request.Form("Author"),"'",""))
	  Rs("ClassID")=NoCSSHackInput(Cstr(Request.Form("ClassID")))
	  Rs.update
	  Rs.close()
	  Response.Write("<script>alert('稿件修改成功'); window.location=""User_contribution.asp"";</script>")
	  Response.End
	  Set Rs=nothing
End If
Dim RsUserObj
Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(Replace(session("MemName"),"'",""),Chr(39),"")&"' and Password = '"& Replace(Replace(session("MemPassword"),"'",""),Chr(39),"") &"'")
If RsUserObj.eof then
	Response.Write("<script>alert(""严重错误！"&CopyRight&""");location=""Login.asp"";</script>")  
    Response.End
End if

Dim ConID,ConModObj
	ConID = Replace(Replace(Request("ContID"),"'",""),Chr(39),"")
	If ConID = "" or isnull(ConID) then
		Response.Write("<script>alert(""参数传递错误"&CopyRight&""");location=""javascript:history.back()"";</script>")
		Response.end
	End If
	Set ConModObj = Conn.Execute("Select * from FS_Contribution where ContID='"&ConID&"'")
	If ConModObj.eof then
		Response.Write("<script>alert(""稿件已经被管理审核通过或是删除,请返回刷新再试"&CopyRight&""");location=""javascript:history.back()"";</script>")
		Response.end
	End If
	
Dim NewsContent
    If Request.Form("Content")<>"" then
		NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
	Else
		NewsContent = Replace(Replace(ConModObj("Content"),"""","%22"),"'","%27")
	End If
%>
<HTML><HEAD>
<TITLE><%=RsConfigObj("SiteName")%> >> 会员中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet>
</HEAD>
<BODY leftmargin="0" topmargin="5">
<div align="center"> </div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
    <TR> 
      <TD height="262" vAlign=top> 
        <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>修改稿件</TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="100%"> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#ff6633 height=4><IMG height=1 src="" 
                              width=1></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
                
              <TD width="100%" height="238" valign="top"> 
                <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    
                  <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=5 width="100%" border=1>
                    <TBODY>
                        <TR> 
                          
                        <TD height="233" vAlign=top> 
                          <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                            <TBODY>
                              <TR> 
                                <TD width="15%" height="26"> 
                                  <div align="left"> <font color="#000000"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="Add_Contribution.asp"><font color="#FF0000">我要投稿</font></a> 
                                    </font> </div></TD>
                                <TD width="17%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution.asp">未审核投稿</a></TD>
                                <TD width="43%"><img src="Images/arr2.gif" width="10"><img src="Images/arr2.gif" width="10"><a href="User_Contribution_Passed.asp">已审核投稿</a></TD>
                                <TD width="25%"> 
                                  <div align="center"></div></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <hr size="1" noshade>
                          <table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
                            <form action="" method="POST" name="NewsForm">
                              <tr bgcolor="#F0F0F0"> 
                                <td width="15%"> 
                                  <div align="right">&#31295;&#20214;&#26631;&#39064; 
                                  </div></td>
                                <td width="85%"> 
                                  <input name="Title" id="Title2" style="width:80% " value="<%=ConModObj("Title")%>"></td>
                              </tr>
                              <tr bgcolor="#F0F0F0"> 
                                <td> 
                                  <div align="right">&#21103; 
                                    &#26631; &#39064; </div></td>
                                <td> 
                                  <input name="SubTitle" id="SubTitle2" style="width:80% " value="<%=ConModObj("SubTitle")%>"></td>
                              </tr>
                              <tr bgcolor="#F0F0F0"> 
                                <td> 
                                  <div align="right">&#20851; 
                                    &#38190; &#23383;</div></td>
                                <td> 
                                  <input name="KeyWords" type="text" id="KeyWords2" style="width:42% " value="<%=ConModObj("KeyWords")%>"> 
                                  &nbsp; &nbsp;&nbsp; &#20316;&#32773; <input name="Author" type="hidden" style="width:41% " value="<%=ConModObj("Author")%>"> 
                                  <input name="Author" type="text" style="width:30% " value="<%=ConModObj("Author")%>" disabled> 
                                </td>
                              </tr>
                              <tr bgcolor="#F0F0F0"> 
                                <td> 
                                  <div align="right">&#25152;&#23646;&#26639;&#30446; 
                                  </div></td>
                                <td> 
        <select name="ClassID" id="ClassID" style="width:50% ">
	   <%
	   Dim UserAddRikerObj
	   Set UserAddRikerObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where DelFlag=0 and Contribution=1 order by AddTime desc")
	   Do While Not UserAddRikerObj.eof
	   %>
	   <option value="<%=UserAddRikerObj("ClassID")%>" <%If  Cstr(ConModObj("ClassID")) = Cstr(UserAddRikerObj("ClassID")) then Response.Write("selected")%>><%=UserAddRikerObj("ClassCName")%></option>
	   <%
	   	UserAddRikerObj.MoveNext
	   Loop
	   UserAddRikerObj.Close
	   Set UserAddRikerObj = Nothing
	   %>
		</select>
</td>
                              </tr>
                              <tr bgcolor="#F0F0F0"> 
                                <td colspan="2"> 
                                  <div align="center"> 
                                    <iframe id='NewsContent' src='Editer/NewsEditer.asp' frameborder=0 scrolling=no width='100%' height='360'></iframe>
                                  </div></td>
                              </tr>
                              <tr bgcolor="#F0F0F0"> 
                                <td colspan="2"> 
                                  <div align="center"> 
								  <input name="submitggg" type="button" onClick="SubmitFun();" value=" &#25552; &#20132; ">
								  <input name="reset" type="reset" value=" &#22797; &#20301; ">
								  <input type="hidden" name="Content" value="<% = NewsContent %>">
								  <input name="action" type="hidden" id="action" value="add">
                                    <input name="ContID" type="hidden" id="ContID" value="<%=Request("ContID")%>">
                                  </div></td>
                              </tr>
                            </form>
                          </table></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> 
      <div align="center">
        <hr size="1" noshade color="#FF6600">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
</BODY></HTML>
<%
RsConfigObj.Close
Set RsConfigObj = Nothing
RsUserObj.close
Set RsUserObj=nothing
Set Conn=nothing
%>
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

