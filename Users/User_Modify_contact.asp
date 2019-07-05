<% Option Explicit %>
<!--#include file="../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Const.asp" -->
<!--#include file="../Inc/Md5.asp" -->
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
	If Request.Form("action")="Update" then
		If Trim(Request.Form("Name"))="" Or Trim(Request.Form("Tel"))=""  Or Trim(Request.Form("Address"))=""  Then
			Response.Write("<script>alert(""请填写完整！"&CopyRight&""");location=""User_Modify_Contact.asp"";</script>")  
			Response.End
		End if
		Dim RsUpdateObj,SqlUpdate
		Set RsUpdateObj = server.CreateObject (G_FS_RS)
		SqlUpdate = "select * from FS_members where id="& Clng(Replace(Replace(Request.Form("id"),"'",""),Chr(39),""))
		RsUpdateObj.Open SqlUpdate,Conn,1,3
		RsUpdateObj("Name")=NoCSSHackInput(Replace(Request.Form("Name"),"'",""))
		If Replace(Request.Form("Sex"),"'","")="0" Then
			RsUpdateObj("Sex")=0
		Else
			RsUpdateObj("Sex")=1
		End if
		RsUpdateObj("Telephone")=NoCSSHackInput(Replace(Request.Form("tel"),"'",""))
		RsUpdateObj("Msn")=NoCSSHackInput(Replace(Request.Form("Msn"),"'",""))
		RsUpdateObj("Oicq")=NoCSSHackInput(Replace(Request.Form("qq"),"'",""))
		RsUpdateObj("Province")=NoCSSHackInput(Replace(Request.Form("Province"),"'",""))
		RsUpdateObj("City")=NoCSSHackInput(Replace(Request.Form("City"),"'",""))
		RsUpdateObj("Address")=NoCSSHackInput(Replace(Request.Form("address"),"'",""))
		RsUpdateObj("postcode")=NoCSSHackInput(Replace(Request.Form("postcode"),"'",""))
		RsUpdateObj("HomePage")=NoCSSHackInput(Replace(Request.Form("HomePage"),"'",""))
		RsUpdateObj("Birthday")=NoCSSHackInput(Replace(Request.Form("Birthday"),"'",""))
		RsUpdateObj.Update
		RsUpdateObj.Close
		Set RsUpdateObj=Nothing
		Response.Write("<script>alert(""更新联系资料成功！"&CopyRight&""");location=""User_Modify_Contact.asp"";</script>")  
		Response.End
	End if
	Dim RsUserObj
	Set RsUserObj = Conn.Execute("Select * From FS_Members where MemName = '"& Replace(session("MemName"),"'","")&"' and Password = '"& Replace(session("MemPassword"),"'","") &"'")
	If RsUserObj.eof then
		Response.Write("<script>alert(""严重错误！"&CopyRight&""");location=""Login.asp"";</script>")  
		Response.End
	End if
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
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>修改联系方式</TD>
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
              <form method=POST action="" name="UserForm1">
                <TD width="100%" height="159"> <div align="left"> 
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="3"></td>
                      </tr>
                    </table>
                    <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
                      <TBODY>
                        <TR> 
                          <TD vAlign=top> <TABLE class=bgup cellSpacing=0 cellPadding=5 width="100%" 
                  background="" border=0>
                              <TBODY>
                                <TR> 
                                  <TD width="95%" height="68"><div align="left"><font color="#000000"> 
                                      </font> 
                                      <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#E7E7E7">
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">会员编号：</div></td>
                                          <td><font color="#FF0000"> 
                                            <% = RsUserObj("UserNo") %>
                                            &nbsp;</font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td width="22%"> <div align="right">用户名：</div></td>
                                          <td width="78%"> <% = RsUserObj("MemName") %> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">姓名：</div></td>
                                          <td> <font color="#666666"> 
                                            <input name="name" type="text" id="name" value="<% = RsUserObj("name") %>">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font><font color="#666666">&nbsp; </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">性别：</div></td>
                                          <td><input type="radio" name="Sex" value="0" <%if RsUserObj("Sex")=0 Then Response.Write("checked")%>>
                                            男 
                                            <input type="radio" name="Sex" value="1" <%if RsUserObj("Sex")=1 Then Response.Write("checked")%>>
                                            女</td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">电话：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="tel" type="text" id="tel" value="<% = RsUserObj("Telephone") %>">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">MSN：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="msn" type="text" id="msn" value="<% = RsUserObj("msn") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">QQ：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="qq" type="text" id="qq" value="<% = RsUserObj("Oicq") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">省份：</div></td>
                                          <td><select name="Province" id="Province">
                                              <option value="四川" <%if RsUserObj("Province")="四川" Then Response.Write("selected")%>>四川</option>
                                              <option value="广西"  <%if RsUserObj("Province")="广西" Then Response.Write("selected")%>> 
                                              广西 </option>
                                              <option value="广东"  <%if RsUserObj("Province")="广东" Then Response.Write("selected")%>> 
                                              广东 </option>
                                              <option value="北京"  <%if RsUserObj("Province")="北京" Then Response.Write("selected")%>> 
                                              北京 </option>
                                              <option value="海南"  <%if RsUserObj("Province")="海南" Then Response.Write("selected")%>> 
                                              海南 </option>
                                              <option value="福建"  <%if RsUserObj("Province")="福建" Then Response.Write("selected")%>> 
                                              福建 </option>
                                              <option value="天津"  <%if RsUserObj("Province")="天津" Then Response.Write("selected")%>> 
                                              天津 </option>
                                              <option value="湖南"  <%if RsUserObj("Province")="湖南" Then Response.Write("selected")%>> 
                                              湖南 </option>
                                              <option value="湖北"  <%if RsUserObj("Province")="湖北" Then Response.Write("selected")%>> 
                                              湖北 </option>
                                              <option value="河南"  <%if RsUserObj("Province")="河南" Then Response.Write("selected")%>> 
                                              河南 </option>
                                              <option value="河北"  <%if RsUserObj("Province")="河北" Then Response.Write("selected")%>> 
                                              河北 </option>
                                              <option value="山东"  <%if RsUserObj("Province")="山东" Then Response.Write("selected")%>> 
                                              山东 </option>
                                              <option value="山西"  <%if RsUserObj("Province")="山西" Then Response.Write("selected")%>> 
                                              山西 </option>
                                              <option value="黑龙江"  <%if RsUserObj("Province")="黑龙江" Then Response.Write("selected")%>> 
                                              黑龙江 </option>
                                              <option value="辽宁"  <%if RsUserObj("Province")="辽宁" Then Response.Write("selected")%>> 
                                              辽宁 </option>
                                              <option value="上海"  <%if RsUserObj("Province")="上海" Then Response.Write("selected")%>> 
                                              上海 </option>
                                              <option value="甘肃"  <%if RsUserObj("Province")="甘肃" Then Response.Write("selected")%>> 
                                              甘肃 </option>
                                              <option value="青海"  <%if RsUserObj("Province")="青海" Then Response.Write("selected")%>> 
                                              青海 </option>
                                              <option value="新疆"  <%if RsUserObj("Province")="新疆" Then Response.Write("selected")%>> 
                                              新疆 </option>
                                              <option value="西藏"  <%if RsUserObj("Province")="西藏" Then Response.Write("selected")%>> 
                                              西藏 </option>
                                              <option value="宁夏"  <%if RsUserObj("Province")="宁夏" Then Response.Write("selected")%>> 
                                              宁夏 </option>
                                              <option value="云南"  <%if RsUserObj("Province")="云南" Then Response.Write("selected")%>> 
                                              云南 </option>
                                              <option value="吉林"  <%if RsUserObj("Province")="吉林" Then Response.Write("selected")%>> 
                                              吉林 </option>
                                              <option value="内蒙古"  <%if RsUserObj("Province")="内蒙古" Then Response.Write("selected")%>> 
                                              内蒙古 </option>
                                              <option value="陕西"  <%if RsUserObj("Province")="陕西" Then Response.Write("selected")%>> 
                                              陕西 </option>
                                              <option value="安徽"  <%if RsUserObj("Province")="安徽" Then Response.Write("selected")%>> 
                                              安徽 </option>
                                              <option value="贵州"  <%if RsUserObj("Province")="贵州" Then Response.Write("selected")%>> 
                                              贵州 </option>
                                              <option value="江苏"  <%if RsUserObj("Province")="江苏" Then Response.Write("selected")%>> 
                                              江苏 </option>
                                              <option value="重庆"  <%if RsUserObj("Province")="重庆" Then Response.Write("selected")%>> 
                                              重庆 </option>
                                              <option value="浙江"  <%if RsUserObj("Province")="浙江" Then Response.Write("selected")%>> 
                                              浙江 </option>
                                              <option value="江西"  <%if RsUserObj("Province")="江西" Then Response.Write("selected")%>> 
                                              江西 </option>
                                              <option value="国外"  <%if RsUserObj("Province")="国外" Then Response.Write("selected")%>> 
                                              国外 </option>
                                              <option value="台湾"  <%if RsUserObj("Province")="台湾" Then Response.Write("selected")%>> 
                                              台湾 </option>
                                              <option value="香港"  <%if RsUserObj("Province")="香港" Then Response.Write("selected")%>> 
                                              香港 </option>
                                              <option value="澳门"  <%if RsUserObj("Province")="澳门" Then Response.Write("selected")%>> 
                                              澳门 </option>
                                            </select> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right"><span class="f41">城市</span>：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="city" type="text" id="city" value="<% = RsUserObj("City") %>">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">地址：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="address" type="text" id="address" value="<% = RsUserObj("address") %>" size="35">
                                            </font><font color="#FF0000">&nbsp; 
                                            * </font><font color="#666666">&nbsp; </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td><div align="right">邮政编码：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="postcode" type="text" id="postcode" value="<% = RsUserObj("postcode") %>">
                                            </font> </td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">网站地址：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="HomePage" type="text" id="HomePage" value="<% = RsUserObj("HomePage") %>" size="35">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td> <div align="right">出生日期：</div></td>
                                          <td><font color="#666666"> 
                                            <input name="Birthday" type="text" id="Birthday" value="<% = RsUserObj("Birthday") %>"  readonly>
                                            <input type="button" name="Submit4" value="选择日期" onClick="OpenWindowAndSetValue('Comm/SelectDate.asp',280,110,window,document.UserForm1.Birthday);document.UserForm1.Birthday.focus();">
                                            </font></td>
                                        </tr>
                                        <tr bgcolor="#FFFFFF"> 
                                          <td>&nbsp;</td>
                                          <td> <input type="submit" name="Submit" value="提交修改"> 
                                            <input name="id" type="hidden" id="id" value="<% = RsUserObj("ID") %>"> 
                                            <input name="action" type="hidden" id="action" value="Update">
<script language="JavaScript" type="text/JavaScript">
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
</script></td>
                                        </tr>
                                      </table>
                                      <font color="#000000"> </font> </div>
                                    <span class="f41"> </span> </TD>
                                </TR>
                              </TBODY>
                            </TABLE></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <strong></strong></div></TD>
              </form>
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

