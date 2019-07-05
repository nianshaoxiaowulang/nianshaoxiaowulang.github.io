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
Dim DBC,conn
Set DBC = new databaseclass
Set conn = DBC.openconnection()
Set DBC = nothing

Dim I,RsConfigObj
Set RsConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright from FS_Config")
'-------------------判断用户名和电子邮件是否注册
Dim RsMemberObj,RsMemberObj1
Set RsMemberObj = Conn.Execute("Select MemName,Email from FS_members where MemName = '" & replace(request.Form("Username"),"'","") &"'")    
If Not RsMemberObj.Eof then
	Response.Write("<script>alert(""用户名已经存在！请重新选择"&CopyRight&""");location=""javascript:history.back()"";</script>")
	Response.End
End if
Set RsMemberObj1 = Conn.Execute("Select MemName,Email from FS_members where Email ='" & trim(replace(request.Form("email"),"'","")) &"'")
If Not RsMemberObj1.Eof then
	Response.Write("<script>alert(""电子邮件已经存在！请重新选择"&CopyRight&""");location=""javascript:history.back()"";</script>")
	Response.End
End if
Dim Action
Action = "CheckLogin.ASP?UrlAddress=" & Request("UrlAddress")
Function GetCode()
	Dim TestObj
	On Error Resume Next
	Set TestObj = Server.CreateObject("Adodb.Stream")
	Set TestObj = Nothing
	If Err Then
		Dim TempNum
		Randomize timer
		TempNum = cint(8999*Rnd+1000)
		Session("GetCode") = TempNum
		GetCode = Session("GetCode")		
	Else
		GetCode = "<img src=""Comm/GetCode.asp"" onclick='this.src=this.src;' style='cursor:pointer'>"		
	End If
End Function
%>
<HTML><HEAD><TITLE><%=RsConfigObj("SiteName")%> >> 填写帐号信息</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="Css/UserCSS.css" type=text/css  rel=stylesheet></HEAD>
<BODY leftmargin="0" topmargin="10">
<div align="center">
  <script language="JavaScript" src="top.js" type="text/JavaScript"></script>
  <script language="javascript" src="Comm/MyScript.js"></script>
</div>
<TABLE cellSpacing=2 width="98%" align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=160>
      <TABLE cellSpacing=0 cellPadding=0 width=102 border=0>
        <TBODY>
        <TR>
          <TD><IMG height=27 src="images/favorite.left.help.jpg" 
        width=190></TD></TR>
        <TR>
          <TD>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="15%" height=30>&nbsp;</TD>
                                
                              <TD width="85%" class="f4"><strong>填写联系资料</strong></TD>
                              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD height=10><IMG height=1 src="" 
width=1></TD></TR></TBODY></TABLE>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#cccccc 
            cellSpacing=0 cellPadding=0 width="100%" border=1>
              <TBODY>
              <TR>
                <TD vAlign=top>
                  <TABLE class=bgup cellSpacing=0 cellPadding=0 width="100%" 
                  background="" border=0>
                    <TBODY>
                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD align=right>&nbsp;</TD></TR>
                    <TR>
                      <TD align=right width="12%" height=30>&nbsp;</TD>
                      <TD width="88%">
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD> 同意注册协议</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>填写帐号信息</TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD><font color="#FF0000">填写联系资料</font></TD>
                                    </TR>
                                    <TR> 
                                      <TD width=14 height=24><A><IMG id=KB1Img height=10 
                              src="images/arr2.gif" width=10></A></TD>
                                      <TD>注册成功</TD>
                                    </TR>
                                    <TR> 
                                      <TD><IMG height=5 src="images/SelfService.aspx" 
                              width=1></TD>
                                      <TD></TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
      <TD vAlign=top> <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center 
                  border=0>
          <TBODY>
            <TR> 
              <TD width="100%"> <TABLE width="100%" border=0>
                  <TBODY>
                    <TR> 
                      <TD width=26><IMG 
                              src="images/Favorite.OnArrow.gif" border=0></TD>
                      <TD 
class=f4>填写联系资料</TD>
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
              <form method=POST action="sRegister_Success.asp" name=UserForm1 onSubmit="return checkdata1()">
                <TD width="100%" height="74"> <div align="left"> <br>
                    <table width="98%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td width="19%"><div align="right"> 
                            <input name="Username" type="hidden" id="Username" value="<% = NoCSSHackInput(Request.Form("Username"))%>">
                            <input name="sPassword" type="hidden" id="sPassword" value="<% = Request.Form("sPassword")%>">
                            <input name="PassQuestion" type="hidden" id="PassQuestion" value="<% = Request.Form("PassQuestion")%>">
                            <input name="PassAnswer" type="hidden" id="PassAnswer" value="<% = Request.Form("PassAnswer")%>">
                            姓名：</div></td>
                        <td width="29%"> <font color="#FF0000"> 
                          <input name="sName" type="text" id="sName">
                          * </font></td>
                        <td width="52%">请填写您的真实姓名。</td>
                      </tr>
                      <tr> 
                        <td><div align="right">性别：</div></td>
                        <td> <input type="radio" name="Sex" value="0">
                          男 
                          <input type="radio" name="Sex" value="1">
                          女 <font color="#FF0000">*</font></td>
                        <td>请您选择性别。</td>
                      </tr>
                      <tr> 
                        <td><div align="right">出生日期：</div></td>
                        <td> <select name="yyear" id="yyear">
						<%
						For I=1949 to Year(Now)-6
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          年 
                          <select name="mmonth" id="mmonth">
						<%
						For I=1 to 12
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          月 
                          <select name="dday" id="dday">
						<%
						For I=1 to 31
						%>
                            <option value="<%=i%>"><%=i%></option>
						<%
						Next
						%>
                          </select>
                          日</td>
                        <td>请填写您的真实生日，该项用于取回密码。</td>
                      </tr>
                      <tr> 
                        <td><div align="right">证件类别：</div></td>
                        <td><font color="#FF0000"> 
                          <select name="VerGetType" id="VerGetType">
                            <option value="身份证" selected>身份证</option>
                            <option value="学生证">学生证</option>
                            <option value="军人证">军人证</option>
                            <option value="护照">护照</option>
                          </select>
                          </font> </td>
                        <td rowspan="2">有效证件作为取回帐号的最后手段，用以核实帐号的合法身份，请您务必如实填写。<br>
                          特别提醒：有效证件一旦设定，不可更改</td>
                      </tr>
                      <tr> 
                        <td><div align="right">证件号码：</div></td>
                        <td><font color="#FF0000"> 
                          <input name="VerGetCode" type="text" id="VerGetCode">
                          </font> </td>
                      </tr>
                      <tr> 
                        <td><div align="right">校验码：</div></td>
                        <td><font color="#FF0000"> 
                          <input name="Ver" type="text" id="Ver" size="10">
                          * 
                          <% = GetCode() %>
                          </font> </td>
                        <td>请将图中数字填入左边输入框中，该步骤有利于防止注册机。<br>
                          <font color="#FF0000">如果你长时间没有操作，请点击验证码刷新验证随机码</font></td>
                      </tr>
                      <tr> 
                        <td><div align="right">电话：</div></td>
                        <td><font color="#FF0000"> 
                          <input name="tel" type="text" id="tel">
                          * </font></td>
                        <td>可以填写多个号码，中间用&quot;,&quot;隔开</td>
                      </tr>
                      <tr> 
                        <td><div align="right">传真：</div></td>
                        <td><font color="#FF0000"> 
                          <input name="fax" type="text" id="fax">
                          </font></td>
                        <td>可以填写多个号码，中间用&quot;,&quot;隔开</td>
                      </tr>
                      <tr>
                        <td><div align="right">省份：</div></td>
                        <td colspan="2">
						     <select name="Province" id="Province">
                            <option value="四川" selected>四川</option>
                            <option value="广西" > 广西 </option>
                            <option value="广东" > 广东 </option>
                            <option value="北京" > 北京 </option>
                            <option value="海南" > 海南 </option>
                            <option value="福建" > 福建 </option>
                            <option value="天津" > 天津 </option>
                            <option value="湖南" > 湖南 </option>
                            <option value="湖北" > 湖北 </option>
                            <option value="河南" > 河南 </option>
                            <option value="河北" > 河北 </option>
                            <option value="山东" > 山东 </option>
                            <option value="山西" > 山西 </option>
                            <option value="黑龙江" > 黑龙江 </option>
                            <option value="辽宁" > 辽宁 </option>
                            <option value="上海" > 上海 </option>
                            <option value="甘肃" > 甘肃 </option>
                            <option value="青海" > 青海 </option>
                            <option value="新疆" > 新疆 </option>
                            <option value="西藏" > 西藏 </option>
                            <option value="宁夏" > 宁夏 </option>
                            <option value="云南" > 云南 </option>
                            <option value="吉林" > 吉林 </option>
                            <option value="内蒙古" > 内蒙古 </option>
                            <option value="陕西" > 陕西 </option>
                            <option value="安徽" > 安徽 </option>
                            <option value="贵州" > 贵州 </option>
                            <option value="江苏" > 江苏 </option>
                            <option value="重庆" > 重庆 </option>
                            <option value="浙江" > 浙江 </option>
                            <option value="江西" > 江西 </option>
                            <option value="国外" > 国外 </option>
                            <option value="台湾" > 台湾 </option>
                            <option value="香港" > 香港 </option>
                            <option value="澳门" > 澳门 </option>
                          </select>
                          城市：
                          <input name="City" type="text" id="City" size="12"></td>
                      </tr>
                      <tr> 
                        <td><div align="right">地址：</div></td>
                        <td colspan="2"><font color="#FF0000"> 
                          <input name="address" type="text" id="address" size="35">
                          * </font> 请务必详细填写该项目</td>
                      </tr>
                      <tr> 
                        <td><div align="right">邮政编码：</div></td>
                        <td colspan="2"><input name="PostCode" type="text" id="PostCode"></td>
                      </tr>
                      <tr> 
                        <td><div align="right">电子邮件：</div></td>
                        <td colspan="2"><font color="#FF0000"> 
                          <input name="email" type="text" id="email" value="<% = Request.Form("email")%>">
                          *</font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2"><input  type=submit name="Submit3" value="下一步" style="cursor:hand;"></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">注意：<br>
                          1．带*的栏目必须填写，否则注册不能继续！ <br>
                          2．推荐您使用网易2G超大免费邮箱@126.com，点击 <a href="http://reg.126.com/reg1.jsp" target="_blank"><font color="#FF0000">快速注册</font></a> 
                        </td>
                      </tr>
                    </table>
                  </div></TD></form>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>
  
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td> <hr size="1" noshade color="#FF6600">
      <div align="center">
        <% = RsConfigObj("Copyright") %>
      </div></td>
  </tr>
</table>
<BR>
</BODY></HTML>
<script language="JavaScript" type="text/JavaScript">
function CheckName(gotoURL) {
   var ssn=UserForm.Username.value.toLowerCase();
	   var open_url = gotoURL + "?Username=" + ssn;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
function CheckEmail(gotoURL) {
   var ssn1=UserForm.email.value.toLowerCase();
	   var open_url = gotoURL + "?email=" + ssn1;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=150,height=80');
}
</script>

