<%
'**************************************
'**		inc_UBB.asp
'**
'** 文件说明：UBB编辑栏
'** 修改日期：2005-04-07
'**************************************
if UBBcfg_font=1 or pagename="发布公告" or pagename="回复留言" then%>字体：<select onchange="if(this.options[this.selectedIndex].value!=''){showfont(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}" name="font">
<option value="宋体" selected>宋体</option>
<option value="楷体_GB2312">楷体</option>
<option value="仿宋_GB2312">仿宋</option>
<option value="黑体">黑体</option>
<option value="幼圆">幼圆</option>
<option value="隶书">隶书</option>
<option value="Verdana">Verdana</option>
<option value="Arial">Arial</option>
<option value="Tahoma">Tahoma</option>
</select>&nbsp;<%end if%><%if UBBcfg_size=1 or pagename="发布公告" or pagename="回复留言" then%>&nbsp;字号：<select name="size" onChange="if(this.options[this.selectedIndex].value!=''){showsize(this.options[this.selectedIndex].value);this.options[1].selected=true;}else {this.selectedIndex=0;}">
<option value="1">1</option>
<option value="2" selected>2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
</select>&nbsp;<%end if%><%if UBBcfg_size=1 or pagename="发布公告" or pagename="回复留言" then%>&nbsp;&nbsp;文字颜色：</font>
<SELECT onchange="if(this.options[this.selectedIndex].value!=''){showcolor(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}" name=color>
<option style="background-color:#555555;color: #555555" value="#555555" selected>#555555</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
<option style="background-color:#8B008B;color: #8B008B" value="#8B008B">#8B008B</option>
<option style="background-color:#95A298;color: #95A298" value="#95A298">#95A298</option>
<option style="background-color:#FF8C00;color: #FF8C00" value="#FF8C00">#FF8C00</option>
<option style="background-color:#1E90FF;color: #1E90FF" value="#1E90FF">#1E90FF</option>
<option style="background-color:#0175B4;color: #0175B4" value="#0175B4">#0175B4</option>
<option style="background-color:#008000;color: #008000" value="#008000">#008000</option>
<option style="background-color:#20B2AA;color: #20B2AA" value="#20B2AA">#20B2AA</option>
<option style="background-color:#778899;color: #778899" value="#778899">#778899</option>
<option style="background-color:#88C609;color: #88C609" value="#88C609">#88C609</option>
<option style="background-color:#0175B4;color: #0175B4" value="#0175B4">#0175B4</option>
<option style="background-color:#FFA500;color: #FFA500" value="#FFA500">#FFA500</option>
<option style="background-color:#DA70D6;color: #DA70D6" value="#DA70D6">#DA70D6</option>
<option style="background-color:#0099FF;color: #0099FF" value="#0099FF">#0099FF</option>
<option style="background-color:#9ACD32;color: #9ACD32" value="#9ACD32">#9ACD32</option>
</SELECT><%end if%><br>
<%if UBBcfg_b=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cbold()" src="images/bold.gif" width="23" height="22" alt="粗体" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_i=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Citalic()" src="images/italicize.gif" width="23" height="22" alt="斜体" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_u=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cunder()" src="images/underline.gif" width="23" height="22" alt="下划线" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_center=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Ccenter()" src="images/center.gif" width="23" height="22" alt="居中" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_URL=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Curl()" src="images/url1.gif" width="23" height="22" alt="超链接" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_email=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cemail()" src="images/email1.gif" width="23" height="22" alt="添加Email链接" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_glow=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cguang()" width="23" height="22" alt="添加发光字" src="images/glow.gif" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_shadow=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cying()" width="23" height="22" alt="添加阴影字" src="images/shadow.gif" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_pic=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cimage()" width="23" height="22" alt="添加图片" src="images/pic.gif" border="0">&nbsp;&nbsp;<%end if%><%if UBBcfg_swf=1 or pagename="发布公告" or pagename="回复留言" then%><img onclick="Cswf()" width="23" height="22" alt="添加 Flash 动画" src="images/swf.gif" border="0"><%end if%><br>