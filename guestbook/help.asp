<% Option Explicit %>
<%response.buffer=true%>
<!--#include file="inc_common.asp"-->
<%
'**************************************
'**		help.asp
'**
'** 文件说明：帮助页面
'** 修改日期：2005-04-07
'**************************************

pagename="留言帮助"
call pageinfo()
mainpic="page_help.gif"
call skin1()
'---------------以下显示页面主体--------
%>
<br>
<div align="center">
	<center>
	<table cellpadding="8" cellspacing="1" width="95%" class="table1">
		<tr>
			<td width="100%" class="tablebody3">
			<p align="center"><b>UBB使用帮助：（以下效果均可用按钮实现）</b></td>
		</tr>
		<tr>
			<td width="100%" class="tablebody1"><br>
			<FONT FACE="Verdana,宋体">	[B]文字[/B]：在文字的位置可以任意加入您需要的字符，显示为粗体效果。 <br>
			<br>
			[I]文字[/I]：在文字的位置可以任意加入您需要的字符，显示为斜体效果。 <br>
			<br>
			[U]文字[/U]：在文字的位置可以任意加入您需要的字符，显示为下划线效果。 <br>
			<br>
			[center]文字[/center]：可以使文字居中显示。<br>
			<br>
			[URL]http://www.sohu.com[/URL] <br>
			<br>	[URL=http://www.sohu.com]搜狐[/URL]：有两种方法可以加入超级连接，可以连接具体地址或者文字连接。 <br>
			<br>
			[EMAIL]howlion@163.com[/EMAIL] <br>
			<br>	[EMAIL=MAILTO:howlion@163.com]王浩亮[/EMAIL]：有两种方法可以加入邮件连接，可以连接具体地址或者文字连接。 <br>
			<br>	[glow=255,red,2]文字[/glow]：在标签的中间插入文字可以实现文字发光特效，glow内属性依次为宽度、颜色和边界大小。 <br>
			<br>	[shadow=255,red,2]文字[/shadow]：在标签的中间插入文字可以实现文字阴影特效，shadow内属性依次为宽度、颜色和边界大小。 <br>
			<br>
			[size=数字]文字[/size]：输入您的字体大小，在标签的中间插入文字可以实现文字大小改变。 <br>
			<br>
			[face=字体]文字[/face]：输入您需要的字体，在标签的中间插入文字可以实现文字字体转换。 </FONT><p></td>
		</tr>
		<tr>
			<td width="100%" class="tablebody3"><B>提示：</B>如果留言没有达到你想要的效果，请注意是否正确使用了UBB代码，或是否支持这样的代码。</td>
		</tr>
	</table>
	</center>
</div>
<br>
<%
'--------------页面主体显示结束--------
call skin2()
%>