<%
'++++++++++////////////注意：单引号"'"不能去掉，请不要使用回车////////////++++++++++
'////////////建议只改字符，不要增加、删除////////////
'==========================生成目录参数=======================================
''-----管理目录,必须带/,不带虚拟目录-----
Const AdminDir = "Foosun/Admin/"
'-----统计 投票和广告的目录-----
Const PlusDir = "Foosun_Plus" 
'-----用户目录-----
Const UserDir = "Users"
'-----生成文件保存路径,后面不能带/,不带虚拟目录-----
Const ClassDir = "info"
'-----系统的虚拟目录,后面不能带/——-----
Const SysRootDir = ""
'-----文件目录,后面不能带/,不带虚拟目录-----
Const UpFiles = "Files"
'-----自由标签样式文件目录,后面不能带/,不带虚拟目录-----
Const StyleFiles = "Templetsskyim/FreeLableStyle"
'-----模板文件目录,,后面不能带/,不带虚拟目录-----
Const TempletDir = "Templetsskyim"
'-----远程图片保存目录,后面不能带/,不能带虚拟目-----
Const BeyondPicDir = "BeyondPic"
'-----新闻采集远程图片保存地址,后面不能带/,不带虚拟目录-----
Const SaveImagePath = "BeyondPic"
'-----下载文件存放目录-----
Const DownLoadDir = "DownLoad"
'-----归档新闻列表文件保存路径,后面不能带/,不带虚拟目录-----
Const RecordNewsListSavePath = "oldrecord"
'-----数据库连接路径,请不要删除,请注意要加虚拟目录-----
Const DataBaseConnectStr = "/FooSun_Data/FooSun_Data.asa"
'-----归档数据库联接路径,请不要删除,请注意要加虚拟目录-----
Const RecordDataBaseConnectStr = "/FooSun_Data/Record.asa"
'-----流量统计IP数据库路径,如有虚拟目录,请注意要加虚拟目录-----
Const IPDataBaseConnStr = "/FooSun_Data/AddressIp.mdb"
'-----新闻自动分页字符数，为0则不分页 分页字符一个汉字算2个 分页字符数不包含Html标记-----
Const AutoPagesNum = 3000
'==========================邮件控制参数======================================= 
Const MailObject = "Jmail"		'邮件发送组件
Const MailServer = "mail.cooin.com"		'用来发送邮件的SMTP服务器
Const MailServerUserName = "skeen@cooin.com"		'登录用户名
Const MailServerPassword = "22222"		'登录密码
Const MailDomain = "cooin"	'域名
'==========================版本控制参数======================================= 
Const Copyright = "\n\n www.Skyim.Com 天空校园网V1.0"				'系统版本信息
'==========================用户自定义参数===================================== 
Const VariableStr = "JustForTest........NoUse!"
'++++++++++----------配置信息结束----------++++++++++ 
%>
<!--#include file="ConstOption.asp" -->
