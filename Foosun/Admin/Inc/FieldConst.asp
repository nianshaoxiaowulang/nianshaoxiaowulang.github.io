<%
Dim NewsFieldName,NewsFieldEName,NewsFieldType,NewsClassFieldName,NewsClassFieldEName,NewsClassFieldType
Dim DownloadFieldName,DownloadFieldEName,DownloadFieldType
NewsFieldName = Array("编号","新闻编号","标题","副标题","标题样式","所属栏目编号","是否标题新闻","标题新闻路径","内容","是否用户投稿","新闻文件名","文件扩展名","新闻路径","添加日期","关键字","新闻来源","新闻作者","责任编辑","点击次数","是否推荐","是否图片新闻","图片路径","图片新闻文字导航","是否审核","是否删除","删除时间","新闻浏览权限","是否显示评论","是否允许评论","是否并排新闻","是否滚动新闻","是否公告新闻","是否显示内部链接","是否幻灯片新闻","新闻模板文件名","是否焦点图片新闻","是否精彩图片新闻","是否今日头条","所属专题编号","标题后显示评论")
NewsFieldEName = Array("ID","NewsID","Title","SubTitle","TitleStyle","ClassID","HeadNewsTF","HeadNewsPath","Content","ManuTF","FileName","FileExtName","Path","AddDate","KeyWords","TxtSource","Author","Editer","ClickNum","RecTF","PicNewsTF","PicPath","NaviWords","AuditTF","DelTF","DelTime","BrowPop","ShowReviewTF","ReviewTF","SBSNews","MarqueeNews","ProclaimNews","LinkTF","FilterNews","NewsTemplet","FocusNewsTF","ClassicalNewsTF","TodayNewsTF","SpecialID","TitleShowReview")
NewsFieldType = Array("16","100","100","100","100","100","116","100","230","116","100","100","100","7","100","100","100","100","116","116","116","100","100","116","116","7","116","116","116","116","116","116","116","116","100","116","116","116","230","116")
NewsClassFieldName = Array("编号","栏目编号","栏目英文名称","栏目中文名称","父栏目编号","子栏目数量","栏目生成模板","新闻捆绑模板","下载捆绑模板","是否允许投稿","是否显示","栏目添加时间","删除标记","删除时间","生成文件保存路径","生成文件扩展名","栏目浏览权限","捆绑二级域名","新闻归档时间","排序数字","是否为外部栏目","外部栏目地址","商品捆绑模板","栏目转向")
NewsClassFieldEName = Array("ID","ClassID","ClassEName","ClassCName","ParentID","ChildNum","ClassTemp","NewsTemp","DownLoadTemp","Contribution","ShowTF","AddTime","DelFlag","DelTime","SaveFilePath","FileExtName","BrowPop","DoMain","FileTime","Orders","IsOutClass","ClassLink","ProductTemp","RedirectList")
NewsClassFieldType = Array("16","100","100","100","100","116","100","100","100","116","116","7","116","7","100","100","116","100","7","116","116","100","100","100")
DownloadFieldName = Array("编号","下载编号","下载名称","所属栏目编号","版本号","下载类型","下载性质","语言","授权","文件大小","评价","系统平台","联系人EMAIL","提供者Url地址","开发商","显示图片","下载权限","简介","解压密码","添加时间","修改时间","是否推荐","是否审核","生成页面文件名","页面文件扩展名","下载次数","生成模板名","是否允许评论","是否显示评论")
DownloadFieldEName = Array("ID","DownLoadID","Name","ClassID","Version","Types","Property","Language","Accredit","FileSize","Appraise","SystemType","EMail","ProviderUrl","Provider","Pic","BrowPop","Description","PassWord","AddTime","EditTime","RecTF","AuditTF","FileName","FileExtName","ClickNum","NewsTemplet","ReviewTF","ShowReviewTF")
DownloadFieldType = Array("16","100","100","100","100","116","116","100","116","100","116","100","100","100","100","100","116","230","100","7","7","116","116","100","100","116","100","116","116")

'返回字段英文名在数组中的序号
Function GetIndexOfField(FieldEName,FieldENameArray)
	Dim i,FiledName
	FiledName = Lcase(FieldEName)
	For i = 0 to UBound(FieldENameArray)
		if lcase(FieldENameArray(i)) = FiledName Then
			GetIndexOfField = i
			Exit Function
		End if
	Next
	GetIndexOfField = -1
End Function
%>