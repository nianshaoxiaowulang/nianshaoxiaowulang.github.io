<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>菜单</title>
</head>
<script language="JavaScript">
var iFrameSize = 1;
var aszFrameCols = "";
var szOldColumnsSize = "157,*";
var iCurrentSize = 0;
var iChangeSize = 25;
var iFadeInterval = 20;
var m_MinFrameSize = 10;
var bShrinkPending = false;
var bBusy = false;
var bShrinkPending = false;
var bEnlargePending = false;
var iCurrentEnlargeCount = 0;
var constMaxEnlargeCount = 200;
var m_OrigMouseOver = "";
var m_OrigMouseOut = "";
var bPinned = true;
var m_loaded = false;
var szVisibleStartTree = "<%=(Request.QueryString("TreeVisible"))%>";
function ShrinkFrame() {
	ResizeFrame();
}
function ResizeFrame(Direction)
{
	var NavFrame = top.document.getElementById('BottomFrameSet');
	if (Direction != null)
	{
		if ((iFrameSize == 1) && (Direction == 1))	return;
		if ((iFrameSize == 0) && (Direction == 0))	return;
	}

	bBusy = true;
	iCurrentSize = 0;
	if (((Direction == null) && (iFrameSize == 1)) || (Direction == 0))
	{
		szOldColumnsSize = NavFrame.cols;
		if (top.IsBrowserIE() && (!((navigator.userAgent.toLowerCase()).indexOf("msie 5.") == -1)) )
		{ 
			var tmpFrameSizes = NavFrame.cols.split(",");
			tmpFrameSizes[0] = (parseInt(document.body.clientWidth) + (iIe5FrameSizeAdjustment * 1));
			szOldColumnsSize = "";
			for (var iLoop = 0; iLoop < tmpFrameSizes.length; iLoop++)
			{
				szOldColumnsSize += tmpFrameSizes[iLoop].toString() + ",";
			}
			szOldColumnsSize = szOldColumnsSize.substring(0, (szOldColumnsSize.length - 1));
		}
		aszFrameCols = szOldColumnsSize.split(",");
		iCurrentSize = aszFrameCols[0];
		setTimeout("FadeFrame(0)", 50);
	}
	else
	{
		document.getElementById("div2").style.display = "";
		aszFrameCols = szOldColumnsSize.split(",");
		iCurrentSize = 0;
		setTimeout("FadeFrame(1)", 50);
	}
}

function FadeFrame(Direction)
{
	var NavFrame = top.document.getElementById('BottomFrameSet');
	if (Direction == 0)
	{
		iCurrentSize = iCurrentSize - iChangeSize;
		if (iCurrentSize <= m_MinFrameSize)
		{
			iCurrentSize = m_MinFrameSize;
			iFrameSize = 0;
			bBusy = false;
			document.body.onmouseover = m_OrigMouseOver;
			document.body.onmouseout = m_OrigMouseOut;
		}
		else
		{
			setTimeout("FadeFrame(" + Direction + ")", iFadeInterval);
		}
	}
	else
	{
		iCurrentSize = iCurrentSize + iChangeSize;
		if (iCurrentSize >= aszFrameCols[0])
		{
			iCurrentSize = aszFrameCols[0];
			iFrameSize = 1;
			bBusy = false;
			document.body.onmouseover = null;
			document.body.onmouseout = null;
		}
		else
		{
			setTimeout("FadeFrame(" + Direction + ")", iFadeInterval);
		}
	}
	NavFrame.cols = iCurrentSize + ",*";
	if (iCurrentSize <= m_MinFrameSize) iCurrentSize = 0;
	document.getElementById("div2").style.left = (iCurrentSize - aszFrameCols[0]) + "px";
}

function StartShrink(e)
{
	var myTarget;
	if (document.all) myTarget = e.toElement;
	else myTarget = e.relatedTarget;
	if (myTarget == null) BeginShrink();
	return;
}

function BeginShrink ()
{
	bEnlargePending = false;
	if (bPinned == false)
	{
		if (iFrameSize == 0)
		{
			if (bBusy == true)
			{
				bShrinkPending = true;
				setTimeout("PollShrink();", 100);
				return;
			}
		}
		else
		{
			if (bBusy == false)
			{
				iCurrentShrinkCount = 0;
				if (bShrinkPending == false)
				{
					bShrinkPending = true;
					setTimeout("DelayShrink()", 10);
				}
			}
		}
	}
}

function StartEnlarge(e)
{
	var myTarget;
	if (!CanShowNavTree())	return;
	if (document.all) myTarget = e.fromElement;
	else myTarget = e.relatedTarget;
	if (myTarget == null) setTimeout("BeginEnlarge();", 1);
	return;
}

function CanShowNavTree()
{
	if (typeof(top.CanShowNavTree) == "function") return (top.CanShowNavTree());
	return (true);
}

function BeginEnlarge ()
{
	bShrinkPending = false;
	if (iFrameSize == 1)
	{
		if (bBusy == true)
		{
			bEnlargePending = true;
			setTimeout("PollEnlarge();", 100);
			return;
		}
	}
	else
	{
		if (bBusy == false)
		{
			iCurrentEnlargeCount = 0;
			if (bEnlargePending == false)
			{
				bEnlargePending = true;
				setTimeout("DelayEnlarge()", 10);
			}
		}
	}
}

function PollEnlarge()
{
	if (bEnlargePending == true)
	{
		if (bBusy == true)
		{
			setTimeout("PollEnlarge();", 100);
			return;
		}
		else
		{
			iCurrentEnlargekCount = 0;
			setTimeout("DelayEnlarge()", 10);
		}
	}
}

function DelayEnlarge ()
{
	if (bEnlargePending == true)
	{
		iCurrentEnlargeCount += 10;
		if (iCurrentEnlargeCount > constMaxEnlargeCount)
		{
			iCurrentEnlargeCount = 0;
			ResizeFrame();
			return;
		}
		setTimeout("DelayEnlarge()", 10);
	}
}

function Startup ()
{
	m_loaded = true;
	m_OrigMouseOver = document.body.onmouseover;
	m_OrigMouseOut = document.body.onmouseout;
	return;
	if (IsTreeLoaded() == false)
	{
		setTimeout("Startup()", 100);
	}
	else
	{
		SetTreeVisible(szVisibleStartTree);
	}
}

function SetTreeVisible (treeName, reload)
{
	if (treeName == null) treeName = "content";
	treeName = treeName.toLowerCase();
}

function IsTreeLoaded ()
{
	var retValue = false;
	var localTreeObj = top.frames["ek_nav_bottom"]["NavIframeContainer"];
	if (localTreeObj != null)
	{
		localTreeObj = localTreeObj.frames["nav_folder_area"];
		if (localTreeObj != null)
		{
			if ((typeof(localTreeObj.IsLoaded)).toLowerCase() != "undefined")
			{
				if (localTreeObj.IsLoaded() == true) retValue = true;
			}
		}
	}
	return (retValue);
}
</script>
<body leftmargin="0" topmargin="0" class="UiNavigation" onload="Startup();" onmouseover="StartEnlarge(event);" onmouseout="StartShrink(event);">
<div id="div2" style="DISPLAY: block; LEFT: 0px; POSITION:relative"> 
  <iframe scrolling="auto" id="NavIframeContainer" name="NavIframeContainer" marginheight="0" marginwidth="0"  frameborder="0" height="100%" width="100%" src="Menu_Manage.asp"> 
  </iframe>
</div>
</body>
</html>
