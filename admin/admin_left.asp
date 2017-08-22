<!--#include file="adminconn.asp"-->
<%
  if session("aleave")="" then
      response.redirect "adminlogin.asp"
	  response.end
  end if
%>
<html><head>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link href="images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
BODY {
	BACKGROUND:799ae1; MARGIN: 0px;
}

.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</style>
</head>

<BODY bgcolor="#0064A8">
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td valign="bottom" height="42"> <img height="38" src="images/title.gif" width="158" border="0"></td>
	</tr>
	<tr>
		
    <td class="menu_title" onmouseover="this.className='menu_title2';" onmouseout="this.className='menu_title';" background="images/title_bg_quit.gif" height="25"> 
      <span><b>&nbsp;<a target="_blank" href="../../../index.asp"><font color="215DC6">&#22238;&#21040;&#39318;&#39029;</font></a></b> 
      | <b> <a target="_top" href="adminlogin.asp?logout=ture"><font color="215DC6">&#36864;&#20986; 
      </font></a></span></td>
	</tr>
	<tr>
		<td align="center" onmouseover="aa('up')" onmouseout="StopScroll()">
		<font face="Webdings" color="#FFFFFF" style=cuorsr:hand>5</font> </td>
	</tr>
</table>

<table cellspacing="0" cellpadding="0" width="160" align="center">
  <tr> 
    <td width="158" height="25" background="images/menudown.gif" class="menu_title" id="imgmenu2" style="cuorsr:hand" onclick="showsubmenu(2)" onmouseover="this.className='menu_title2';" onmouseout="this.className='menu_title';"> 
      <span>&#25991;&#31456;&#31649;&#29702; </span></td>
  </tr>
  <tr> 
    <td id="submenu2" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"><span class="bh"> 
        </span>  
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a target="main" href="admin_addinfo.asp"><font color="000000">&#28155;&#21152;&#25991;&#31456;</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a target="main" href="admin_info.asp"><font color="000000">&#31649;&#29702;&#25991;&#31456;</font></a></td>
          </tr>
        </table>
        <span class="bh"> </span> </div>
      <br> </td>
  </tr>
  <tr> 
    <td id="imgmenu3" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>关于我们&#31649;&#29702;</span></td>
  </tr>
  <tr> 
    <td id="submenu3" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"><span class="bh"> 
        </span>  
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a target="main" href="about_addinfo.asp"><font color="#000000">&#28155;&#21152;&#20027;&#39064;</font> 
              </a></td>
          </tr>
          <tr> 
            <td height="20"><a href="about_info.asp" target="main"><font color="#000000">&#31649;&#29702;&#20027;&#39064;</font></a></td>
          </tr>
        </table>
      </div>
      <br> </td>
  </tr>
   <tr> 
    <td id="imgmenu15" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(15)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <b>
		<p style="margin-left: 10px"><font color="#0E75E7">图片视频&#31649;&#29702; </span>
		</font> </td>
  </tr>
  <tr> 
    <td id="submenu15" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a href="Product_add.asp" target="main"><font color="#000000">&#28155;&#21152;信息</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a href="Product_info.asp" target="main"><font color="#000000">信息&#31649;&#29702;</font></a></td>
          </tr>
          <tr>
            <td height="20">　</td>
          </tr>
        </table>
      </div>
      <br> </td>
  </tr>  

   <tr> 
    <td id="imgmenu104" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(104)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>下载文件管理 </span></td>
  </tr>
  <tr> 
    <td id="submenu104" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a href="Download_add.asp" target="main"><font color="#000000">添加文件</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a href="Download_info.asp" target="main"><font color="#000000">下载管理</font></a></td>
          </tr>
        </table>
      </div>
      <br> </td>
  </tr><tr> 
    <td id="imgmenu7" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(7)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>&#30041;&#35328;&#31649;&#29702;</span></td>
  </tr>
  <tr> 
    <td id="submenu7" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a target="main" href="guest_admin.asp"><font color="000000">&#30041;&#35328;&#31649;&#29702;</font></a><a target="main" href="vote_add.asp"></a><a target="main" href="../../../vote/admin_addvote.asp"></a></td>
          </tr>
        </table>
      </div>
      　</td>
  </tr>
  <tr> 
    <td id="imgmenu8" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(8)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>友情链接管理</span></td>
  </tr>
  <tr> 
    <td id="submenu8" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a href="friend_admin.asp" target="main"><font color="#000000">友情链接管理</font></a></td>
          </tr>
        </table>
      </div>
      <br> </td>
  </tr>
  <% if session("aleave")="super" then %>
  <tr> 
    <td id="imgmenu12" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(12)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>&#25968;&#25454;&#24211;&#31649;&#29702; </span></td>
  </tr>
  <tr> 
    <td valign="top" id="submenu12" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a href="databack.asp" target="main"><font color="#FF0000">&#25968;&#25454;&#24211;&#31649;&#29702;</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a href="Edit_Admin_UploadFile.asp?sUploadDir=databack" target="main">&#22791;&#20221;&#25968;&#25454;&#24211;&#31649;&#29702;</a></td>
          </tr>
          <tr> 
            <td height="20"><a target="main" href="datarar.asp"><font color="000000">&#22791;&#20221;&#21387;&#32553;&#20462;&#22797;</font></a></td>
          </tr>
        </table>
      </div>
      <br></td>
  </tr>
    <tr> 
    <td id="imgmenu13" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(13)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>&#31995;&#32479;&#35774;&#32622; </span></td>
  </tr>
  <tr> 
    <td valign="top" id="submenu13" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a target="main" href="SiteSetup.asp"><font color="000000">&#32593;&#31449;&#21442;&#25968;&#35774;&#32622;</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a href="ClassManage.asp" target="main"><font color="#0000FF">&#26639;&#30446;&#31649;&#29702;</font></a></td>
          </tr>
          <tr> 
            <td height="20"><a target="main" href="admin_admin.asp"><font color="000000">&#31649;&#29702;&#21592;&#31649;&#29702;</font></a></td>
          </tr>
        </table>
      </div>
      <br>
     </td>
  </tr>
  <%else%>
  <tr> 
    <td id="imgmenu13" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(13)" onmouseout="this.className='menu_title';" style="cuorsr:hand" background="images/menudown.gif" height="25"> 
      <span>&#23494;&#30721;&#20462;&#25913; </span></td>
  </tr>
  <tr> 
    <td valign="top" id="submenu13" style="DISPLAY: none"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td height="20"><a href="admin_adminmodify_check.asp" target="main"><font color="#0000FF">&#23494;&#30721;&#20462;&#25913;</font></a></td>
          </tr>
        </table>
      </div>
      <br>
     </td>
  </tr>
  <%end if%>
  <tr> 
    <td height="25" background="images/menudown.gif"><strong><font color="#0000FF">&#12288;</font><font color="#376ed0">&#31243;&#24207;&#21046;&#20316;</font></strong></td>
  </tr>
  <tr> 
    <td height="17"> <div class="sec_menu" style="WIDTH: 158px"> 
        <table width="94%" height="47" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="20"> <div align="left">&#20848;&#24030;&#31185;&#28304;&#65306;<font color="#FF0000">lzkeyuan.cn</font></div></td>
          </tr>
          <tr> 
            <td height="20">Q&#12288;&#12288;Q&#65306;664126937</td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
&nbsp; </div> 
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//&#36825;&#37324;100&#20026;&#28378;&#21160;&#36895;&#24230;
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}



function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}







function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="images/menudown.gif";
}
}



function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title="&#32593;&#31449;&#21518;&#21488;&#31649;&#29702;&#31995;&#32479;"; 
</script>
</BODY></HTML>
