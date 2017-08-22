<!--#include file="conn.asp"-->
<%
sql="select * from content where firstImageName<>'' order by ID DESC"
set rs=Server.CreateObject("ADODB.recordset")
Rs.Open sql,Conn,1,3
if not Rs.eof then
for i=1 to 5
  if NOT Rs.EOF then
 if i=1 then 
  img1=rs("firstImageName")
  url1="zwview.asp?id="&rs("id")
   texts1=left(RS("title"),25)

  else if i=2 then
     img2=rs("firstImageName")
     url2="zwview.asp?id="&rs("id")
           texts2=left(RS("title"),25)
	 else if i=3 then
	    img3=rs("firstImageName")
        url3="zwview.asp?id="&rs("id")
                texts3=left(RS("title"),25)

		else if i=4 then
		   	    img4=rs("firstImageName")
                url4="zwview.asp?id="&rs("id")
                texts4=left(RS("title"),25)
                else if i=5 then
		   	    img5=rs("firstImageName")
                url5="zwview.asp?id="&rs("id")
                texts5=left(RS("title"),25)
                 else if i=6 then
		   	    img6=rs("firstImageName")
                url6="zwview.asp?id="&rs("id")
                texts6=left(RS("title"),25)

				 else if i=7 then
		   	    img7=rs("firstImageName")
                url7="zwview.asp?id="&rs("id")
                texts7=left(RS("title"),25)

          end if

         end if
          end if

		end if
	end if
end if
end if
RS.MoveNEXT
end if
next
%>

	<script type="text/javascript">
	<!--
	
	var focus_width=320
	var focus_height=280
	var text_height=30
	var swf_height = focus_height+text_height
	var config='3|0xFF0000|0xFFFFFF|80|0xcccccc|0x0099ff|0x000000';
				//-- config 参数设置 -- 自动播放时间(秒)|文字颜色|文字背景色|文字背景透明度|按键数字颜色|当前按键颜色|普通按键色彩 --
				var pics='',links='', texts='';

	var pics='<% =img1 %>|<% =img2 %>|<% =img3 %>|<% =img4 %>|<% =img5 %> '
 var 
links='<% =url1 %>|<% =url2 %>|<% =url3 %>|<% =url4 %>|<% =url5 %>'
 var 
texts='<% =texts1 %>|<% =texts2 %>|<% =texts3 %>|<% =texts4 %>|<% =texts5 %>'
	
	document.write('<object ID="focus_flash" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
	document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="playswfnew.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">');
	document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
	document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
	document.write('<embed ID="focus_flash" src="playswfnew.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#C5C5C5" quality="high" width="'+ focus_width +'" height="'+ swf_height+'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');		
	document.write('</object>');
	
	//-->
	</script>

<%
else Response.Write ("暂无本类文章.") end if
RS.close 
Set RS = Nothing
%>