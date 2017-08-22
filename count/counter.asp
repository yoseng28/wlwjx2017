<%
'******************************************************************************************************
'asp简单文本记数器
'Powered lzkeyuan.cn
'Date：2013-05-26
'counter.asp：统计点击次数
'counter.txt：保存点击的次数的文本文件
'使用方法：在需要进行记数的页面中加入：<script language="javascript" src="counter.asp"></script>
'注意counter.asp，couner.txt两个文件引用的路径要正确
'******************************************************************************************************

dim path,myFile,read,write,cntNum
path=server.mappath("bijie520.txt")
read=1
write=2
Set myFso = Server.CreateObject("Scripting.FileSystemObject")
set myFile = myFso.opentextfile(path,read)
cntNum=myFile.ReadLine
myFile.close
cntNum=cntNum+1
set myFile = myFso.opentextfile(path,write,TRUE)
myFile.write(cntNum)
myFile.close
set myFile=nothing
set myFso=nothing
counterlen=len(cntNum)
%>
<%for i=1 to 10-counterlen%>	
document.write('<img src=\"images/0.gif\" border=0></img>');
<%next
for i=1 to counterlen%>
document.write('<img src=\"images/<%=mid(cntNum,i,1)%>.gif\" border=0></img>');
<%next%>
