<%
'******************************************************************************************************
'asp���ı�������
'Powered lzkeyuan.cn
'Date��2013-05-26
'counter.asp��ͳ�Ƶ������
'counter.txt���������Ĵ������ı��ļ�
'ʹ�÷���������Ҫ���м�����ҳ���м��룺<script language="javascript" src="counter.asp"></script>
'ע��counter.asp��couner.txt�����ļ����õ�·��Ҫ��ȷ
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
