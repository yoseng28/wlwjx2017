
<head>
<style>
.container-fluid {padding-right: 0px; padding-left: 0px;}
.myTable{margin-left: 0px; background-color: #F0F0F0;}
.bgF0{background-color: #F0F0F0;}
.padding5 {padding: 5px;}
.bc1369c0 {background-color: #1369c0;}
.cFFF {color: #FFF;}
.h20 {height: 20px;}
</style>
</head>

<div class="container-fluid">
	<div class="row">
		<div class="col-md-1"></div>
		<div class="col-md-10 bgF0">
			<table class="myTable" width="100%" height="72" cellspacing="0" cellpadding="0">
				<tr>
					<td height="49" background=""> 
						<h4><span class="bc1369c0 cFFF padding5" style="border-radius: 5px;">”—«È¡¥Ω”</span></h4>
						</td>
				</tr>
				<tr>
		            <TD> 
		            	<p align="center" style="margin-top: 0; margin-bottom: 0"> 
		                        <%
									sql2="select *  from frendlink where reply='-1' "
									Set oRS= Server.CreateObject("ADODB.recordset")
									oRS.Open sql2,conn,1,3
									 if NOT oRS.EOF then 
								%>
		                    <table class="myTable" align="left" cellpadding="0" cellSpacing=0>
		                      <%
								 do while not oRS.eof
								  if NOT oRS.EOF then 
								   %>
		      					<tr> 
		                            <% for m=1 to 6
									  if NOT oRS.EOF then
									%>
		                            <td valign="middle"> 
		                            	<div align="center"> 
			                                <TABLE class="myTable" border=0 align="center" cellPadding=0 cellSpacing=0 width="188">
			                                  <TBODY>
			                                    <TR> 
			                                      <TD width="188" height=32 align="center">
													<p align="center" style="margin-top: 0; margin-bottom: 0">
													<font color="#666">&nbsp;</font><a href="<% =ors("weburl")  %>" target="_blank" style="text-decoration: none""><font color="#666"><%=left(oRS("webname"),38)%></font></a><font color="#666">
													</font> 
			                                      </TD>
			                                    </TR>
			                                  </TBODY>
			                                </TABLE>
		                              	</div>
		                            </td>
		                            <% oRS.MoveNEXT
									end if
									next %>
		                        </tr>
		                      	<%
								end if
								loop
								%>
		                    </table>
		                    <%
							else Response.Write ("‘›Œﬁƒ⁄»›.") end if
							oRS.close 
							Set oRS = Nothing
							%>
		            </TD>
		        </tr>
			</table>
		</div>
		<div class="col-md-1"></div>
	</div>
	<div class="h20"></div>
	<div class="row bg-primary">
		<div class="col-md-12">
			<TABLE width="100%" height="95" border=0 cellPadding=0 cellSpacing=0 bgcolor="">
	  			<TBODY>
				    <TR> 
				      <TD align=middle height=95 bgcolor=""> 
				        <TABLE width="100%" height="131" border=0 cellPadding=0 cellSpacing=0 bgcolor="" >
				          <TBODY>
				            <TR> 
				              <TD align=center height=16 bgcolor=""></TD>
				            </TR>
		            <TR> 
		              <TD align=center height=2></TD>
		            </TR>
		            <tr>
		              <TD align=center width="100%" bgcolor="">
						<table border="0" width="100%" height="96">
							<tr>
		              <TD align=center height=25 class="foot" > 
						<font face="Œ¢»Ì—≈∫⁄" color=""> 
						<p style="margin-top: 0; margin-bottom: 0; "> 
						<span style="font-weight: 400">√˚≥∆£∫<% =company %>°°µÿ÷∑: <% =address %>°°°°” ±‡£∫<% =code %> </span> </font> 
						<font color=""> </font></font></TD>
		            		</tr>
							<tr>
		              <TD align=center height=25 class="foot"> 
						<p style="margin-top: 0; margin-bottom: 0; "> 
						<font color="#FFFFFF"><span style="font-weight: 400">
						µÁª∞</span></font><font face="Œ¢»Ì—≈∫⁄" color="#FFFFFF"><a style="text-decoration: none" href="tel:09312135926"><span style="font-weight: 400"><font color="#FFFFFF">: 
		                <% =fax %>&nbsp;&nbsp;
		                &nbsp;</font></span></a></font><span style="font-weight: 400"><font face="Œ¢»Ì—≈∫⁄" color="#FFFFFF"> ÷ª˙£∫</font><font face="Œ¢»Ì—≈∫⁄" color="#DDEEFF"><font color="#DDEEFF"><a href="tel:13609316443" style="text-decoration: none"><font color="#FFFFFF"><% =phone %>°°°°</font></a></font><font face="Œ¢»Ì—≈∫⁄" color="#FFFFFF">” œ‰£∫<% =email %>&nbsp;&nbsp; 
		               °°			</font></font></span></TD>
		            		</tr>
							<tr>
		              	<TD align=center height=25 class="foot"> 
						<p style="margin-top: 0; margin-bottom: 0"> 
						<span style="font-weight: 400"> 
						<font color="#FFFFFF" face="Œ¢»Ì—≈∫⁄">
						Õ¯’æ∞Ê»®£∫</font></span>
							<span style="color:#FFFFFF; font-family: Œ¢»Ì—≈∫⁄; font-size: 14px;">ŒÔ¡™Õ¯ΩÃ—ßÕ≈∂”</span>
		              	</TD>
		            		</tr>
							</table>
			            </TR>
			          		</tr>
			            <TR> 
			              <TD align=center width="100%" bgcolor="">
							°°</TR>
			          </TBODY>
		        </TABLE></TD>
			    </TR>
			  </TBODY>
			</TABLE>
		</div>
	</div>
</div>