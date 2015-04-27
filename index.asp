<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="imlm.asp"-->
<%
response.cookies("wh")=Request.QueryString
mm=request.Form("mm")
if mm=1 then
	ausername=request.Form("username")
	apassword=request.Form("password")
	bmpnum=Lcase(request.Form("bmpnum"))
	If Trim(Request.Form("validatecode"))=Empty Or Trim(Session("cnbruce.com_ValidateCode"))<>Trim(Request.Form("validatecode")) Then
         response.write"<SCRIPT language=JavaScript>alert('验证码有误！请重新输入！');"
         Response.Write"this.location.href='vbscript:history.back()';</SCRIPT>"
	 	response.end   
	end if
	if ausername="" or apassword="" then
	     response.write "<SCRIPT language=JavaScript>alert('对不起,帐号和密码不可以为空！');"
	     Response.Write"this.location.href='vbscript:history.back()';</SCRIPT>"
	     Response.End
	else
		exea="select * from imlm_user where username='"&ausername&"' and password='"&md5(apassword)&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open exea,conn,1,3
		if rs.eof then
			response.write "<SCRIPT language=JavaScript>alert('对不起,帐号或密码错误！请重新输入！');"
			Response.Write"this.location.href='vbscript:history.back()';</SCRIPT>"
			Response.End
		else
			rs("loginnum")=rs("loginnum")+1
			rs.update
			Session("imlmusername")=ausername
			Response.redirect "accounts.asp"
			response.end
		end if
		rs.close
		set rs=nothing
	end if
mm=0
end if
%>
<%
id=request("id")
souser=request.form("souser")
pagenum=request("pagenum")

if pagenum="" or pagenum<1 then
	pagenum=1
else
	pagenum=cint(pagenum)
end if

Set rs=Server.CreateObject("ADODB.recordset")
rs.cachesize=30
rs.cursortype=1

mysq="select * from imlm_jsml order by id desc"
rs.open mysq,conn,1,1
rs.pagesize=30
if pagenum>rs.pagecount then
pagenum=1
end if
if rs.eof then
else
rs.absolutepage=pagenum
end if
%>
<!--#include file="top.asp"-->
	<!--步骤图-->
	<div class="Topstep"><img src="images/index/step.gif"></div>
	<!--步骤图-->
	<!--首页通栏幻灯+登录窗-->
	<div class="IndexPpt">
		<div class="IpptBox">
			<div id="focus" style="overflow:hidden; width:990px; height:288px; overflow:hidden; position:relative;">
				<ul class="PptImg">
					<%=txt10%>
				</ul>
			</div>
			<%if not login then%>
			<!--登陆框-->
			<form method="POST">
			<div class="LoginBox" style="padding-top:10px; height:246px; top:16px;">
								<div class="LoginIpt" style="height:auto; background:url(images/index/yz_bg.png去掉输入背景);">
					<p>用户名：&nbsp;<input type="text" name="username" size="25" class="form" maxlength="16"></p>
					<p>密<span class="a1"></span>码：&nbsp;<input type="password" name="password" size="25" class="form"></p>
					<p id="tbLoginCode" >验证码：&nbsp;<input name="validatecode" type="text" size="15" class="yzm">
            &nbsp;&nbsp;<img src="validatecode.asp" width="41" height="11" border="0" align="absmiddle"></p>
				</div>
				<div class="LoginBtn"><p><input type=image src="images/Bgsprit1.png" name="B1" /></p></div>

				<div class="LoginTxt"><span><a href="reg.asp" class="LRegLink">免费注册</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span><a href="pass.asp" class="w12l">忘记密码</a></span></div>
				<div class="LoginTips">
										<p><a href="reg.asp" class="w12l" target="_blank">注册赠送20元</a></p>
										<p><a href="login.asp" class="w12l" target="_blank">每日登陆赠送3元</a></p>
			  </div>
							</div>
							<input type="hidden" name="mm" value="1">
							</form>
			<!--登陆框-->
			
			<%else%>
			<!--登陆登录后-->
			<div class="LoginBox" style="padding-top:10px; height:246px; top:16px;">
								<div class="jiand">
				  <div class="jiand_l"><img src="images/per_headimg.gif" width="100" height="100" /><div class="jiand_zh"><a href="accounts.asp">+ 我的账户</a></div></div>
				  <div class="jiand_r">
					<p>欢迎您！<span><a href="logout.asp" class="tuic">[退出]</a></span></p>
					<p style="margin-top:18px;"><div style="float:left;">普通会员：</div><div style="float:left;font-weight:bold; color:#ff6600; font-family:'宋体'; font-size:14px;"><%=imlmusername%></div></p>
					<p style="margin-top:18px; font-weight:bold; font-family:'宋体'; font-size:14px;">推广人数：<%
Set rsb=Server.CreateObject("ADODB.recordset")
mysq="select * from imlm_user where yesno and sx1='"&imlmusername&"'"
rsb.open mysq,conn,1,1
response.write rsb.recordcount&"  人"
rsb.close
set rsb=nothing%></p>
				  </div>
				</div>
								
				<div class="jiand2">
					<div class="jiand2_l" id="signinDiv">
						<div class="jiand2_l_top"><a target="_blank" href="/sys/1373439892"><img src="images/index/signin_txt.png" border="0" /></a></div>
												<div class="jiand2_l_tab">
												  <div class="jiand2_l_tab_an" id="signsub"><a href="accounts.asp"><img style="cursor:pointer;" onclick="signin();"  src="images/index/signin.png" border="0" /></a></div>
					  </div>
											</div>
					<div class="jiand2_r">
						<ul>
						  <li>
							<div class="txt">您的收入：<span style="font-weight:bold; color:#ff6600; font-family:'宋体'; font-size:14px;" id="iUserG"><%=formatnumber(ajine,2,-1)%>&nbsp;<%=adbname%></span><span class="jb"></span></div>
							<div class="txt2"><a href="change.asp">[我要提现]</a></div>
						  </li>
						  <li style="margin-top:8px;">
							<div class="txt">提现总额：<span style="font-weight:bold; color:#ff6600; font-family:'宋体'; font-size:14px;"><%=formatnumber(atjine,2,-1)%>&nbsp;<%=adbname%></span></div>
							<div class="txt2"><a href="mlog.asp">[提现记录]</a></div>
						  </li>
						</ul>
					</div>
					<script type="text/javascript">
					function formatNum(nStr){nStr += '';x = nStr.split('.');x1 = x[0];x2 = x.length > 1 ? '.' + x[1] : '';var rgx = /(\d+)(\d{3})/;while (rgx.test(x1)) {x1 = x1.replace(rgx, '$1' + ',' + '$2');}return x1 + x2;}
					function signin(){
						$('#signsub').html('<img src="images/index/signin.png" border="0" />');
						$.ajax({
							type: "POST",
							url: "ajax.php",
							dataType: "json",
							data : "act=signin&key="+Math.random(),
							success: function(strJson){
								if(strJson==null){
									return false;
								}
								else if(strJson.error==10000&&strJson.times!="undefined"){
									$('#signinDiv').html('<div class="jiand2_l_top" title="连续天数越久，签到奖励越丰厚！"><img src="images/index/signin_txt.png" /></div><div class="jiand2_l_tab2"><p style="color:#663333;">已连续签到</p><p style="font-size:14px; font-family:\'微软雅黑\'; font-weight:bold; color:#663333;">'+strJson.times+'天</p></div>');
									$('<div id="tscrollG" style="position:absolute;z-index:10; text-align:center; margin-top:-50px; margin-left:-170px;left:100%;font-weight:bold; color:#ff6600; font-family:\'宋体\'; font-size:12px;  background-color:#FFFFFF">签到有礼，随机奖励'+strJson.dayG+'乐币，赠送1次抽奖机会</div>').appendTo($("#signinDiv")).animate({opacity:'hide'},5000,'',function(){if(strJson.fullG>0){$('#tangc').show();}});
																		$('#iUserG').html(formatNum(strJson.userG));
									return true;
								}else{
									alert(strJson.msg);
								}
							}
						})
					}
					</script>
				</div>
							</div>
			<!--登陆框-->
			<%end if%>
		</div>
	</div>
	<!--首页通栏幻灯+登录窗-->
	<!--主体-->
	<div id="MainBox">
		<!--左边-->
		<div class="IleftBox">
			<!--体验-->
			<div class="TyBox">
				<h1>最新推荐任务<label><a href="task.asp">更多></a></label></h1>
				<ul class="TyMode">
					<%=txt12%>
				</ul>
				
				<h1 style="border-top:1px solid #dddddd;">最新高价任务<label><a href="task.asp">更多></a></label></h1>
				<ul class="TyMode">
					<%=txt13%>
				</ul>
				<script type="text/javascript">function upDiv(id,i){var name = 'on';if(i==0){name = 'on last';}$("#"+id).removeClass();$("#"+id).addClass(name);}function downDiv(id,i){var name = '';if(i==0){name = 'last';}$("#"+id).attr('class',name);}function hitsJump(type,ID){$.ajax({type:"POST",url:"ajax.php",dataType:"json",data:"act=hitsJump&type="+type+"&ID="+ID+"&key="+Math.random(),success:function(strJson){}})}</script>
			</div>
			<!--体验-->
		</div>
		<!--左边-->
		<!--右边-->
		<div class="IrightBox">
			<!--最新公告-->
			<div class="Inews" style="margin-bottom:10px;">
				<p class="title0" style="margin-bottom:3px;">最新任务<label><a href="task.asp">更多></a></label></p>
				<ul>
                            <%
exea="select top 9 * from imlm_renwu order by id desc"
set rsa=server.createobject("adodb.recordset")
rsa.open exea,conn,1,1
if not rsa.eof then
do while not rsa.eof
%>				
				<li><a href="tasktj.asp?id=<%=rsa("id")%>">最新可做：<font color="#999999"><%=strvalue(rsa("title"),26)%></font></a></li>
<%
rsa.movenext
loop
else
	response.write "<tr><td align=""center"">更新中...</td></tr>"
end if
rsa.close
set rsa=nothing
%>				
			  </ul>
			</div>
			<!--最新公告-->
			<!--最新公告-->
			<div class="Inews" style="height:192px;"><!--不间断滚动-原172-140-->
				<p class="title0" style="margin-bottom:2px;">最新支付</p>
				<div style="position:relative;height:152px;overflow:hidden; top:5px;">
					<div id="showList" style="position:relative;height:152px;overflow:hidden;">
						<ul id="showCash">
	  	<%
for j=1 to rs.pagesize
if rs.eof then exit for
usernametemp=rs("username")
lengs=len(usernametemp) 
usernametemp=left(usernametemp,lengs-2)&"**"
%>						
						<li><%=usernametemp%>&nbsp;&nbsp;已支付&nbsp;&nbsp;&nbsp;<span class="Rede1"><%=rs("jine")%> <%=adbname%></span></li>
			<%
rs.movenext
if rs.eof then exit for
next
%>														
					  </ul>
						<div id="showCash2"></div>
					</div>
				</div>
				<script language="javascript" type="text/javascript">
				function getId(obj){return typeof(obj) == "string" ? document.getElementById(obj) : obj;}
				getId('showCash2').innerHTML=getId('showCash').innerHTML;
				var speed=40;function Marqueeg(){if(getId("showCash2").offsetTop-getId("showList").scrollTop<=0){getId("showList").scrollTop-=getId('showCash').offsetHeight;}else{getId("showList").scrollTop++;}}var MyMarG=setInterval(Marqueeg,speed);getId("showList").onmouseover=function(){clearInterval(MyMarG)};getId("showList").onmouseout=function(){MyMarG=setInterval(Marqueeg,speed)};</script>
			</div>
			<!--最新公告-->
		</div>
		<!--右边-->
		<div  class="clear"></div>
		<!--合作商家-->
		<div class="Hzsj">
			<h1>合作商家</h1>
			<ul class="Hzlist">
							<li ><img src="attach/event/192.jpg?1395308653" alt="闽乐游" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/191.jpg?1395308653" alt="蚂蚁游" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/190.jpg?1395308653" alt="小说阅读网" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/189.jpg?1395308653" alt="网易游戏" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/188.jpg?1395308653" alt="PPwan" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/187.jpg?1395308653" alt="到武林" width="106" height="50" border="0" /></li>					
								<li ><img src="attach/event/173.jpg?1395308653" alt="乐都网" width="106" height="50" border="0" /></li>					
								<li class="last"><img src="attach/event/139.jpg?1395308653" alt="盛大网络" width="106" height="50" border="0" /></li>					
							</ul>
			<h1 style="margin-bottom:0px;">友情链接:</h1>
			<ul class="Yqlist">
			<%=txt2%>
			</ul>
		</div>
		<!--合作商家-->
	</div>
	<!--主体-->
<!--#include file="bottom.asp"-->
