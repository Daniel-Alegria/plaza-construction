<%@ Language=vbScript%>

<%option explicit
Response.Buffer=true
Response.Expires=-1

%>

<% Dim i %>
<!-- #include virtual="/PartnerNet/Impe/configuration.inc" -->
<!-- #include file="../IMPE_Core.asp" --> 
<%
Dim errmsg,info,login,password,firstname,lastname,email,rsdb,EMailer
if request("your_email")<>"" and request("send")<>"" then
sSQL="Select top 1 user_email,user_fir_nam as firstname,user_lst_nam as lastname, Left(user_signin_id,(len(user_signin_id)-len(user_pwd))) as login ,user_pwd as password from user_access ua inner join user_det ud " &_
"on ud.user_id=ua.user_id where ud.user_email='" & request("your_email") & "'"
set rsdb = cn.execute(sSQL)
If  rsdb.eof and rsdb.bof then
	errmsg="Sorry your email does not appear to be listed, please contact support"
else
	Do while not rsdb.eof
	firstname=rsdb("firstname")
	lastname=rsdb("lastname")
	login=rsdb("login")
	password=rsdb("password")
	email=rsdb("user_email")
	rsdb.movenext
	loop
	
end if
set rsdb=nothing

if Errmsg = "" then 
		
		SET EMailer = Server.CreateObject("CDONTS.NewMail")								
							
		EMailer.ContentLocation = 0
			'Send ClickSafety Staff notification				
		EMailer.To = Email
		EMailer.From = "support@clicksafety.com"
		Emailer.CC="support@clicksafety.com"
		EMailer.Subject = "Your ClickSafety login information"
		EMailer.BodyFormat = 1
		EMailer.MailFormat = 0
		EMailer.Body = "As requested from the ClickSafety website here is your login information" & vbcrlf & vbcrlf &_
						"Name:" & firstname & " " & lastname & vbcrlf &_
						"username:" & login & vbcrlf &_
						"password:" & password & vbcrlf & vbcrlf &_
						"If you have any further questions please contact us at support@clicksafety.com" & vbcrlf &_
						"Thank you," & vbcrlf & vbcrlf & "ClickSafety Support"
		EMailer.Send
		Set EMailer = Nothing
	info="Your username and password has been sent and should arive shortly at the email address given."
	End if
elseif request("send")<>"" then
errmsg="Sorry your email does not appear to be listed, please contact support"
End if

%>

      <div align="left">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="755">
          <tr>
            <td height="8"><img border="0" src="images/trans.gif"></td>
          </tr>
        </table>
      </div>      
      <div align="left">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="755">
          <tr>
            <td colspan="3" bgcolor="#C0C0C0">
            <img border="0" src="images/trans.gif"></td>
          </tr>
          <tr>
            <td bgcolor="#C0C0C0" width="1">
            <img border="0" src="images/trans.gif"></td>
            <td bgcolor="#FFFFFF" width="753">
            <div align="left">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="753">
                <tr>
                  <td bgcolor="<%=LayotCellColor3%>" height="12">
                  <b><p class="tabletextcolhead">&nbsp;&nbsp;&nbsp;HELP CENTER:</p></b></td>
                </tr>
                <tr>
                  <td bgcolor="#C0C0C0"><img border="0" src="images/trans.gif"></td>
                </tr>
              </table>
            </div>
            <div align="left">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="753">
                <tr>
                  <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td width="20">&nbsp;</td>
                  <td width="713">
                  <div align="left">
                    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="710">
                      <tr>
                        <td width="488">
                        <p class="header">Help Topics</td>
                        <td width="22">&nbsp;</td>
                        <td width="200">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="488">
                        <div align="left">
                          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="488">
                            <tr>
                              <td>
                              <p class="header">Frequently Asked Questions:</td>
                            </tr>
                            <tr>
                              <td>
                              &nbsp;</td>
                            </tr>                            
                            <tr>
                              <td>
                              <p class="helptext">
                              <a href="#2">Forgot my username or password?</a><br>
                              <br>
                              <a href="#4">How do I register for training? </a>
                              <br>
                              <br>
                              <a href="#7">Do I have to finish the course in one 
                              sitting? </a><br>
                              <br>
                              <a href="#8">Can I train a group of employees 
                              using one login? </a><br>
                              <br>
                              <a href="#23">What are the minimum technical 
                              requirements?</a><br>
                              <br>
                              <a href="#25">When I try to log in I am getting a 
                              message that says, “Session expired”. </a><br>
                              &nbsp;</td>
                            </tr>
                            
                            <tr>
                              <td>
                              &nbsp;</td>
                            </tr>
                            <tr>
                              <td>
                              <p class="header"><a name="2"></a>Forgot my 
                              username or password?</td>
                            </tr>
                            <tr>
                              <td>
                              <p class="stdtext">Your default username is your 
                              first initial of your first name + your entire 
                              last name. Example: John Smith, username = JSmith<br>
                              <br>
                              Your default password is your last 4 digits of 
                              your Social Security Number. Example: 745-85-2586, 
                              password = 2586</td>
                            </tr>
                            <tr>
                              <td>
                              &nbsp;</td>
                            </tr>                            
							<tr>
								<td><p class="stdtext">If your email address is on-file with ClickSafety 
								you can request we email you your username and password 
								to you.</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td>
                        
								<form action="<%=HelpNavigationURL%>" method="post" name="email_check">
								<input type="text" value="<%=request("your_email")%>" name="your_email" size=40>
								<input type="submit" value="Submit" name="send">
						
								<%if errmsg<>"" then%>
								<p class="errcd">
								<%=errmsg%></p>
							<%elseif info <>"" then%>
				  
				  				<p class="errcd">
								 <%=info%></p>
							<%end if%>
							</form>
						
						</td>
                      </tr>
                            <tr>
                              <td>
                              <p class="header"><a name="4"></a>How do I register for training? </td>
                            </tr>
                            <tr>
                              <td>
                              <p class="stdtext">Once someone from your company 
                              creates a company account you can click on the “Register for Training” 
                              button on the home page. This will create a 
                              username for you and allow you to log in and take 
                              the course.</td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr>
                              <td>
                              <p class="header"><a name="7"></a>Do I have to 
                              finish the course in one session? </td>
                            </tr>
                            <tr>
                              <td>
                              <p class="stdtext">You don’t have to finish the 
                              course all at once. You may quit partway through 
                              the course, our system will bookmark where you 
                              left off.</td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr>
                              <td>
                              <p class="header"><a name="8"></a>Can I train a 
                              group of employees using one login? </td>
                            </tr>
                            <tr>
                              <td>
                              <p class="stdtext">No. Everyone has his or her own 
                              unique username and password. Everyone taking the 
                              course online will need to log in INDIVIDUALLY. 
                              The name on the printed certificate will be the 
                              name of the person logged into ClickSafety. </td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr>
                              <td>
                              <p class="header"><a name="23"></a>What are the 
                              minimum technical requirements?</td>
                            </tr>
                              <tr>
                              <td>
                              <p class="stdtext">Processor:<br>
                              1.0GHz (or equivalent) minimum
                              <br>
                              <br>
                              Operating System: <br>
                              Windows 7 or Mac OSX <br>
                              <br>
                              Memory: <br>
                              512 MB RAM<br>
                              <br>
                              Video Display: <br>
                              1024 x 768 Screen resolution <br>
                              <br>
                              Hardware Accessories: <br>
                              Sound card and speakers/headphones (recommended)
                              <br>
                              <br>
                              Internet Connection: <br>
                              Any broadband (DSL, Cable Modem, Satellite, T1, 
							  etc.)<br>
                              <br>
                              Internet Browser: <br>
                              Google Chrome or Internet Explorer 9.0 (IE9)<br>
							  <br>
                              <br>
                              Internet Browser Plug-in (For Online Training 
                              Courses): <br>
                              Adobe Flash 10.x Player™, Adobe Acrobat Reader 
							  10.x™</td>
                              </tr>
							  <tr>
                              <td>&nbsp;</td>
                              </tr>
							  <tr>
                              <td>
                              <p class="header"><a name="25"></a>When I try to 
                              log in I am getting a message that says, “Session 
                              expired”. </td>
                              </tr>
							  <tr>
                              <td>
                              <p class="stdtext">This problem is most likely due 
                              to your browser settings not accepting cookies. 
                              Our site requires cookies to keep track of your 
                              login information. It may also be that your 
                              administrator has deleted you out of the system.</td>
                              </tr>
                            <tr>
                              <td>
                              &nbsp;</td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            </table>
                        </div>
                        </td>                        
                        <td width="22">&nbsp;</td>
                        <td width="200" valign="top">
                        <div align="left">
                          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="200">
                            <tr>
                              <td colspan="3" bgcolor="#C0C0C0">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                            <tr>
                              <td bgcolor="#C0C0C0" width="1">
                              <img border="0" src="images/trans.gif"></td>
                              <td width="198">
                              <div align="left">
                                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="198">
                                  <tr>
                                    <td bgcolor="<%=LayotCellColor2%>" height="12">
										<p class="tabletextcolhead">&nbsp;&nbsp;&nbsp;SUPPORT:</p></td>
                                  </tr>
                                  <tr>
                                    <td bgcolor="#C0C0C0">
                                    <img border="0" src="images/trans.gif"></td>
                                  </tr>
                                </table>
                              </div>
                              <div align="left">
                                <table border="0" cellpadding="0" cellspacing="8" style="border-collapse: collapse" bordercolor="#111111" width="198">
                                  <tr>
                                    <td>
                                    <div align="left">
                                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                                        <tr>
                                          <td width="100%">
                                          <img border="0" src="images/csSupport.gif"></td>
                                        </tr>                                        
                                         <tr>
											<td>
												<p class="smallText">ClickSafety.com, Inc.<br>
          2185 N. California Blvd.<br>
          Suite 425<br>
		  Walnut Creek, CA 94596<br>
		  800.971.1080<br>
                   <a href="mailto:support@clicksafety.com?subject=CS/GA Help:">
           support@clicksafety.com</a></td>
</tr></td>
											</tr>
                                      </table>
                                    </div>
                                    </td>
                                  </tr>
                                </table>
                              </div>
                              </td>
                              <td bgcolor="#C0C0C0" width="1">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                            <tr>
                              <td colspan="3" bgcolor="#C0C0C0">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                            <tr>
                              <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr>
                              <td colspan="3" bgcolor="#C0C0C0">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                            <tr>
                              <td bgcolor="#C0C0C0" width="1">
                              <img border="0" src="images/trans.gif"></td>
                              <td width="198">
                              <div align="left">
                                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="198">
                                  <tr>
                                    <td bgcolor="<%=LayotCellColor2%>" height="12">
										<p class="tabletextcolhead">&nbsp;&nbsp;&nbsp;DOWNLOADS:</p></td>
                                  </tr>
                                  <tr>
                                    <td bgcolor="#C0C0C0">
                                    <img border="0" src="images/trans.gif"></td>
                                  </tr>
                                </table>
                              </div>
                              <div align="left">
                                <table border="0" cellpadding="0" cellspacing="8" style="border-collapse: collapse" bordercolor="#111111" width="198">
                                  <tr>
                                    <td width="100%">
                                    <div align="left">
                                      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                                        <tr>
                                          <td width="100%">
                                          <p class="stdtext">Adobe Flash 
                                          Player</td>
                                        </tr>
                                        <tr>
                                          <td width="100%">
                                          <a href="https://get.adobe.com/flashplayer/">
                                          <img border="0" src="images/buttons/get_flash_player.gif"></a></td>
                                        </tr>
                                        <tr>
                                          <td width="100%">&nbsp;</td>
                                        </tr>
                                        <tr>
                                          <td width="100%">
                                          <p class="stdtext">Adobe Acrobat 
                                          Reader</td>
                                        </tr>
                                        <tr>
                                          <td width="100%">
                                          <a href="https://get.adobe.com/reader/">
                                          <img border="0" src="images/buttons/get_adobe_reader.gif"></a></td>
                                        </tr>
                                        <tr>
                                          <td width="100%">&nbsp;</td>
                                        </tr>
                                      </table>
                                    </div>
                                    </td>
                                  </tr>
                                </table>
                              </div>
                              </td>
                              <td bgcolor="#C0C0C0" width="1">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                            <tr>
                              <td colspan="3" bgcolor="#C0C0C0">
                              <img border="0" src="images/trans.gif"></td>
                            </tr>
                          </table>
                        </div>
                        </td>
                      </tr>
                    </table>
                  </div>
                  <p>&nbsp;</td>
                  <td width="20">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="3">&nbsp;</td>
                </tr>
              </table>
            </div>
            </td>
            <td bgcolor="#C0C0C0" width="1">
            <img border="0" src="images/trans.gif"></td>
          </tr>
          <tr>
            <td colspan="3" bgcolor="#C0C0C0">
            <img border="0" src="images/trans.gif"></td>
          </tr>
        </table>
      </div>
      <div align="left">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="755">
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>
            <p class="smallText" align="center">&nbsp;© 2003 ClickSafety.com, 
            Inc., all rights reserved</td>
          </tr>
        </table>
      </div>
      