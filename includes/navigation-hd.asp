              <tr> 
                <td valign="top" bgcolor="#D00707" width="5%"><img src="./images/hand.gif" border="0"></td>
                <td valign="middle" bgcolor="#D00707" class="txtfnt"><b><a href="index.asp" class="txtfnt">Home</a> 
                  <% if Role_ID = "3" then %>
				  | <a href="helpdesk_activity.asp" class="txtfnt">Daily Activity</a> 
				  <% end if	%>				  
				  <% if Role_ID = "4" then %>
				  | <a href="dashboard_spc.asp" class="txtfnt">My Dashboard</a> 
				  <% elseif	Role_ID = "3" then %>
				  | <a href="dashboard_user.asp" class="txtfnt">My Dashboard</a> 					 
				  <% end if	%>
				  | <a href="dashboard_user_all.asp" class="txtfnt">All Locations Dashboard</a> 
                  | <a href="dashboard_admin.asp" class="txtfnt">Location wise Dashboard</a> 
                  |</b></td>
                <td valign="middle" width="12%" bgcolor="#D00707" class="txtfnt" align="center"><b><a href="/logout.asp" class="txtfnt">Logout</a></b></td>
              </tr>
           