<% Option Explicit %>
<script language="javascript" runat="server" src="json2.js"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->

<%
'' JSON 2 Library from: 
''   https://github.com/nagaozen/asp-xtreme-evolution/tree/master/lib/axe/classes/Parsers
''

main

function main
	dim app_id
	dim app_secret
	dim my_url
	dim dialog_url
	dim token_url
	dim resp
	dim token
	dim expires
	dim graph_url
	dim json_str
	dim user
	dim code
	dim strLocation 
	dim strEducation
	dim strEmail
	dim strFirstName
	dim strLastName
	dim strID

    dim cookie 
    set cookie = get_facebook_cookie()
 
    do_login
	token = cookie("token")

	if token = "" then 
		response.write "Facebook login error"	
		exit function
	end if

	graph_url = "https://graph.facebook.com/me?access_token=" & token

	json_str = get_page_contents( graph_url )


	set user = JSON.parse( json_str )

	'' These properties should always be there provided
	'' we ask the right questions user.id & user.name
	strFirstName = user.first_name
	strLastName = user.last_name
	strID = user.id
	
	'' Handling properties that might not be there
	on error resume next
	strLocation = user.location.name
	strEducation = user.education.get(0).school.name
	strEMail = user.email
	strEmail = replace( strEmail, "\u0040", "@")

	on error goto 0

	response.write "USER ID: " & strID & "<br/>"
	response.write "First Name: " & strFirstName & "<br/>"
	response.write "Last Name: " & strLastName & "<br/>"

	response.write "Location: " & strLocation & "<br/>"
	response.write "Education: " & strEducation & "<br/>"
	response.write "Email: " & strEMail & "<br/>"
   
    response.write "<p/>"
    response.write "JSON String: <br/>"
   	response.write json_str
end function    


%>