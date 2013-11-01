<!--#include file="oauth/_inc/_base.asp"-->

<%
	Dim blnLoggedIn : blnLoggedIn = False
	Dim strScreenName : strScreenName = Session(TWITTER_SCREEN_NAME)
	
	If Not IsNull(strScreenName) And strScreenName <> "" Then
		blnLoggedIn = True
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html> 
<head><title>Test</title>
        
        <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.2/jquery.min.js"></script>
    
		<link href="http://yui.yahooapis.com/3.1.0/build/cssreset/reset-min.css" rel="stylesheet" type="text/css">
		<link href="css/base.css" rel="stylesheet" type="text/css">
		<meta content="This page hosts an example of a Generic Classic ASP VBScript OAuth Library in action. The example uses Twitter's OAuth Authentication Flow to illustrate usage. The project in its entirety, with full source code, is available for download." name="description">
		<script type="text/javascript">
			var OAUTH_VBSCRIPT = {
				loggedIn: <%=LCase(CStr(blnLoggedIn))%>,
				screenName: '<%=strScreenName%>'
			};
            
    
</head> 
<body> 
            <div id="sign_out_container">
				<b>Logged in as:</b> <span id="screen_name"></span> [<a href="sign_out.asp" target="_blank" id="sign_out">sign out</a>]
				<br>
				<textarea id="tweet_textarea" name="tweet_textarea">test tweet</textarea>
				<br>
				<input type="button" value="post to twitter" id="tweet_tools_button" NAME="tweet_tools_button">
			</div>
			<div id="sign_in_container">
				<a href="twitter/authenticate.asp" target="_blank" id="sign_in">
					<img src="css/assets/Sign-in-with-Twitter-lighter.png" border=0></img>
				</a>
			</div>

			<script src="http://yui.yahooapis.com/3.1.0/build/yui/yui-min.js"></script>
			<script src="js/base.js"></script>

            <div id="fb-root"></div>
</body> 
</html>