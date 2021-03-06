<%@ Language="VBScript"%>
<!--#include file="oauth/_inc/_base.asp"-->
<!--#include file="facebook/utils.asp"-->
<%

	Dim blnLoggedIn : blnLoggedIn = False
    Dim blnLoggedIn_fb : blnLoggedIn_fb = False
	Dim strScreenName : strScreenName = Session(TWITTER_SCREEN_NAME)
	Dim strScreenName_fb : strScreenName_fb = Session(FACEBOOK_SCREEN_NAME)
	

	If Not IsNull(strScreenName) And strScreenName <> "" Then
		blnLoggedIn = True
	End If

    dim app_id 
    app_id = "174723912729791"
    dim app_secret
    app_secret = "a41ad1afedaf0fd99bb7a5a666947c3e"
    dim scope 
    scope = "publish_stream,manage_pages"
    dim curURl
    curURl = "http://localhost:8383/Default.asp"'curPageURL()
    if (Request.QueryString("code") <> "") then
        on error resume next
        dim HttpReq, apiURI
        set HttpReq=Server.CreateObject("MSXML2.ServerXMLHTTP")
        apiURI="https://graph.facebook.com/oauth/access_token?client_id=" & app_id & "&redirect_uri=" & curURl & "&scope=" & scope & "&code=" & Request.QueryString("code") & "&client_secret=" & app_secret
        
        Dim sResult : sResult = GetTextFromUrl(apiURI)
        if sResult <> "" then
            dim arrSplitted, sub_arrSplitted, access_token
            arrSplitted =split(sResult,"&")
        
            sub_arrSplitted=split(arrSplitted(0),"=")
            access_token=sub_arrSplitted(1)
            Session("access_token")=access_token
            apiURI="https://graph.facebook.com/me?access_token=" & access_token 
            
        
            sResult = GetTextFromUrl(apiURI)

        
            dim fbitems
            set fbitems=JSON.parse(sResult)

            blnLoggedIn_fb = True
            strScreenName_fb = fbitems.name
            Session(FACEBOOK_SCREEN_NAME) = strScreenName_fb
        end if
        on error goto 0
    end if
            
    If Not IsNull(strScreenName_fb) And strScreenName_fb <> "" Then
		blnLoggedIn_fb = True
	End If
    


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html> 
<head><title>Test</title>
        
        <script type="text/javascript" src="_private/js/jquery-1.10.2.min.js"></script>
    
		<link href="http://yui.yahooapis.com/3.1.0/build/cssreset/reset-min.css" rel="stylesheet" type="text/css">
		<link href="css/base.css" rel="stylesheet" type="text/css">
		<meta content="This page hosts an example of a Generic Classic ASP VBScript OAuth Library in action. The example uses Twitter's OAuth Authentication Flow to illustrate usage. The project in its entirety, with full source code, is available for download." name="description">
        <script language="JScript" runat="server" src='/_private/js/json2.js'></script>
        <script src="../_private/js/jquery.simplemodal.1.4.4.min.js"></script>
        <style>
            #simplemodal-container a.modalCloseImg {
	            background:url(../_private/js/x.png) no-repeat; /* adjust url as required */
	            width:25px;
	            height:29px;
	            display:inline;
	            z-index:3200;
	            position:absolute;
	            top:-15px;
	            right:-18px;
	            cursor:pointer;
            }
        </style>
		<script type="text/javascript">
			var OAUTH_VBSCRIPT = {
				loggedIn: <%=LCase(CStr(blnLoggedIn))%>,
				screenName: '<%=strScreenName%>'
			};
            var OAUTH_VBSCRIPT_FB = {
				loggedIn_fb: <%=LCase(CStr(blnLoggedIn_fb))%>,
				screenName_fb: '<%=strScreenName_fb%>'
			};
            function fbpost()
            {
                $.ajax({
                    type:"POST",
                    url: "facebook/post.asp",
                    dataType: "html",
                    data: "txtmsg=" + encodeURIComponent($("#txtmsg").val()) + "&txtlink=" + encodeURIComponent($("#txtlink").val()),
                    async: true,
                    success: function(msg)
                    {
                        alert(msg); 
                        $("#txtmsg").val('');
                    },
                    error: function (xhr, ajaxOptions, thrownError) 
                    {
                        alert(thrownError);
                        alert(xhr.status);
                    }
                });
    
            }
            function search_customer() {
                var src = "http://365.ericmmartin.com/";
                $.modal('<iframe src="' + src + '" height="450" width="830" style="border:0">', {
	                closeHTML:"",
	                containerCss:{
		                backgroundColor:"#fff", 
		                borderColor:"#fff", 
		                height:450, 
		                padding:0, 
		                width:830
	                },
	                overlayClose:true
                });
            }
        </script>
</head> 
<body> 
            

            <div id="sign_in_container_fb">
				<a href="facebook/authenticate.asp?redirect_uri=http://localhost:8383/Default.asp" target="" id="sign_in_fb">
					<img src="css/assets/facebook_signin.png" border=0></img>
				</a>
			</div>

            <div id="sign_out_container_fb">
				<b>Logged in as:</b> <span id="screen_name_fb"></span> [<a href="sign_out_fb.asp" target="_blank" id="sign_out_fb">sign out</a>]
                <br />
                Message: <br />
                <textarea id="txtmsg" name="txtmsg" rows="4" cols="130"></textarea>
                <br /> Link: <br />
                <textarea id="txtlink" name="txtlink" rows="1" cols="130">http://www.shopgoodwill.com/auctions/Detailed-Carved-Wood-Elephant-Stand-Decor-14658581.html</textarea>
                <br />
				<input type="button" name="fbpost" value="Post to Facebook" onclick="fbpost();"/>
                
			</div>
            
                

            <div id="sign_out_container">
				<b>Logged in as:</b> <span id="screen_name"></span> [<a href="sign_out.asp" target="_blank" id="sign_out">sign out</a>]
				<br> Message/Link: <br />
				<textarea id="tweet_textarea" name="tweet_textarea" rows="4" cols="130">http://www.shopgoodwill.com/auctions/Detailed-Carved-Wood-Elephant-Stand-Decor-14658581.html</textarea>
				<br>
				<input type="button" value="Post to Twitter" id="tweet_tools_button" NAME="tweet_tools_button">
			</div>
			<div id="sign_in_container">
				<a href="twitter/authenticate.asp" target="_blank" id="sign_in">
					<img src="css/assets/Sign-in-with-Twitter-lighter.png" border=0></img>
				</a>
			</div>
            <input type="button" value="Search Buyer" id="Search_Customer" name="Search_Customer" onclick="search_customer();"/>
			<script src="http://yui.yahooapis.com/3.1.0/build/yui/yui-min.js"></script>
			<script src="js/base.js"></script>
            
</body> 
</html>