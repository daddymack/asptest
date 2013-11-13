<%@ Language="VBScript"%>
<%option explicit%>
<!--#include file="utils.asp"-->
<%
            
    dim app_id 
    app_id = "174723912729791"
    dim app_secret
    app_secret = "a41ad1afedaf0fd99bb7a5a666947c3e"
    dim scope 
    scope = "publish_stream,manage_pages"
    dim curURl
    curURl = "http://localhost:8383/Default.asp"'curPageURL()
            
    dim strURL, para, strLink
    'strLink = "http://www.shopgoodwill.com/auctions/Detailed-Carved-Wood-Elephant-Stand-Decor-14658581.html"
                
    strURL = "https://graph.facebook.com/me/feed"

    Dim str : str = Server.UrlEncode(request.form("txtmsg"))
    para = "?access_token=" & Session("access_token") & "&message=" & str '& "&link=" & strLink 
    Dim xmlHttp 
    Dim res

    set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP") 


    xmlHttp.Open "POST", strURL & para, false
    xmlHttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    xmlHttp.send
    response.write "'" & str & "' has been successfully posted to Facebook!"

 %>
