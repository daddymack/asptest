<%option explicit%>
<%
function curPageURL()
 dim s, protocol, port

 if Request.ServerVariables("HTTPS") = "on" then 
   s = "s"
 else 
   s = ""
 end if  
 
 protocol = strleft(LCase(Request.ServerVariables("SERVER_PROTOCOL")), "/") & s 

 if Request.ServerVariables("SERVER_PORT") = "80" then
   port = ""
 else
   port = ":" & Request.ServerVariables("SERVER_PORT")
 end if  

 curPageURL = protocol & "://" & Request.ServerVariables("SERVER_NAME") &_ 
              port & Request.ServerVariables("SCRIPT_NAME")
end function

function strLeft(str1,str2)
 strLeft = Left(str1,InStr(str1,str2)-1)
end function
Function GetTextFromUrl(url)
  Dim oXMLHTTP
 
  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

  oXMLHTTP.Open "GET", url, False
  oXMLHTTP.Send

  If oXMLHTTP.Status = 200 Then

    GetTextFromUrl = oXMLHTTP.responseText

  End If
end function

    
            dim app_id 
            app_id = "174723912729791"
            dim app_secret
            app_secret = "a41ad1afedaf0fd99bb7a5a666947c3e"
            dim scope 
            scope = "publish_stream,manage_pages"
            dim curURl
            curURl = curPageURL()
            
            if (Request.QueryString("code") = "") then
                Response.Redirect("https://graph.facebook.com/oauth/authorize?client_id=" & app_id & "&redirect_uri=" & curURl & "&scope=" & scope)
            else
                dim HttpReq, apiURI
                set HttpReq=Server.CreateObject("MSXML2.ServerXMLHTTP")
                apiURI="https://graph.facebook.com/oauth/access_token?client_id=" & app_id & "&redirect_uri=" & curURl & "&scope=" & scope & "&code=" & Request.QueryString("code") & "&client_secret=" & app_secret
            
                Dim sResult : sResult = GetTextFromUrl(apiURI)
                
                dim arrSplitted, sub_arrSplitted, access_token
                arrSplitted =split(sResult,"&")
                sub_arrSplitted=split(arrSplitted(0),"=")
                access_token=sub_arrSplitted(1)
                
                dim strURL, strMessage, para, strLink
                strLink = "http://www.shopgoodwill.com/auctions/Detailed-Carved-Wood-Elephant-Stand-Decor-14658581.html"
                
                strURL = "https://graph.facebook.com/me/feed"
                strMessage = "Facebook posting test"
                para = "?access_token=" & access_token & "&message=" & strMessage & "&link=" & strLink

                Dim xmlHttp 
                Dim res

                set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP") 


                xmlHttp.Open "POST", strURL & para, false
                xmlHttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
                xmlHttp.send

                response.write "Successfully posted to facebook"
            end if
 %>
