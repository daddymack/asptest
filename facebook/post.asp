<%option explicit%>
<!--#include file="utils.asp"-->
<%


            dim strMessage
            strMessage = Request.Form("msg")
            response.write "strMessage:" & strMessage

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
                if sResult <> "" then
                    dim arrSplitted, sub_arrSplitted, access_token
                    arrSplitted =split(sResult,"&")
                    sub_arrSplitted=split(arrSplitted(0),"=")
                    access_token=sub_arrSplitted(1)
                
                    dim strURL, para, strLink
                    strLink = "http://www.shopgoodwill.com/auctions/Detailed-Carved-Wood-Elephant-Stand-Decor-14658581.html"
                
                    strURL = "https://graph.facebook.com/me/feed"
                    
                    
                    'para = "?access_token=" & access_token & "&message=" & strMessage & "&link=" & strLink 
                    para = "?access_token=" & access_token & "&message=" & strMessage
                    Dim xmlHttp 
                    Dim res

                    set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP") 


                    xmlHttp.Open "POST", strURL & para, false
                    xmlHttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
                    xmlHttp.send
                    response.write "Successfully posted to facebook"
                    
                end if
                
            end if
 %>
