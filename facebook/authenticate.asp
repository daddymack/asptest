<%option explicit%>
<%


    
            dim app_id 
            app_id = "174723912729791"
            dim app_secret
            app_secret = "a41ad1afedaf0fd99bb7a5a666947c3e"
            dim scope 
            scope = "publish_stream,manage_pages"
            
            
            if (Request.QueryString("code") = "") then
                Response.Redirect("https://graph.facebook.com/oauth/authorize?client_id=" & app_id & "&redirect_uri=" & Request.QueryString("redirect_uri") & "&scope=" & scope)
            else
                Response.Redirect("http://localhost:8383")
            end if

            
 %>
    