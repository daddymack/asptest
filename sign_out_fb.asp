<!--#include file="oauth/_inc/_base.asp"-->
<%
	Session(FACEBOOK_SCREEN_NAME) = ""
	Session.Contents.Remove(FACEBOOK_SCREEN_NAME)
	Session.Contents.RemoveAll()
	Session.Abandon()

	Response.ContentType = "text/javascript"
	Response.Write "<script></script>"
	Response.End
%>
