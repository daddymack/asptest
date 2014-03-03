<%@ Language="VBScript"%>
<%option explicit%>
<%Session("Payment_Amount")=0.01 
response.Write "Total = " & Session("Payment_Amount")%>
<form action='expresscheckout.asp' METHOD='POST'>
<input type='image' name='submit' src='https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif' border='0' align='top' alt='Check out with PayPal'/>
</form>
