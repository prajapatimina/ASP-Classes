<%@Language=VBScript %>
<% option explicit %>
<%
Response.Cookies("username")= "MBM College"

dim numberOfVisits
Response.Cookies("NumberOfVisits").Expires= date() +20
numberOfVisits = request.Cookies("NumberOfVisits")

if numberOfVisits = "" then
    Response.Cookies("NumberOfVisits")= 1
    Response.Write "Welcome! this is the first time you visited us"

else
    Response.Cookies("NumberOfVisits") = numberOfVisits 
    Response.Write "You have visited this web page & numberOfVisits & times before"
end if
%>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Response Object</h2>
        <%
           
            Response.Write "Hello ""World!"""
            Response.Write "<br/>"
            Response.Write "Hello" & chr(34) _
            & "World!" & chr(34)
            Response.Write "<br/>"
            Response.Write "Hello &quot;World!&quot;"
        %>
        <br/>
        <%
        'Response.Redirect "https://www.google.com"
        %>
        </body>
    </html>