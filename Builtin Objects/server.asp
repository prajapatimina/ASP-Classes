<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP Server Object</h2>
        <%
            Server.scriptTimeout = 10
            Response.Write Server.scriptTimeout
             
        %>
        <hr>
        <%
            Response.Write "I am on server.asp"
            Server.Execute "requestobjform.asp"
            Response.Write "I am on the top" 
        %>

         <hr>
        <%
           ' Response.Write "I am on server.asp"
            'Server.Transfer "requestobjform.asp"
           ' Response.Write "I am on the top" 
        %>

        <hr/>
        <%
            Response.Write Server.MapPath("server.asp") & "<br>"
             Response.Write (Server.MapPath("respo.asp") & "<br>")
              Response.Write( Server.MapPath("\respo.asp") & "<br>")
               Response.Write (Server.MapPath("/") & "<br>")
                Response.Write (Server.MapPath("./") & "<br>")
        %>

        <hr>
        <%
            dim content
            content = "The image tag: <img> You & me, ""ok""?"
        %>
        </body>
    </html>