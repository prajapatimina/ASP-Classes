<% @ language="VBScript" %>
<% option explicit %>

<html>
    <head>
        <title>ASP introduction</title>
    </head>
    <body>
        <h1>Welcome To ASP Programming</h1>
        <h2>Raising Error</h2>
        <%
            dim firstName 
            firstName =""
            if len(firstName) < 3 or len(firstName) > 20 then
                Err.Raise 3001, "Validation", "Invalid first name."
            end if
        %>
        </body>
        </html>