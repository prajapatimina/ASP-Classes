<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Dictionaary objects</h2>
        <form action="" method="POST">

        
        <input type="text" name="username">
        
        <input type="password" name="password">
        
        <input type="submit" name="login" value="login">
        
        </form>
        
        
        <%

        dim logindic
        set logindic=server.createObject("scripting.dictionary")
        logindic.add "username","mina"
        logindic.add "password","123"
        if request.form("login")="login" then
            dim u,p
            u=Request.form("username")
            p=Request.form("password")

            if u= logindic.Item("username") and p=logindic.Item("password") then    
                response.redirect ("welcome.asp")
             else
                response.write ("Invalid")

              end if  



              end if  






        %>
        </body>
    </html>