<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>

    <%
        dim userdetails
        set userdetails = server.CreateObject("Scripting.Dictionary")
        userdetails.add "mina" ,"12345"
    %>
        <h1>Welcome to ASP Dictionary Object</h1>
        <h2>Login Form  </h2>
        <form action="LoginValidation.asp" method="post">
            <p>User Name : <input type="text" name="username">
            
           <p> Password : <input type= "password" name="password"></p>
           
            <input type="submit" value="Login" name="login"/>
        </form>

    <%
    dim u , p

    if request.form("login")="Login" then
        u=request.form("username")
        p= request.form("password")

        if userdetails(u)= p then
            response.write "Logged in"
            response.write "welcome"
            response.end
        else
            response.write "Something wrong"
        end if
    end if
      

     %>    
        </body>
    </html>