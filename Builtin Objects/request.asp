<%@Language=VBScript %>
<% option explicit %>

 <%
            
            Response.Cookies("user")("firstname") = "Hari"
            Response.Cookies("user")("lastname") = "Anton"
            Response.Cookies("user")("age") = 23
            Response.Cookies("user")("isFemale") = false
 %>
<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Request Object</h2>
       
       <%
        dim cookie
        if Request.Cookies("user").Haskeys then
            for each cookie in Request.Cookies("user")
                Response.Write Request.Cookies("user")(cookie)
                Response.Write "<br/>"
            next
        end if
       %>
        <hr/>
       <a href="request.asp?name=ram">Click Me</a>
       <%
        dim name
        name= Request.QueryString("name")
        if name <> "" then
            Response.Write "Hey!" & name
        end if


        dim operand1,operand2, sum
        operand1 = Request.QueryString("op1")
        operand2 = Request.QueryString("op2")
        sum = cint(operand1) + cint(operand2)
       %>

       <%
        dim variable
        for each Variable in Request.ServerVariables
            Response.Write variable & "<br/>"
        next
       %>

       
       

       
        </body>
    </html>