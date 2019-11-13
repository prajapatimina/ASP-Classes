<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2> Form inputs </h2>
        <form action="requestobjform.asp" method="post">
            <p>First Name : <input type="text" name="firstname">
            Last Name : <input type="text" name="lastname"></p>
            <p>Fav. Color :
                <select name="favColor">
                    <option>Blue</option>
                    <option>Yellow</option>
                    <option>Green</option>
                    <option>Pink</option>
                </select>
            </p>
            <P> Select your gender :
                <input type="radio" name="gender" value="Male">Male 
                                    <input type="radio" name="gender" value="Female">Female
                                    <input type="radio" name="gender" value="Others">Others
             </p>
            <input type="submit" value="Add User"/>
        </form>
        <%
            dim firstname,lastname,color,gender
            firstname = Request.Form("firstname")
            lastname = Request.Form("lastname")
            color = Request.Form("favColor")
            gender = Request.Form("gender")

        if firstname <> "" and lastname <> "" and color <> "" then
            if gender="Male" then
            Response.Write "Mr." & firstname & "" & lastname & "<br>"
            Response.Write "You chose" &  " " & color & ", great color" & "<br>"
            

            elseif gender="Female" then
            Response.Write "Miss" & firstname & "" & lastname & "<br>"
            Response.Write "You chose"  & "" & color & ", great color" & "<br>"

            else
             Response.Write "hi" & firstname & "" & lastname & "<br>"
            Response.Write "You chose" & color & ", great color" & "<br>"
            end if
         end if
        %>
        </body>
    </html>