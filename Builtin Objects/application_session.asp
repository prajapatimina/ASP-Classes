<%@ Language=VBScript %>
<% option explicit %>

<html>
	<head>
		<title>ASP introduction</title>
	</head>
	<body>
	<h1>Welcome To ASP Programming</h1>
	<h1>ASP Application and Session object</h1>

    <% 
    ' Application is collection of asp files and other files.
        Application("CollegeName") = "MBM, Anamnagar" 'Application variable declaration and defination, these variables can be accessed by all files in this application.
        Response.Write  Application.Contents("CollegeName") 'Contents contains all the application variables.
        Application.contents.remove("CollegeName") 'Contents.remove removes the application variables.
        Response.Write  Application.Contents("CollegeName") 'this wont be displayed.


    'Session is like local variable, each user has its own session variable,
        Session("info") = "Its a new Session"
        Response.Write session.contents("info") & "</br>"
    'Session.SessionID gives 
        Response.Write session.sessionID & "</br>"
    %>

    <hr>
    Your page request/Reloads: 
    <%

        Session("hits") = session("hits")+1
        Response.Write session("hits")


        Application.Lock
        Application("hits")= application("hits") + 1
        Application.Unlock
        Response.write "</br>Server request: " &application("hits")
        

    %>
    <%
        if request.form("DestorySession") = "Destory Session" then
            session.abandon
        end if
    %>
    <%  

        if(request.form("DestorySession") = "Destory Session") then 
            Response.write("</br>You are Disconnected.<br/>Session destoried.")
        else
            Response.write("</br>People Online: ")
            Response.write Application.Contents("noOfUserOnline")
        end if
    %>
    <form action="appandsession.asp" method="POST" >
        <input type="submit" value="Destory Session" name="DestorySession" /> 
    </form>

	</body>
</html>

