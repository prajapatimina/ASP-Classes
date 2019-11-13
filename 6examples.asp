<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>VBScript random examples using built-in APIs</h2>
        <h3>Datetime</h3>
        <p> 
        Today is:<%Response.Write date()%>
        <br/>
        And its <%Response.Write time()%>
        <br/>
        <%
        Response.Write(FormatDateTime(date(),vbgeneraldate))
        Response.Write("<br>")
        Response.Write(FormatDateTime(date(),vblongdate))
        Response.Write("<br>")
        Response.Write(FormatDateTime(date(),vbshortdate))
        Response.Write("<br>")
        Response.Write(FormatDateTime(now(),vblongtime))
        Response.Write("<br>")
        Response.Write(FormatDateTime(now(),vbshorttime))
        Response.Write("<br>")
        %>
        </p>
        <p>
            Date after 45 days from today:
            <% Response.Write DateAdd("d",45, date())%>     
            <br/>
             Date after 45 months from today:
            <% Response.Write DateAdd("m",45, date())%>  
            <br/>

             Date after 45 years from today:
            <% Response.Write DateAdd("yyyy",45, date())%>
            <br/> 
            <br/> 

            <%
                dim year2500
                year2500 = cdate("1/1/2500 00:00:00")

            %>
            It is <%  Response.Write datediff("yyyy", now(), year2500)%> years to year 2500!
            <br/>
            It is <%  Response.Write datediff("m", now(), year2500)%> months to year 2500!
            <br/>
            It is <%  Response.Write datediff("ww", now(), year2500)%> weeks to year 2500!
            <br/>
            It is <%  Response.Write datediff("d", now(), year2500)%> days to year 2500!
            <br/>
            It is <%  Response.Write datediff("h", now(), year2500)%> hours to year 2500!
            <br/>
            It is <%  Response.Write datediff("n", now(), year2500)%> mins to year 2500!
            <br/>

            <h3> Strings</h3>
            <%
            dim sometext
            sometext =" Madan Bhandari Memorial College "
            sometext = trim(sometext)                          'trim:-remove unnecessary 
            Response.Write strReverse(sometext)
            Response.Write "<br/>"
            Response.Write mid(sometext, 16, 8)
            %>

            <br/>
            <%
            randomize()
            Response.Write cint(100* rnd())
            %>
        </body>
    </html>