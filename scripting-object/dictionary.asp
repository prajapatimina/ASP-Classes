<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Dictionaary objects</h2>
        <%
           dim countryCapitals ,country
           set countryCapitals = server.CreateObject("Scripting.Dictionary")

           countryCapitals.add "Nepal", "Kathmandu"
            countryCapitals.add "India", "New Delhi"
             countryCapitals.add "Russia", "Mosco"
              countryCapitals.add "Germany", "Berlin"

            for each  country in countryCapitals
                response.write countryCapitals(country) &_
                " is capital city of " & country & ".<br/>"

            next
            countryCapitals.Items

        %>
        </body>
    </html>