<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>control structure</h2>

        
        <%
        '  dim i : i=45
        '    if i=45 then
         '       Response.Write hour(time)
        '    end if
        %>

        <%
            dim i : i=hour(time)
            if i=10 then
                Response.Write "class started..."
            elseif i=12 then
                Response.Write"Afternoon, lunch time..."
            elseif i=3 then
                Response.Write"Time to go home.."
            else
                Response.Write"whatever.."
            end if

            'select cas
            
            dim hairColor
            hairColor ="Yellow"
            select case hairColor
            case "Black"
                Response.Write("Smae as the cat")
            case "Blue"
                Response.Write("Same as the parrot")
            case "Yellow"
                Response.Write("Same as the taxi")
            case else
                Response.Write("Ah well,whatever...")
            end select
        %>
        </body>
    </html>