<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1> 
        <h2>Loops</h2>
        <%
            'For...Next-uses a counter to runstatements a specified number of times.
             

            dim i
            for i= 0 to 10 step 2

            Response.Write "value of i: " & i
            Response.write "<br/>"
        next    
            


        %>

        <%
            dim j,k
            for j=1 to 6 
            for k=1 to j
                Response.write k
                
            next
            Response.write "<br/>"
            next


        'For Each....Next -Repeats a group of statements for each item in a collection or each element of 
        'an array
            dim fruits, item
            fruits= Array("Apple","Orange","Cherries")
            for each item in fruits
                Response.Write item & "<br/>"
            next
            %>

            <% 
            'do...loop - If you don't know how many repetitions you want 

            'Repeat code while a condition is true
                
            dim ii
            do while ii < 10
            ii = ii + 2
            Response.Write "value of ii: " & ii
            Response.write "<br/>"
        loop   
            

             dim iii
            do 
             iii = iii + 2
          
            Response.Write "value of iii: " & iii
            Response.write "<br/>"
        loop  while iii< 10 



        'Repeat code until a condition becomes true

        dim jj
        do until jj=5
            jj = jj + 1
            Response.Write "value of jj :" & jj
            Response.Write "<br/>" 
            loop


        dim jjj
        do until jjj=10
            jjj = jjj + 1
            if jjj= 3 then exit do
            Response.Write "value of jjj :" & jjj
            Response.Write "<br/>" 
            loop

            %>
        </body>
    </html>