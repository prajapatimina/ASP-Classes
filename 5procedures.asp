<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>VBScript Procedures</h2>
        <%
           'sub procedure:
           'series of statements, enclosed by the sub and End Sub statements can 
           'perform actions, but does not return a value

           

                sub WriteHello(helloString)
                Response.Write helloString
               end sub
             WriteHello "Hello World!" 

             'function:
             'is a series of statements,enclosed by the function and End Function statements 
             'can perform actions and can return a value returns a value by assigning a value to its name

             function add(a,b)
             add=a+b
             end function
             dim sum : sum= add(213, 343)
             Response.Write sum
             Response.Write "<br/>"



            'Area of circle

             function area(r)
                const pi=3.14
                area = pi* r*r
            end function
            dim circle : circle =area(3)
            Response.Write "Area of circle :" & circle
            Response.Write "<br/>"


            'print pattern

            sub pattern(a)
                 dim j,k
                for j=1 to 5
                    for k=1 to j
                Response.write a
                    next
                Response.Write "<br>"
             next    
            end sub
            pattern "*"


            
            
        %>
        </body>
    </html>