
<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
    
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>Login Here</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" rel='stylesheet'>
        <style>
            td{
                padding: 15px;
            }
            div{
                width: 51%;
                border: 1px solid black;
                border-radius: 5px;
                background-color: lightcoral;
            }
        </style>
    </head>
    <body background="bg5.jpg">
    <center> <h1 style="color:#D5BF03;">Login Here!</h1></center>
    <center>  <p> ${message}</p></center>
    <center>
        <div>
            <form action="login" method="post">
                <table>
                    <tr>
                        <td> User Name: </td>
                        <td> <input type="text"class='form-control' name = "uname" placeholder="User name"> </td>
                    </tr>                   
                    <tr>
                        <td> Email: </td>
                        <td> <input type="text"class='form-control' name = "email" placeholder="Email"> </td>
                    </tr>
                    <tr>
                        <td colspan="2" style ='text-align: center'> <input class="btn btn-success" type="submit" value = "submit"> </td>
                    </tr>
                </table>
            </form>
        </div>
    </center>
</body>
</html>
