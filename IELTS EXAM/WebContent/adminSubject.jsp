<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Admin Login Form</title>
		<link rel="stylesheet" href="style.css">
    </head>
	<body background="bg.jpg">
		<ul>
		    <li><a href="logout.html">Logout</a></li>
		    <li><a href="#">Contact Us</a></li>
		    <li><a href="adminSubject.jsp" class="active">Subjects</a></li>
            <li><a href="#">About Us</a></li>
            <li><a href="adminPage.html">Home</a></li> 
        </ul>

       <div class="menuBox">
       <p><center><h1>Add Any Section Questions</h1></center></p>
			<a href="questionForm.jsp">
				<div id="listening">
				</div>
			</a>
			
			<a href="questionForm2.jsp">
				<div id="reading" >
				</div>
			</a>
			
			<a href="questionForm3.jsp">
				<div id="writing" >
				</div>
			</a>
			
			<a href="questionForm4.jsp">
				<div id="speaking" >
				</div>
			</a>
		</div>
     
	
	</body>
</html>