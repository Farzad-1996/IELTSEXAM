<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Admin Login Form</title>
		<link rel="stylesheet" href="style.css">
    </head>
	<body background="subback.jpg">
		<ul>
		    <li><a href="logout.html">Logout</a></li>
		    <li><a href="#">Contact Us</a></li>
		    <li><a href="adminSubject.jsp" class="active">Subjects</a></li>
            <li><a href="#">About Us</a></li>
            <li><a href="#home" >Home</a></li> 
        </ul>

       <div class="menuBox1">
			<center><h1><font color="#680b87">Online Examination</font></h1></center>
    <p>
    <form name="quiz" action="adminQstnInDatabase3.jsp" method="get">
    
			<b><font color="#a3087a">Question 1.</font>
			        <br><input type = "text" name="0" value = "" placeholder="Enter your qstn here"/><br></b>
			        <blockquote>
					<p><font color="blue">1)</font> <input type="text" name="01" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="02" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="03" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="04" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="05" placeholder="answer"/></p>
			    
			    </blockquote>
			
			<p><b>
			<hr>
			<font color="#a3087a">Question 2.</font>
			<br><input type = "text" name="1" value = "" placeholder="Enter your qstn here"/><br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="11" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="12" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="13" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="14" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="15" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 3.</font>
			<br><input type = "text" name="2" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="21" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="22" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="23" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="24" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="25" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 4.</font>
			<br><input type = "text" name="3" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="31" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="32" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="33" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="34" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="35" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 5.</font>
			<br><input type = "text" name="4" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="41" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="42" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="43" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="44" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="45" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 6.</font>
			<br><input type = "text" name="5" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="51" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="52" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="53" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="54" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="55" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<font color="#a3087a">Question 7.</font>
			<br><input type = "text" name="6" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="61" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="62" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="63" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="64" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="65" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 8.</font>
			<br><input type = "text" name="7" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="71" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="72" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="73" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="74" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="75" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 9.</font>
			<br><input type = "text" name="8" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="81" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="82" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="83" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="84" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="85" placeholder="answer"/></p>
			</blockquote>
			<p><b>
			<hr>
			<font color="#a3087a">Question 10.</font>
			<br><input type = "text" name="9" value = "" placeholder="Enter your qstn here"/></br></b>
			<blockquote>
					<p><font color="blue">1)</font> <input type="text" name="91" placeholder="option-1"/></p>
					<p><font color="green">2)</font> <input type="text" name="92" placeholder="option-2"/></p>
					<p><font color="purple">3)</font> <input type="text" name="93" placeholder="option-3"/></p>
					<p><font color="#ba4506">4)</font> <input type="text" name="94" placeholder="option-4"/></p>
					<p><b><font color="#7c07a3">(**)Set Answer</font></b> <input type="text" name="95" placeholder="answer"/></p>
			</blockquote>
			<br/>
			
			<center><input type="submit" class="sb" value="Submit Questions" style="height:50px; width:250px"></center>
			<br/><br/>

	</form>

	</div>
     
	
	</body>
</html>