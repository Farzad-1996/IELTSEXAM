<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<%@ page import="java.sql.*" %>
<!DOCTYPE HTML> 
<html> 
<head> 
<style> 
body{ 
    text-align: center; 
    background: #00ECB9; 
  font-family: sans-serif; 
  font-weight: 100; 
} 
h1{ 
  color: #396; 
  font-weight: 100; 
  font-size: 40px; 
  margin: 40px 0px 20px; 
} 
 #clockdiv{ 
    font-family: sans-serif; 
    color: #fff; 
    display: inline-block; 
    font-weight: 100; 
    text-align: center; 
    font-size: 30px; 
} 
#clockdiv > div{ 
    padding: 10px; 
    border-radius: 3px; 
    background: #00BF96; 
    display: inline-block; 
} 
#clockdiv div > span{ 
    padding: 15px; 
    border-radius: 3px; 
    background: #00816A; 
    display: inline-block; 
} 
smalltext{ 
    padding-top: 5px; 
    font-size: 16px; 
} 
</style> 
</head> 
<body> 
	<form action="answer2.jsp" method="get">
		<h1>Your Time Ends in</h1> 
		<h4 style="color:red;">You must submit your answers within time</h2>
		
		<div id="clockdiv"> 
		  <div> 
		    <span class="days" id="day"></span> 
		    <div class="smalltext">Days</div> 
		  </div> 
		  <div> 
		    <span class="hours" id="hour"></span> 
		    <div class="smalltext">Hours</div> 
		  </div> 
		  <div> 
		    <span class="minutes" id="minute"></span> 
		    <div class="smalltext">Minutes</div> 
		  </div> 
		  <div> 
		    <span class="seconds" id="second"></span> 
		    <div class="smalltext">Seconds</div> 
		  </div> 
		</div> 
		  
		<p id="demo"><b> Start Time </b></p> 
		  
		<script> 
		
		var ok = 0;
		
				var today, date, time1, startTime, minute, time, endTime, deadline;

				today = new Date();
				date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
				time1 = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
				startTime = date+' '+time1;
				
				if(ok == 0){
					minute = parseInt(today.getMinutes()) + 1;
				}
				else{
					minute = parseInt(today.getMinutes());
				}
				
				time = today.getHours() + ":" + minute.toString() + ":" + today.getSeconds();
				endTime = date+' '+time;
				
				document.getElementById("demo").innerHTML = startTime; 
				
				
				deadline = new Date(endTime).getTime();
			
			  
			var x = setInterval(function() { 
			  
			var now = new Date().getTime(); 
			var t = deadline - now; 
			var days = Math.floor(t / (1000 * 60 * 60 * 24)); 
			var hours = Math.floor((t%(1000 * 60 * 60 * 24))/(1000 * 60 * 60)); 
			var minutes = Math.floor((t % (1000 * 60 * 60)) / (1000 * 60)); 
			var seconds = Math.floor((t % (1000 * 60)) / 1000); 
			document.getElementById("day").innerHTML =days ; 
			document.getElementById("hour").innerHTML =hours; 
			document.getElementById("minute").innerHTML = minutes;  
			document.getElementById("second").innerHTML =seconds;  
			if (t < 0) { 
			        clearInterval(x); 
			        document.getElementById("demo").innerHTML = "TIME UP"; 
			        document.getElementById("day").innerHTML ='0'; 
			        document.getElementById("hour").innerHTML ='0'; 
			        document.getElementById("minute").innerHTML ='0' ;  
			        document.getElementById("second").innerHTML = '0'; 
			        }
				
			}, 1000); 
			
			setTimeout( function () { window.close(); }, 60000);
			setTimeout("window.open('answer2.jsp')", 60000);
			
			
			function go(){

				var confirmation = confirm("Do you want to submit your paper?");
				
				if(confirmation == true)
					{
					
						//setTimeout( function () { window.close(); });
						
						//window.open("answer.jsp");
						return true;
					}
				
				else
					{
						//document.write("Ok you can carry on your exam..")
						window.open("timer2.jsp");
						return false;
					}
					
			}
			
			
			
		</script>
		
		<%-- question papers start --%>
		
		<center><h1>Quiz Questions</h1></center>
    <p>
    
    		<%
    		Connection conn = null;
			  Statement st = null;
			  ResultSet rs = null;
    			try
    			{
    				  
    				  
    				  Class.forName("com.mysql.jdbc.Driver");
    				  conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
    				  st = conn.createStatement();
    				  
    			}catch(Exception e)
    				{
    					e.printStackTrace();
    				}
    		%>
    
			<b>Question 1.
			
			<%
			     String qry; 
			     qry = "";
			     
				 qry = "select * from `add_qstn2` where id = 1";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			        <label for="0"><%=rs.getString("question")%></label><br></b>
			        <blockquote>
					<p><input type="radio" name="radio0" value="1"/><label for="01"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio0" value="2"/><label for="02"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio0" value="3"/><label for="03"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio0" value="4"/><label for="04"><%=rs.getString("Op4")%></label></p>
			    
			    </blockquote>
			
			<p><b>
			<hr>
			Question 2.
			
			<%
				 qry = "select * from `add_qstn2` where id = 2";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="1"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio1" value="1"/><label for="11"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio1" value="2"/><label for="12"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio1" value="3"/><label for="13"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio1" value="4"/><label for="14"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 3.
			
			<%
				 qry = "select * from `add_qstn2` where id = 3";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="2"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio2" value="1"/><label for="21"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio2" value="2"/><label for="22"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio2" value="3"/><label for="23"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio2" value="4"/><label for="24"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 4.
			
			<%
				 qry = "select * from `add_qstn2` where id = 4";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="3"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio3" value="1"/><label for="31"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio3" value="2"/><label for="32"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio3" value="3"/><label for="33"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio3" value="4"/><label for="34"><%=rs.getString("Op4")%></label></p>
		
			</blockquote>
			<p><b>
			<hr>
			Question 5.
			
			<%
				 qry = "select * from `add_qstn2` where id = 5";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="4"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio4" value="1"/><label for="41"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio4" value="2"/><label for="42"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio4" value="3"/><label for="43"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio4" value="4"/><label for="44"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 6.
			
			<%
				 qry = "select * from `add_qstn2` where id = 6";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="5"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio5" value="1"/><label for="51"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio5" value="2"/><label for="52"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio5" value="3"/><label for="53"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio5" value="4"/><label for="54"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 7.
			
			<%
				 qry = "select * from `add_qstn2` where id = 7";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="6"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio6" value="1"/><label for="61"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio6" value="2"/><label for="62"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio6" value="3"/><label for="63"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio6" value="4"/><label for="64"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 8.
			
			<%
				 qry = "select * from `add_qstn2` where id = 8";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="7"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio7" value="1"/><label for="71"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio7" value="2"/><label for="72"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio7" value="3"/><label for="73"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio7" value="4"/><label for="74"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 9.
			
			<%
				 qry = "select * from `add_qstn2` where id = 9";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="8"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio8" value="1"/><label for="81"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio8" value="2"/><label for="82"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio8" value="3"/><label for="83"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio8" value="4"/><label for="84"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			Question 10.
			
			<%
				 qry = "select * from `add_qstn2` where id = 10";
			 	 rs = st.executeQuery(qry);
			  	 rs.next();
			%>
			
			<label for="9"><%=rs.getString("question")%></label><br></b>
			<blockquote>
					<p><input type="radio" name="radio9" value="1"/><label for="91"><%=rs.getString("Op1")%></label></p>
					<p><input type="radio" name="radio9" value="2"/><label for="92"><%=rs.getString("Op2")%></label></p>
					<p><input type="radio" name="radio9" value="3"/><label for="93"><%=rs.getString("Op3")%></label></p>
					<p><input type="radio" name="radio9" value="4"/><label for="94"><%=rs.getString("Op4")%></label></p>
					
			</blockquote>
			<p><b>
			<hr>
			
			

<%-- question papers end --%>
		<center><input type="submit" class="sb" value="Submit Answers" onclick="go()" style="height:50px; width:250px"></center>
			<br/><br/>
		
	</form>
</body> 
</html> 