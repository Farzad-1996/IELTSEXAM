<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
    
<%@page import="java.sql.DriverManager"%>
<%@page import="java.sql.ResultSet"%>
<%@page import="java.sql.Statement"%>
<%@page import="java.sql.Connection"%>

<%@ page import="java.sql.*" %>

<body background="bg6.jpg">

<%

int listen=0, read=0, write=0, speak=0;

String name = (String)session.getAttribute("uname");
String email = (String)session.getAttribute("email");

Connection conn = null;
Statement st = null;
ResultSet rs = null;

try {
	  Class.forName("com.mysql.jdbc.Driver");
	  conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
	  st = conn.createStatement();
	  
} catch (ClassNotFoundException e) {
e.printStackTrace();
}

%>
<h2 align="center" style="color:#D83F06;"><font><strong>Your Scores in respective Subjects</strong></font></h2>
<table align="center" cellpadding="10" cellspacing="15" border="3">
<tr>

</tr>
<tr bgcolor="#A52A2A">
<td><b>Subject</b></td>
<td><b>Score</b></td>
</tr>
<%
try{ 

String sql ="SELECT * FROM listening";
rs = st.executeQuery(sql);
boolean ok = false;
while(rs.next()){
	String temp = rs.getString("email");
	if(temp.equals(email))
	{
		rs.getString("email");
		ok = true;
		break;
	} 
}

if(ok)
{
	listen = Integer.parseInt(rs.getString("score"));
%>

<tr bgcolor="#DEB887">

<td>Listening</td>
<td><%=rs.getString("score") %></td>

</tr>

<%
	}
}
catch (Exception e) {
e.printStackTrace();
}
%>

<%-- reading part --%>
<%
try{ 

String sql ="SELECT * FROM reading";
rs = st.executeQuery(sql);
boolean ok = false;
while(rs.next()){
	String temp = rs.getString("email");
	if(temp.equals(email))
	{
		rs.getString("email");
		ok = true;
		break;
	} 
}

if(ok)
{
	read = Integer.parseInt(rs.getString("score"));
%>
<tr bgcolor="#DEB887">

<td>Reading</td>
<td><%=rs.getString("score") %></td>

</tr>

<%
	}
}
catch (Exception e) {
e.printStackTrace();
}
%>


<%-- writing part --%>
<%
try{ 

String sql ="SELECT * FROM writing";
rs = st.executeQuery(sql);
boolean ok = false;
while(rs.next()){
	String temp = rs.getString("email");
	if(temp.equals(email))
	{
		rs.getString("email");
		ok = true;
		break;
	} 
}

if(ok)
{
	write = Integer.parseInt(rs.getString("score"));
%>
<tr bgcolor="#DEB887">

<td>Writing</td>
<td><%=rs.getString("score") %></td>

</tr>

<%
	}
}
catch (Exception e) {
e.printStackTrace();
}
%>

<%-- Speaking part --%>
<%
try{ 

String sql ="SELECT * FROM speaking";
rs = st.executeQuery(sql);
boolean ok = false;
while(rs.next()){
	String temp = rs.getString("email");
	if(temp.equals(email))
	{
		rs.getString("email");
		ok = true;
		break;
	} 
}

if(ok)
{
	speak = Integer.parseInt(rs.getString("score"));
%>
<tr bgcolor="#DEB887">

<td>Speaking</td>
<td><%=rs.getString("score") %></td>

</tr>

<%
	}
}
catch (Exception e) {
e.printStackTrace();
}
%>

<%-- Band Score 
<tr bgcolor="#DEB887">

<td>Band Score(9)</td>
<%
	int total = listen + read + write + speak;
	double band = (total * 9) / 10;
	String sc = Double.toString(band);
%>
<td><%=sc %></td>

</tr>

--%>

</table>

</body>