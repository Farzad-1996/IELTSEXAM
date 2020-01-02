import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class AddServlet extends HttpServlet{
	
	public void doPost(HttpServletRequest req, HttpServletResponse res) throws IOException {
		
		int a = Integer.parseInt(req.getParameter("fnum"));
		int b = Integer.parseInt(req.getParameter("snum"));
		
		int sum = a+b;
		
		PrintWriter out = res.getWriter();
		out.println("The Sum is: "+ sum);
	}
}
