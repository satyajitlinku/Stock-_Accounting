package DriverFactory;

import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import commonfunctionslibrary.ERP_Functions;
import utilities.Excelfileutil;

public class ERP_DataDriven {
	ERP_Functions erp=new ERP_Functions();
	@BeforeTest
	public void launchapp(){
		String launch=erp.launchapp("http://webapp.qedge.com/login.php");
		System.out.println(launch);
		String login=erp.login("admin", "master");
		System.out.println(login);	
	}
	@Test
	public void verify_supplier() throws Exception{
		Excelfileutil xl=new Excelfileutil();
		int rc= xl.rowcount("supplier");
		int cc=xl.colcount("supplier");
		System.out.println("no of rows::"+rc+"  "+"no of columns ::"+cc);
		for (int i = 1; i <=rc; i++) {
			String sname=xl.getData("supplier", i, 0);
			String address=xl.getData("Supplier", i, 1);
			String city=xl.getData("Supplier", i, 2);
			String country=xl.getData("Supplier", i, 3);
			String cperson=xl.getData("Supplier", i, 4);
			String pnumber=xl.getData("Supplier", i, 5);
			String mail=xl.getData("Supplier", i, 6);
			String mnumber=xl.getData("Supplier", i, 7);
			String note=xl.getData("Supplier", i, 8);
			String results=erp.supplier(sname, address, city, country, cperson, pnumber, mail, mnumber, note);
			xl.setData("Supplier", i, 9, results);
			
		}
}


@AfterMethod
public void logout() throws Exception
{
	erp.logout();
			
		}
		
	}



