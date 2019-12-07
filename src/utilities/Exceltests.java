package utilities;

public class Exceltests {

	public static void main(String[] args) throws Exception {
Excelfileutil efu=new Excelfileutil();
		
		int totalRows=efu.rowcount("EmployeeDetails");
		System.out.println(totalRows);

		int totalColumns=efu.colcount("EmployeeDetails");
		System.out.println(totalColumns);
		
		String celldata=efu.getData("EmployeeDetails", 3, 2);
		System.out.println(celldata);
		efu.setData("EmployeeDetails", 3, 6, "FAIL");
		efu.setData("EmployeeDetails", 2, 6, "PASS");
		efu.setData("EmployeeDetails", 1, 6, "Not Executed");
		
		

	}

}
