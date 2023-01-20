package com.actitime.genric;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

import java.io.*;

public class GetExcellData
{

        public static FileInputStream fis;
        public static FileOutputStream fos;
        public static XSSFWorkbook wb;
        public static XSSFSheet ws;
        public static XSSFRow row;
        public static XSSFCell cell;

        static FileInputStream FileInput() throws FileNotFoundException
        {
            FileLib f=new FileLib();
             fis= new FileInputStream(f.getExcellPath());
            return fis;

            }

    public static int getRowCount(String xlsheet) throws IOException
    {
            FileInput();
            wb=new XSSFWorkbook(fis);
            ws=wb.getSheet(xlsheet);
            int rowCount=ws.getLastRowNum();
            wb.close();
            fis.close();
            return rowCount;
        }

        public static int getCellCount(String xlsheet,int rownum) throws IOException
        {
            FileInput();
            wb=new XSSFWorkbook(fis);
            ws=wb.getSheet(xlsheet);
            row=ws.getRow(rownum);
            int cellCount=row.getLastCellNum();
            wb.close();
            fis.close();
            return cellCount;
        }

        public static String getcellData( String XlSheet, int i, int j) throws IOException
        {
            // TODO Auto-generated method stub
            FileInput();
            wb=new XSSFWorkbook(fis);
            String data = wb.getSheet(XlSheet).getRow(i).getCell(j).getStringCellValue();


            return data;
        }

     // @DataProvider(name="LoginData")
	public String[][] getExcellDataIn_2D_Array() throws IOException
	{
	 // String path1=System.getProperty("user.dir")+"/src/test/java/com/inetBanking/testData/New.xlsx";

	    int rownum=getRowCount("Sheet1");
	    int colcount=getCellCount("Sheet1",1);

	    String LoginData[][]=new String[rownum][colcount];

	    for(int i=1;i<=rownum;i++)
	    {

	    	for(int j=0;j<colcount;j++)
	    	{
	    		LoginData[i-1][j]=getcellData("Sheet1",i,j);

	    	}
	    }
	    return LoginData;
	}

    public void getPrintExcellTable() throws IOException {
        String ar[][]= getExcellDataIn_2D_Array();
        for(int i=0; i<ar.length;i++)
        {
            for(int j=0;j<ar[0].length;j++)
            {
                System.out.print(ar[i][j]+" ");
            }
            System.out.println();
        }
    }


        public static void main(String[] args) throws IOException
        {
           GetExcellData e=new GetExcellData();
           int row= getRowCount("Sheet1");
           int cell= getCellCount("Sheet1",1);
           String data=getcellData("Sheet1",3,1);
            e.getPrintExcellTable();
            System.out.println(data);
            System.out.println(row);
            System.out.println(cell);

    }

}
