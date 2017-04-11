import java.sql.DriverManager;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.beust.jcommander.JCommander;
import com.beust.jcommander.ParameterException;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;


public class extract {

	public static void main(String[] args) {
		
		String server="";
		String username="";
		String password="";
		String directory="";
		try{		
			InputParameter params = new InputParameter();
		    new JCommander(params,args);

		    password=params.password;
			server=params.server;
			username=params.username;
			directory=params.directory;
		
		File[] path = new File(directory+"/").listFiles();

		for(File fileName:path) {
	         
            // prints file and directory paths
            System.out.println(fileName);
         
		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		try {
			
			BufferedReader in = new BufferedReader(new FileReader(fileName));
			String str;
			StringBuffer sb = new StringBuffer();
			while ((str = in.readLine()) != null) {
			sb.append(str + "\n ");
			}
			in.close();
			
			String connurl = ("jdbc:oracle:thin:@" + server);
			String jdbcurl = (connurl + "," + username + "," + password);

			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection c = DriverManager.getConnection(connurl, username, password);

			System.out.println("Opened database successfully");
			Statement stmt = c.createStatement(
					ResultSet.TYPE_SCROLL_INSENSITIVE,
					ResultSet.CONCUR_UPDATABLE);

			String MyQuery = sb.toString();

			ResultSet rs = stmt.executeQuery(MyQuery);
			ResultSetMetaData rsmd = rs.getMetaData();
			// Create a blank sheet
			XSSFSheet sheet = workbook.createSheet("sheet1");

			// GET METADATA FOR RESULTSET
			int columnCount = rsmd.getColumnCount();// GET COLUMN CONT( NUMBER
													// OF COLUMNS RETURNED BY
													// QUERY
			int firstrow = 1;// SET VERIABLE FOR FORMATING THE HTML TABLE
								// ACCORDING TO THE TABLE ROWS

			XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
			XSSFColor grey =new XSSFColor(new java.awt.Color(192,192,192));
			style.setFillForegroundColor(grey);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			style.setBorderTop(HSSFCellStyle.BORDER_THIN);
			style.setBorderRight(HSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		
			
			XSSFCellStyle border = (XSSFCellStyle) workbook.createCellStyle();
			border.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			border.setBorderTop(HSSFCellStyle.BORDER_THIN);
			border.setBorderRight(HSSFCellStyle.BORDER_THIN);
			border.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			
			Row header = sheet.createRow(0);

			for (int i = 1; i <= rsmd.getColumnCount(); i++) {
				// System.out.println(rsmd.getColumnLabel(i));
				Cell headerCell = header.createCell(i - 1);
				headerCell.setCellValue(rsmd.getColumnLabel(i));
				sheet.autoSizeColumn(i-1);
				headerCell.setCellStyle(style);
			}
			
			
			if (!rs.next()) { // IF NO RECORDS ARE RETURNED. THIS MOVES CURSOR
								// TO NEXT ROW
				System.out.println("Query returned no data/n");
			} else {
				// Iterate through the data in the result set and display it.
				rs.beforeFirst();
				while (rs.next()) {
					Row data = sheet.createRow(firstrow);
					// Print one row
					for (int i = 1; i <= columnCount; i++) {
						Cell dataCell = data.createCell(i - 1);
						dataCell.setCellValue(rs.getString(i));
						sheet.autoSizeColumn(i-1);
						dataCell.setCellStyle(border);
						// System.out.print(rs.getString(i) + " "); //Print row
					}
					firstrow++;
				}
				c.close();
			}
		} catch (NullPointerException e) {
			System.out
					.println(" is not updated because cells cannot be left null/empty ");
		}

		catch (Exception e) {
			System.err.println(e.getClass().getName() + ": " + e.getMessage());
			System.exit(0);
		}

		try {
			// Write the workbook in file system
			String fileNameWithoutExtension =fileName.toString().substring(0, fileName.toString().lastIndexOf("."));
			String fileNameWithoutfolder= fileNameWithoutExtension.substring(fileNameWithoutExtension.indexOf("\\"));
			Date dateobj = new Date();
			SimpleDateFormat timeStamp = new SimpleDateFormat("_ddMMyyyy");
			String date=timeStamp.format(dateobj);
			
			File  extractDir= new File("extracts"+date+"/"+fileNameWithoutfolder+date+".xlsx");
			if (extractDir.getParentFile() != null) {
				extractDir.getParentFile().mkdirs();
			}
			extractDir.createNewFile();
			
			
			FileOutputStream out = new FileOutputStream(extractDir);
			
			workbook.write(out);
			out.close();
			System.out.println(fileNameWithoutfolder+date+".xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
		} catch (ParameterException ex) {
		    System.out.println(ex.getMessage());
		    InputParameter usages = new InputParameter();
		    usages.usage();
	}
		
	}
		
}