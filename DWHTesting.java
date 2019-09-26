import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;

public class DWHTesting {
    private FileInputStream file;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private Connection conn;
    private Statement st;
    private String reportName;

    public void createTableInDB(String excelPath, int sheetNumber, int columnIndex) throws IOException, ClassNotFoundException, SQLException {
        file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(sheetNumber);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//Read column number 1; index start from 0;
            XSSFRow row = sheet.getRow(i);
            if (row.getCell(columnIndex).getCellType() == CellType.STRING)
                // System.out.println(((XSSFRow) row).getCell(0).getStringCellValue());
                st.executeQuery(((XSSFRow) row).getCell(columnIndex).getStringCellValue());
            System.out.println("Table: " + i + "- Created(0)");
            if (row.getCell(columnIndex).getCellType() == CellType.NUMERIC) {
                Double value = row.getCell(columnIndex).getNumericCellValue();
                Long lv = value.longValue();
                String data2 = lv.toString();
                st.executeQuery(data2);
                System.out.println("Table: " + i + "- Created(1)");
            }
        }
        file.close();
    }

    public void inserDataFromExcel(String excelPath, int sheetNumber, int columIndex) throws IOException, ClassNotFoundException, SQLException {
        file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(sheetNumber);
        reportName = workbook.getSheetAt(sheetNumber).getSheetName();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//Read column number 1; index start from 0;
            Row row = sheet.getRow(i);
            if (row.getCell(columIndex).getCellType() == CellType.STRING)
                // System.out.println(((XSSFRow) row).getCell(0).getStringCellValue());
                //  st.executeQuery("INSERT INTO "+ tableName+" ("+columnName+") Values ("+((XSSFRow) row).getCell(0).getStringCellValue()+")");
                st.executeQuery("INSERT INTO GLOBL.RED_SAS_MEMBER (MEMBER_ID) VALUES (" + (((XSSFRow) row).getCell(columIndex).getStringCellValue()) + ")");
//            System.out.println("0.Data Inserted"+i);
            if (row.getCell(columIndex).getCellType() == CellType.NUMERIC) {
                Double value = row.getCell(columIndex).getNumericCellValue();
                Long lv = value.longValue();
                String data2 = lv.toString();
                //  st.executeQuery("INSERT INTO "+tableName+" ("+") Values ("+data2+")");
                st.executeQuery("INSERT INTO GLOBL.RED_SAS_MEMBER (MEMBER_ID) VALUES (" + lv.toString() + ")");
//                System.out.println("1.Data Inserted"+i);
            }
        }
        file.close();
    }

    public void dropDBTable(String excelPath, int sheetNumber, int columnIndex) throws IOException, ClassNotFoundException, SQLException {
        file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(sheetNumber);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//Read column number 1; index start from 0;
            XSSFRow row = sheet.getRow(i);
            if (row.getCell(columnIndex).getCellType() == CellType.STRING)
                // System.out.println(((XSSFRow) row).getCell(0).getStringCellValue());
                st.executeQuery("DROP TABLE " + ((XSSFRow) row).getCell(columnIndex).getStringCellValue());
            System.out.println("Table: " + i + "- Dropped(0)");
            if (row.getCell(columnIndex).getCellType() == CellType.NUMERIC) {
                Double value = row.getCell(columnIndex).getNumericCellValue();
                Long lv = value.longValue();
                String data2 = lv.toString();
                st.executeQuery("DROP TABLE " + data2);
                System.out.println("Table: " + i + "- Dropped(1)");
            }
        }
        file.close();
    }

    public void connectDB(String driver, String dbURL, String user, String password) throws ClassNotFoundException, SQLException {
        conn = DriverManager.getConnection(dbURL, user, password);
        st = conn.createStatement();
    }

    public void displayDataOption2(int dbColumnIndex, String resultTextFile, String excelPath, int sheetNumber) throws SQLException, IOException {
        file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(sheetNumber);
        FileWriter file = new FileWriter(resultTextFile);
        PrintWriter pw = new PrintWriter(file);
        // ResultSet rs = st.executeQuery("Select * from medecon_ta_b");
        pw.println("Following Member_ID could not found or mismatched with DWH");
        pw.println("=========================================================");
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//Read from row number 1; index start from 0;
            Row row = sheet.getRow(i);
            if (row.getCell(2).getCellType() == CellType.STRING)//Read from Column number 2;
                System.out.println(((XSSFRow) row).getCell(2).getStringCellValue());
            ResultSet rs = st.executeQuery(((XSSFRow) row).getCell(2).getStringCellValue());
            while (rs.next()) {
                pw.println(rs.getString(dbColumnIndex));
            }
            if (row.getCell(2).getCellType() == CellType.NUMERIC) {
                Double value = row.getCell(2).getNumericCellValue();
                Long lv = value.longValue();
                String data2 = lv.toString();
                System.out.println(data2);
                rs = st.executeQuery(data2);
                while (rs.next()) {
                    pw.println(rs.getString(dbColumnIndex));
                }
            }
            rs.close();
        }
        file.close();
        pw.close();
    }

    public void displayDataOption0(int dbColumnIndex, String resultTextFile, String sqlQuery) throws SQLException, IOException {
        FileWriter file = new FileWriter(resultTextFile);
        PrintWriter pw = new PrintWriter(file);
        ResultSet rs = st.executeQuery(sqlQuery);
        String result = null;
        while (rs.next()) {
            pw.println(rs.getString(dbColumnIndex));
            result = rs.getString(dbColumnIndex);
        }
        if (result == null) {
            pw.println(reportName + " Report: Testing has PASSED..!");
        }
        if (result != null) {
            pw.println(reportName + " Report: Testing has FAILED: Oops above Data Could Not Found/Mismatched between Report and DWH..!");
            pw.println("=========================================================================================");
        }
        pw.close();
    }

    public void displayReturnResult(int dbColumnIndex, String sqlQuery, String resultPath) throws SQLException, IOException {
        FileWriter file = new FileWriter(resultPath+reportName+"-TestingResult"+".txt");
        PrintWriter pw = new PrintWriter(file);
        ResultSet rs = st.executeQuery(sqlQuery);
        String result = null;
        while (rs.next()) {
            pw.println(rs.getString(dbColumnIndex));
            result = rs.getString(dbColumnIndex);
        }
        if (result == null) {
            pw.println(reportName + " Report: Testing has PASSED, No Mismatched Found..!");
        }
        if (result != null) {
            pw.println("=========================================================================================");
            pw.println(reportName + " Report: Testing has FAILED: Oops above Data Could Not Found/Mismatched between Report and DWH..!");
        }
        pw.close();
    }
    //public void testResultStatus(String testStatus){
//    if (testStatus==null){pw.println("Testing has PASSED..!");}
//    if (testStatus!=null){
//        pw.println("Testing has FAILED: Oops following Data Could Not Found/Mismatched between Report and DWH..!");
//        pw.println("=========================================================================================");
//    }
//}
    public void last(String excelPath, int sheetNumber) throws IOException, ClassNotFoundException, SQLException {
        file = new FileInputStream(new File(excelPath));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheetAt(sheetNumber);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//Read column number 1; index start from 0;
            XSSFRow row = sheet.getRow(i);
            if (row.getCell(2).getCellType() == CellType.STRING)
                // System.out.println(((XSSFRow) row).getCell(0).getStringCellValue());
                st.executeQuery(((XSSFRow) row).getCell(2).getStringCellValue());
            System.out.println(((XSSFRow) row).getCell(2).getStringCellValue());
            if (row.getCell(2).getCellType() == CellType.NUMERIC) {
                Double value = row.getCell(2).getNumericCellValue();
                Long lv = value.longValue();
                String data2 = lv.toString();
                st.executeQuery(data2);
                System.out.println(data2);
            }
        }
        file.close();
    }

    public void closeDBConnection() throws SQLException {
        if (st != null) {
            st.close();
        }
        if (conn != null) {
            conn.close();
        }
        System.out.println("Execution has been COMPLETED..!");
    }
}