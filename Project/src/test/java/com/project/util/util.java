package com.project.util;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * 'Description : This class helps to read write values from the excel sheet
 * 'Author: 	VIK
 * '************************************************************************
 * ************************** ' C H A N G E H I S T O R Y
 * '************************************************************************
 * ************************** ' Date Change made by Purpose of change
 * '-------- -------------------
 * ------------------------------------------------- '
 * '************************************************************************
 * **************************
 */

public class util {
	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	public static XSSFWorkbook workbook;
	private static org.apache.poi.ss.usermodel.Cell Cell;
	private static XSSFRow Row;
	private static FileOutputStream fileOut = null;
	private static FileInputStream ExcelFile = null;
	private static final String TASKLIST = "tasklist";
	private static final String KILL = "taskkill /F /IM ";

	/*
	 * 'Description : Create the object of the excel file 'Author: VIK
	 */
	public static XSSFWorkbook setExcelFile(String Path) throws NoSuchMethodException {
		try {
			ExcelFile = new FileInputStream(Path);
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return ExcelWBook;
	}

	/*
	 * 'Description : Create the object of the excel file 'Author: VIK
	 */
	public static XSSFWorkbook setExcelFile1(String Path) throws NoSuchMethodException {
		try {
			ExcelFile = new FileInputStream(Path);
			workbook = new XSSFWorkbook(ExcelFile);
			return workbook;
		} catch (Exception e) {
			// Log.error("Class Utils | Method setExcelFile | Exception desc : " +
			// e.getMessage());
			try {
				try {
					// new DriverScript().bResult = false;
				} catch (SecurityException e1) {
					// Log.info(e1.getMessage());
				}
			} finally {
			}
			workbook = null;
			return workbook;
		} finally {
		}
	}

	/*
	 * 'Description : GetData from the excel sheet 'Author: VIK
	 */
	public static String getCellData(int RowNum, int ColNum, String SheetName) throws NoSuchMethodException {
		try {
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = Cell.getStringCellValue();
			return CellData;
		} catch (Exception e) {
			// Log.error("Class Utils | Method getCellData | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			}
			return "";
		} finally {
		}
	}

	/*
	 * 'Description : GetData from the excel sheet 'Author: VIK
	 */
	public static String getCellData(int RowNum, int ColNum, String SheetName, XSSFWorkbook workbook)
			throws NoSuchMethodException {
		try {
			ExcelWSheet = workbook.getSheet(SheetName);
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = Cell.getStringCellValue();
			return CellData;
		} catch (Exception e) {
			// Log.error("Class Utils | Method getCellData | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			}
			return "";
		} finally {
		}
	}

	/*
	 * 'Description : Total Row Count 'Author: VIK
	 */
	public static int getRowCount(String SheetName) throws NoSuchMethodException {
		int iNumber = 0;
		try {
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			iNumber = ExcelWSheet.getLastRowNum() + 1;
		} catch (Exception e) {
			// Log.error("Class Utils | Method getRowCount | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			} finally {
			}
		} finally {
		}
		return iNumber;
	}

	public static void killDrivers() throws Exception {
		String processName = "chromedriver.exe";
		String[] aprocess = processName.split("\\|");
		for (int i = 0; i <= aprocess.length - 1; i++) {
			if (isProcessRunning(aprocess[i])) {
				killProcess(aprocess[i]);
			}
		}
	}

	public static boolean isProcessRunning(String serviceName) throws IOException {

		Process p = Runtime.getRuntime().exec(TASKLIST);
		BufferedReader reader = new BufferedReader(new InputStreamReader(p.getInputStream()));
		String line;
		while ((line = reader.readLine()) != null) {
			if (line.contains(serviceName)) {
				return true;
			}
		}
		return false;
	}

	public static void killProcess(String serviceName) throws Exception {
		Runtime.getRuntime().exec(KILL + serviceName);
	}

	public static int getRowCount(String SheetName, XSSFWorkbook workbook) throws NoSuchMethodException {
		int iNumber = 0;
		try {
			ExcelWSheet = workbook.getSheet(SheetName);
			iNumber = ExcelWSheet.getLastRowNum() + 1;
		} catch (Exception e) {
			// Log.error("Class Utils | Method getRowCount | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			} finally {
			}
		} finally {
		}
		return iNumber;
	}

	public static int getColCount(String SheetName, int row) throws NoSuchMethodException {
		int iNumber = 0;
		try {
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			iNumber = ExcelWSheet.getRow(row).getLastCellNum();
		} catch (Exception e) {
			// Log.error("Class Utils | Method getRowCount | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			} finally {
			}
		} finally {
		}
		return iNumber;
	}

	public static int getRowContains(String sTestCaseName, int colNum, String SheetName) throws NoSuchMethodException {
		int iRowNum = 0;
		try {
			int rowCount = util.getRowCount(SheetName);
			for (; iRowNum < rowCount; iRowNum++) {
				if (util.getCellData(iRowNum, colNum, SheetName).equalsIgnoreCase(sTestCaseName)) {
					break;
				}
			}
		} catch (Exception e) {
			// Log.error("Class Utils | Method getRowContains | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			} finally {
			}
		} finally {
		}
		return iRowNum;
	}

	public static int getTestStepsCount(String SheetName, String sTestCaseID, int iTestCaseStart)
			throws NoSuchMethodException {
		try {
			for (int i = iTestCaseStart; i <= util.getRowCount(SheetName); i++) {
				if (!sTestCaseID.equals(util.getCellData(i, 1, SheetName))) {
					int number = i;
					return number;
				}
			}
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			int number = ExcelWSheet.getLastRowNum() + 1;
			return number;
		} catch (Exception e) {
			// Log.error("Class Utils | Method getRowContains | Exception desc : " +
			// e.getMessage());
			try {
				// new DriverScript().bResult = false;
			} catch (SecurityException e1) {
				// Log.info(e1.getMessage());
			}
			return 0;
		}
	}

	// @SuppressWarnings("static-access")
	public static void setCellData(String Result, int RowNum, int ColNum, String SheetName, String datasheetloc)
			throws Exception {
		try {
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			Row = ExcelWSheet.getRow(RowNum);
			Cell = Row.getCell(ColNum);
			if (Cell == null) {
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);
			} else {
				Cell.setCellValue(Result);
			}
			try {
				fileOut = new FileOutputStream(datasheetloc);
				XSSFFormulaEvaluator.evaluateAllFormulaCells(ExcelWBook);
				ExcelWBook.write(fileOut);
				ExcelWSheet = null;
				ExcelWBook = null;
				ExcelWBook = new XSSFWorkbook(new FileInputStream(datasheetloc));
			} catch (Exception e) {
				// TODO: handle exception
			} finally {
			}
		} catch (Exception e) {
			// new DriverScript().bResult = false;
		} finally {

		}

	}

	public static void setCellData1(String Result, int RowNum, int ColNum, String SheetName, String datasheetloc)
			throws Exception {
		try {
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			Row = ExcelWSheet.getRow(RowNum);
			Cell = Row.getCell(ColNum);
			if (Cell == null) {
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);
			} else {
				Cell.setCellValue(Result);
			}
			FileOutputStream fileOut = new FileOutputStream(datasheetloc);
			ExcelWBook.write(fileOut);
			ExcelWSheet = null;
			ExcelWBook = null;
			ExcelWBook = new XSSFWorkbook(new FileInputStream(datasheetloc));
		} catch (Exception e) {
			// new DriverScript().bResult = false;
		}
	}
}