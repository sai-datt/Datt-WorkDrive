package com.citi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataTruncation_Phase2 {

	static List<Map<Integer, Object[]>> dataList = null;
	static Sheet sheet = null;
	static File destFile = null;
	static SXSSFWorkbook workbook = null;
	static Map<Integer, Object[]> data = null;

	public static void main(String[] args)
			throws InvalidFormatException, IOException {

		destFile = new File("G://testing/dest.xlsx");
		long startTime = System.currentTimeMillis();
		File folder = new File("G://Citi_1L_10k");
		File[] fileList = folder.listFiles();
		dataList = new ArrayList<Map<Integer, Object[]>>();
		for (File srcFile : fileList) {
			data = new HashMap<Integer, Object[]>();
			data = readDataFromFile(data, srcFile);
			dataList.add(data);
			System.out.println("Data list size:" + dataList.size());
		}
		System.out.println("File read successful!");

		writeDataToFile(dataList);

		long endTime = System.currentTimeMillis();

		System.out.println("Total time:" + (endTime - startTime));
	}
	

	/**
	 * This method reads data from source excel file
	 * 
	 * @param data Map<Integer, Object[]> 
	 * @param srcFile File
	 * @return Map<Integer, Object[]> 
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private static Map<Integer, Object[]> readDataFromFile(
			Map<Integer, Object[]> data, File srcFile)
			throws InvalidFormatException, IOException {

		Workbook srcWorkbook = new XSSFWorkbook(srcFile);
		Sheet srcSheet = srcWorkbook.getSheetAt(0);
		long srcRowCount = srcSheet.getPhysicalNumberOfRows();
		int cellCount = srcSheet.getRow(0).getPhysicalNumberOfCells();

		DataFormatter formatter = new DataFormatter();

		for (int rowNum = 0; rowNum < srcRowCount; rowNum++) {
			Object[] obj = new Object[cellCount];
			Row row = srcSheet.getRow(rowNum);
			if (row != null) {
				for (int cellNum = 0; cellNum < cellCount; cellNum++) {
					Cell cell = row.getCell(cellNum);
					if (cell != null) {
						obj[cellNum] = formatter.formatCellValue(cell);
					}
				}
				data.put(rowNum, obj);
			}
		}
		srcWorkbook.close();
		return data;
	}

	/**
	 * This method writes data to excel 
	 * @param dataList List<Map<Integer, Object[]>>
	 * @return boolean
	 */
	private static boolean writeDataToFile(
			List<Map<Integer, Object[]>> dataList) {

		workbook = new SXSSFWorkbook(500);
		sheet = workbook.createSheet();
		int rowNum = 0;

		for (int i = 0; i < dataList.size(); i++) {
			// System.out.println("Row num after file:" + rowNum);
			Map<Integer, Object[]> tempMap = dataList.get(i);
			System.out.println("Getting file data:" + i);

			for (Map.Entry<Integer, Object[]> entry : tempMap.entrySet()) {

				Row row = sheet.createRow(rowNum++);
				int cellNum = 0;
				Object[] objArr = entry.getValue();

				for (Object obj : objArr) {
					Cell cell = row.createCell(cellNum++);
					if (obj instanceof String) {
						cell.setCellValue((String) obj);
					} else if (obj instanceof Integer) {
						cell.setCellValue((Integer) obj);
					} else if (obj instanceof Date) {
						cell.setCellValue((Date) obj);
					} else if (obj instanceof Double) {
						cell.setCellValue((Double) obj);
					} else if (obj instanceof Long) {
						cell.setCellValue((Long) obj);
					}
				}
			}

		}
		OutputStream stream = null;
		try {
			System.out.println("Writing to file...");
			stream = new FileOutputStream(destFile);
			if (null != workbook && null != stream) {
				workbook.write(stream);// Write the data out
				stream.close();
				workbook.close();
				System.out.println("Writing over!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return true;
	}

}
