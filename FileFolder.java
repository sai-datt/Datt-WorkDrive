package com.citi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileFolder {

	static Row row = null;
	static Sheet srcSheet = null;
	static Sheet destSheet = null;
	static FileOutputStream fos = null;
	static File srcFile = null;
	static File destFile = null;
	static XSSFWorkbook srcWorkbook = null;
	static XSSFWorkbook destWorkbook = null;
	static SXSSFWorkbook mainWorkbook = null;
	static int srcRowCount = 0;
	static int destRowCount = 0;
	static int cellCount = 0;
	static Map<Integer, Object[]> data = null;
	static boolean fileAppendStatus = true;

	public static void main(String[] args)
			throws InvalidFormatException, IOException {

		destFile = new File("G://testing/dest.xlsx");
		destWorkbook = new XSSFWorkbook(destFile);
		destSheet = destWorkbook.getSheetAt(0);
		mainWorkbook = new SXSSFWorkbook(destWorkbook);

		File folder = new File("G://CitiFolder");
		File[] fileList = folder.listFiles();
		data = new HashMap<Integer, Object[]>();
		long startTime = System.currentTimeMillis();
		for (File srcFile : fileList) {
			data = readDataFromFile(data, srcFile);
			System.out.println("File read successful!");
			fileAppendStatus = writeDataToFile(data);
			data.clear();
			System.out.println("File append successful!");
		}
		mainWorkbook.close();
		long endTime = System.currentTimeMillis();
		System.out.println(endTime - startTime);
		System.out.println("End of Operation!");
	}

	/**
	 * @param data
	 *            Map<Integer, Object[])
	 * @param srcFile
	 *            File
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private static Map<Integer, Object[]> readDataFromFile(
			Map<Integer, Object[]> data, File srcFile)
			throws IOException, InvalidFormatException {

		System.out.println(srcFile.getName());
		srcWorkbook = new XSSFWorkbook(srcFile);
		srcSheet = srcWorkbook.getSheetAt(0);

		srcRowCount = srcSheet.getPhysicalNumberOfRows();
		cellCount = srcSheet.getRow(0).getPhysicalNumberOfCells();

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
		System.out.println("Total rows read:" + srcRowCount);
		System.gc();
		return data;
	}

	/**
	 * @param data
	 *            Map<Integer, Object[]>
	 * @param count
	 *            Integer
	 * @return boolean
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private static boolean writeDataToFile(Map<Integer, Object[]> data)
			throws IOException, InvalidFormatException {

		destRowCount = destSheet.getPhysicalNumberOfRows();
		System.out.println("Existing row count in dest file:" + destRowCount);

		for (int rowNum = 0; rowNum < srcRowCount; rowNum++) {
			Row row = destSheet.createRow(destRowCount++);
			for (int cellNum = 0; cellNum < cellCount; cellNum++) {
				if (data.get(rowNum) != null) {
					Object[] objData = data.get(rowNum);
					for (Object obj : objData) {
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
		}
		try {
			fos = new FileOutputStream(destFile, true);
			mainWorkbook.write(fos);
			fos.close();
			return true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return false;

		} catch (IOException io) {
			return true;

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

}
