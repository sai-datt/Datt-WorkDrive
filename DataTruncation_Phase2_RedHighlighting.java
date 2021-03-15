package com.citi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataTruncation_Phase2_RedHighlighting {

	static List<Map<Integer, Object[]>> dataList = null;
	static Sheet sheet = null;
	static File destFile = null;
	static SXSSFWorkbook workbook = null;
	static Map<Integer, Object[]> data = null;
	static CellStyle existingHeaderStyle = null;
	public static void main(String[] args)
			throws InvalidFormatException, IOException {

		destFile = new File("G://testing/dest.xlsx");
		long startTime = System.currentTimeMillis();
		File folder = new File("G://CitiFolder_100_2");
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
	 * @param data
	 *            Map<Integer, Object[]>
	 * @param srcFile
	 *            File
	 * @return Map<Integer, Object[]>
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private static Map<Integer, Object[]> readDataFromFile(
			Map<Integer, Object[]> data, File srcFile)
			throws InvalidFormatException, IOException {

		System.out.println("src file name:" + srcFile.getName());
		Workbook srcWorkbook = new XSSFWorkbook(srcFile);
		Sheet srcSheet = srcWorkbook.getSheetAt(0);
		existingHeaderStyle = srcSheet.getRow(0).getCell(0).getCellStyle();
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
	 * 
	 * @param dataList
	 *            List<Map<Integer, Object[]>>
	 * @return boolean
	 */
	private static boolean writeDataToFile(
			List<Map<Integer, Object[]>> dataList) {

		workbook = new SXSSFWorkbook(500);
		sheet = workbook.createSheet();
		int rowNum = 0;
		List<Object> srcList = null;
		List<Object> tgtList = null;

		// Creating style for header row
		CellStyle newHeaderStyle = workbook.createCellStyle();
		newHeaderStyle.cloneStyleFrom(existingHeaderStyle);

		// Creating style for target cells
		CellStyle targetCellStyle = workbook.createCellStyle();
		targetCellStyle.setFillForegroundColor(
				IndexedColors.GREY_25_PERCENT.getIndex());
		targetCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// Creating red highlighter font
		Font font = workbook.createFont();
		short index = HSSFColorPredefined.RED.getIndex();
		font.setColor(index);

		// Creating style for highlighting cells
		CellStyle highlightingCellStyle = workbook.createCellStyle();
		highlightingCellStyle.setFillForegroundColor(
				IndexedColors.GREY_25_PERCENT.getIndex());
		highlightingCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		highlightingCellStyle.setFont(font);

		CellStyle styleToUse = null;

		// Code to iterate the datalist, retrieve each datalist element and map
		// to new cell in destination workbook

		List<Object> headerCellList = Arrays.asList(dataList.get(0).get(0));

		// Loop to create header row
		int cellNum = 0;
		Row headerRow = sheet.createRow(rowNum++);
		for (Object headerCellObj : headerCellList) {
			styleToUse = newHeaderStyle;
			Cell cell = headerRow.createCell(cellNum++);
			if (headerCellObj instanceof String) {
				cell.setCellValue((String) headerCellObj);
				cell.setCellStyle(styleToUse);
			} else if (headerCellObj instanceof Integer) {
				cell.setCellValue((Integer) headerCellObj);
				cell.setCellStyle(styleToUse);
			} else if (headerCellObj instanceof Date) {
				cell.setCellValue((Date) headerCellObj);
				cell.setCellStyle(styleToUse);
			} else if (headerCellObj instanceof Double) {
				cell.setCellValue((Double) headerCellObj);
				cell.setCellStyle(styleToUse);
			} else if (headerCellObj instanceof Long) {
				cell.setCellValue((Long) headerCellObj);
				cell.setCellStyle(styleToUse);
			}
		}
		// End of loop to create header row

		for (int i = 0; i < dataList.size(); i++) { //
			// System.out.println("Row num after file:" + rowNum);

			// Each Map refers to a file where key represents rownum and value
			// represents cols in that row
			Map<Integer, Object[]> tempMap = dataList.get(i);
			System.out.println("Getting file data:" + i);

			srcList = new ArrayList<Object>();
			tgtList = new ArrayList<Object>();

			Object srcElementObj = null;
			Object tgtElementObj = null;

			// Code to loop between different map elements and writing to
			// workbook
			for (int s = 1, t = 2; s <= tempMap.size()
					&& t <= tempMap.size(); s = s + 2, t = t + 2) {

				srcList = Arrays.asList(tempMap.get(s));
				tgtList = Arrays.asList(tempMap.get(t));

				Row srcRow = null;
				cellNum = 0;

				// Writing srcList and tgtList to workbook
				if (srcList.contains("SRC") && !srcList.contains("TGT")
						&& !srcList.contains("RESULTRECORDID")
						&& !srcList.contains("COL_SRC")) {
					styleToUse = null;
					srcRow = sheet.createRow(rowNum++);
					for (Object obj : srcList) {
						Cell cell = srcRow.createCell(cellNum++);
						if (obj instanceof String) {
							cell.setCellValue((String) obj);
							cell.setCellStyle(styleToUse);
						} else if (obj instanceof Integer) {
							cell.setCellValue((Integer) obj);
							cell.setCellStyle(styleToUse);
						} else if (obj instanceof Date) {
							cell.setCellValue((Date) obj);
							cell.setCellStyle(styleToUse);
						} else if (obj instanceof Double) {
							cell.setCellValue((Double) obj);
							cell.setCellStyle(styleToUse);
						} else if (obj instanceof Long) {
							cell.setCellValue((Long) obj);
							cell.setCellStyle(styleToUse);
						}
					}
				}
				Row tgtRow = null;

				if (tgtList.contains("TGT")
						&& !tgtList.contains("RESULTRECORDID")
						&& !tgtList.contains("SRC")
						&& !tgtList.contains("COL_SRC")) {

					tgtRow = sheet.createRow(rowNum++);
					cellNum = 0;

					if (String.valueOf(srcList.get(2)) == String
							.valueOf(tgtList.get(2))
							&& (!srcList.contains("RESULTRECORDID")
									&& (!tgtList.contains("RESULTRECORDID")))) {

						System.out.println("src and tgt are related!");

						for (int listElement = 0; listElement < tgtList
								.size(); listElement++) {

							// cellNum = 0;
							srcElementObj = srcList.get(listElement);
							tgtElementObj = tgtList.get(listElement);

							styleToUse = targetCellStyle;

							// Verifying the elements having cellNum >2
							if (listElement > 2) {
								if (!srcElementObj.equals(tgtElementObj)) {
									System.out.println("Difference found at:"
											+ tgtElementObj);
									styleToUse = highlightingCellStyle;
								} else {
									styleToUse = targetCellStyle;
								}
							}

							// Writing tgtList to workbook
							Cell cell = tgtRow.createCell(cellNum++);
							if (tgtElementObj instanceof String) {
								cell.setCellValue((String) tgtElementObj);
								cell.setCellStyle(styleToUse);
							} else if (tgtElementObj instanceof Integer) {
								cell.setCellValue((Integer) tgtElementObj);
								cell.setCellStyle(styleToUse);
							} else if (tgtElementObj instanceof Date) {
								cell.setCellValue((Date) tgtElementObj);
								cell.setCellStyle(styleToUse);
							} else if (tgtElementObj instanceof Double) {
								cell.setCellValue((Double) tgtElementObj);
								cell.setCellStyle(styleToUse);
							} else if (tgtElementObj instanceof Long) {
								cell.setCellValue((Long) tgtElementObj);
								cell.setCellStyle(styleToUse);
							}
						}
					}
				}
				// End of writing srcList and tgtlist to workbook
			}
		}

		// Writing data to destination workbook
		OutputStream stream = null;
		try {
			System.out.println("Writing to file...");
			stream = new FileOutputStream(destFile);
			if (null != workbook && null != stream) {
				workbook.write(stream);
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
