
package com.jnj.extracter.transform;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

@Service
public class Transform {

	private static final String MASTER_FILE = "excel/Detailed Costing.xlsx";
	private static final String JOURNAL_FILE = "excel/Journal.xlsx";
	private static final String OUTPUT_FILE = "excel/temp/Detailed Costing_transformed.xlsx";

	public String transformJournalWithMaster() {
		try (
			FileInputStream journalFis = new FileInputStream(new File(JOURNAL_FILE));
			Workbook journalWb = new XSSFWorkbook(journalFis);
			Workbook outputWb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream(new File(OUTPUT_FILE));
		) {
			// Load master data
			Map<String, Map<String, String>> masterData = loadMasterData(MASTER_FILE);
			Sheet journalSheet = journalWb.getSheetAt(0);
			Sheet outputSheet = outputWb.createSheet("Transformed");

			// 1. Read 3rd row for company/entity info
			Row thirdRow = journalSheet.getRow(2); // 0-based index
			if (thirdRow == null) {
				return "Journal file does not have a 3rd row.";
			}
			// Example: "7490 JJCM\t14.05.2025\tAUD\tPayrun No 11\t01.05.2025 - 31.05.2025"
			String companyEntity = "";
			String entity = "";
			String paidDate = "";
			String payrun = "";
			String period = "";
			if (thirdRow.getPhysicalNumberOfCells() >= 5) {
				companyEntity = thirdRow.getCell(0).toString().trim();
				entity = companyEntity.split(" ").length > 1 ? companyEntity.split(" ")[1] : "";
				paidDate = thirdRow.getCell(1).toString().trim();
				// third cell is currency, skip
				payrun = thirdRow.getCell(3).toString().trim();
				period = thirdRow.getCell(4).toString().trim();
			}

			// 2. Match with masterData
			Map<String, String> entityData = masterData.get(entity);
			if (entityData == null) {
				return "Entity '" + entity + "' not found in master data.";
			}

			// 3. Map rows below accordingly (starting from 4th row)
			int outputRowNum = 0;
			// Write header row (combine master and journal headers for demo)
			Row outputHeader = outputSheet.createRow(outputRowNum++);
			int col = 0;
			for (String key : entityData.keySet()) {
				outputHeader.createCell(col++).setCellValue(key);
			}
			outputHeader.createCell(col++).setCellValue("Paid Date");
			outputHeader.createCell(col++).setCellValue("Payrun");
			outputHeader.createCell(col++).setCellValue("Period");

			// Write entity data row
			Row outputEntityRow = outputSheet.createRow(outputRowNum++);
			col = 0;
			for (String key : entityData.keySet()) {
				outputEntityRow.createCell(col++).setCellValue(entityData.get(key));
			}
			outputEntityRow.createCell(col++).setCellValue(paidDate);
			outputEntityRow.createCell(col++).setCellValue(payrun);
			outputEntityRow.createCell(col++).setCellValue(period);

			// Copy the rest of the journal rows (from 4th row onwards)
			for (int i = 3; i <= journalSheet.getLastRowNum(); i++) {
				Row journalRow = journalSheet.getRow(i);
				if (journalRow == null) continue;
				Row outRow = outputSheet.createRow(outputRowNum++);
				for (int j = 0; j < journalRow.getLastCellNum(); j++) {
					Cell cell = journalRow.getCell(j);
					if (cell != null) {
						outRow.createCell(j).setCellValue(cell.toString());
					}
				}
			}

			// Save output
			outputWb.write(fos);
			return "Transformation completed. Output: " + OUTPUT_FILE;
		} catch (Exception e) {
			e.printStackTrace();
			return "Error: " + e.getMessage();
		}
	}


	public byte[] generateJournalFromMaster() {
		try (
			FileInputStream masterFis = new FileInputStream(new File(MASTER_FILE));
			Workbook masterWb = new XSSFWorkbook(masterFis);
			Workbook outputWb = new XSSFWorkbook();
			java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
		) {
			// Prepare styles
			// Data wrap style
			CellStyle wrapStyle = outputWb.createCellStyle();
			wrapStyle.setWrapText(true);

			// Info row style (yellow + wrap)
			CellStyle infoStyle = outputWb.createCellStyle();
			infoStyle.setWrapText(true);
			infoStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			infoStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// Header row style (blue + wrap + bold)
			CellStyle headerStyle = outputWb.createCellStyle();
			headerStyle.setWrapText(true);
			headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			Font headerFont = outputWb.createFont();
			headerFont.setBold(true);
			headerStyle.setFont(headerFont);
			Sheet masterSheet = masterWb.getSheetAt(0);
			Sheet outputSheet = outputWb.createSheet("Journal");

			// Define the exact headers as in your screenshot
			String[] journalHeaders = new String[] {
				"Account Code", "Account Name", "Doc Currency Amount", "Local Currency Amount", "Tax Code", "Calculated Tax", "Assignment", "Line Item Text", "Cost Centre", "Profit Centre"
			};

			// Prepare to read master data
			Iterator<Row> masterRows = masterSheet.iterator();
			Row masterHeader = masterRows.hasNext() ? masterRows.next() : null;
			List<String> headers = new ArrayList<>();
			if (masterHeader != null) {
				for (Cell cell : masterHeader) {
					headers.add(cell.getStringCellValue());
				}
			}

			int outputRowNum = 0;
			// Row 1: JOURNAL ENTRY 1 (bold, merged)
			Row row0 = outputSheet.createRow(outputRowNum++);
			row0.createCell(0).setCellValue("JOURNAL ENTRY 1");
			outputSheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, journalHeaders.length - 1));
			CellStyle boldStyle = outputWb.createCellStyle();
			Font font = outputWb.createFont();
			font.setBold(true);
			font.setFontHeightInPoints((short)14);
			boldStyle.setFont(font);
			row0.getCell(0).setCellStyle(boldStyle);

			// Row 2: blank
			outputSheet.createRow(outputRowNum++);

			// Use the first data row for the info row (row 3)
			Row firstDataRow = masterRows.hasNext() ? masterRows.next() : null;
			if (firstDataRow == null) {
				outputWb.write(baos);
				return new byte[0];
			}
			String company = getCellValueByHeader(firstDataRow, headers, "Company");
			String entity = getCellValueByHeader(firstDataRow, headers, "Entity");
			String paidDate = getCellValueByHeader(firstDataRow, headers, "Paid_Date");
			String currency = "AUD"; // Default/fixed, update if you have a column
			String payrun = getCellValueByHeader(firstDataRow, headers, "Pay_No");
			String period = getCellValueByHeader(firstDataRow, headers, "Date_Frm") + " - " + getCellValueByHeader(firstDataRow, headers, "Date_To");


			// Row 3: Company + Entity | Paid Date | Currency | Payrun | Period (highlighted)
			Row infoRow = outputSheet.createRow(outputRowNum++);
			infoRow.createCell(0).setCellValue((company + " " + entity).trim());
			infoRow.createCell(1).setCellValue(paidDate); // Paid Date as string
			infoRow.createCell(2).setCellValue(currency);
			infoRow.createCell(3).setCellValue(payrun);
			infoRow.createCell(4).setCellValue(period);
			for (int i = 0; i < 5; i++) {
				infoRow.getCell(i).setCellStyle(infoStyle);
			}

			// Row 4: Journal headers (highlighted)
			Row headerRow = outputSheet.createRow(outputRowNum++);
			for (int i = 0; i < journalHeaders.length; i++) {
				Cell cell = headerRow.createCell(i, CellType.STRING);
				cell.setCellValue(journalHeaders[i]);
				cell.setCellStyle(headerStyle);
			}

			// Data rows: for all rows in master, fill mapped columns
			// Mapping: Account Code=Account, Account Name=Emp_Name, Doc Currency Amount=Amount, Local Currency Amount=Amount, Tax Code=Txn_Code, Calculated Tax=Txn_Cat, Assignment=Txn_Grp, Line Item Text=Posn_Title, Cost Centre=Cost_Prd, Profit Centre=OU_Lvl_1
			if (firstDataRow != null) {
				List<Row> allRows = new ArrayList<>();
				allRows.add(firstDataRow);
				while (masterRows.hasNext()) {
					allRows.add(masterRows.next());
				}
				for (Row masterRow : allRows) {
					Row dataRow = outputSheet.createRow(outputRowNum++);
					String[] values = new String[] {
						getCellValueByHeader(masterRow, headers, "Account"), // Account Code
						getCellValueByHeader(masterRow, headers, "Account"), // Account Name
						getCellValueByHeader(masterRow, headers, "Amount"), // Doc Currency Amount
						getCellValueByHeader(masterRow, headers, "Amount"), // Local Currency Amount
						getCellValueByHeader(masterRow, headers, "Txn_Code"), // Tax Code
						getCellValueByHeader(masterRow, headers, "Txn_Cat"), // Calculated Tax
						getCellValueByHeader(masterRow, headers, "Txn_Grp"), // Assignment
						getCellValueByHeader(masterRow, headers, "Posn_Title"), // Line Item Text
						getCellValueByHeader(masterRow, headers, "Cost_Prd"), // Cost Centre
						getCellValueByHeader(masterRow, headers, "OU_Lvl_1") // Profit Centre
					};
					for (int i = 0; i < values.length; i++) {
						Cell cell = dataRow.createCell(i, CellType.STRING);
						cell.setCellValue(values[i] != null ? values[i] : "");
						cell.setCellStyle(wrapStyle);
					}
				}
			}

			// Auto-size all columns to fit the longest text
			for (int i = 0; i < journalHeaders.length; i++) {
				outputSheet.autoSizeColumn(i);
			}

			outputWb.write(baos);
			return baos.toByteArray();
		} catch (Exception e) {
			e.printStackTrace();
			return new byte[0];
		}
	}

	private Map<String, Map<String, String>> loadMasterData(String filePath) {
		Map<String, Map<String, String>> masterData = new HashMap<>();
		try (FileInputStream fis = new FileInputStream(new File(filePath));
			 Workbook wb = new XSSFWorkbook(fis)) {
			Sheet sheet = wb.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			Row headerRow = rowIterator.hasNext() ? rowIterator.next() : null;
			if (headerRow == null) return masterData;
			List<String> headers = new ArrayList<>();
			for (Cell cell : headerRow) {
				headers.add(cell.getStringCellValue());
			}
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Map<String, String> rowData = new HashMap<>();
				String entity = null;
				for (int i = 0; i < headers.size(); i++) {
					Cell cell = row.getCell(i);
					String value = cell != null ? cell.toString() : "";
					rowData.put(headers.get(i), value);
					if (headers.get(i).equalsIgnoreCase("Entity")) {
						entity = value;
					}
				}
				if (entity != null && !entity.isEmpty()) {
					masterData.put(entity, rowData);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return masterData;
	}

	private String getCellValueByHeader(Row row, List<String> headers, String headerName) {
		DataFormatter formatter = new DataFormatter();
		for (int i = 0; i < headers.size(); i++) {
			if (headers.get(i).equalsIgnoreCase(headerName)) {
				Cell cell = row.getCell(i);
				return cell != null ? formatter.formatCellValue(cell) : "";
			}
		}
		return "";
	}
		// Removed duplicate generateJournalFromMaster and getCellValueByHeader methods
}
