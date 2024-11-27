package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.opencsv.CSVWriter;

public class ReadXlsFile {
	
	private static String fileName = "";

	@SuppressWarnings("deprecation")
	public static void main(String[] args) {
		
		// Path to the .xls file
		String filePath ="C://Users/test/OneDrive - The Outsource Pro/Desktop/DeliverablesMehul/The Dammous Family Farm Trust/The Dammous Family Farm Trust - MYOB AE - FA Schedule (2023).xls";
		
		
		String[] strArr = filePath.split("/");
		fileName = strArr[strArr.length - 1];
		readXLSFile(filePath);
	}
	
	private static void readXLSFile(String filePath) {
		
		try (FileInputStream fis = new FileInputStream(new File(filePath)); Workbook workbook = new HSSFWorkbook(fis)) {

			// Get the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			int empCellCount = 0;
			int rowCount = 0;
			Set<AssetData> assetList = new HashSet<>();
			
			// Iterate through each row
			AssetData assetData = null;
			int headerRowNum = 10;
			String regex = "\\d{3}";
			boolean flag = false;
			for (Row row : sheet) {
				if (row.getRowNum() > 10) {
					Cell empCell = row.getCell(0);
					empCellCount = empCell == null ? empCellCount+1 : 0;
					if(empCellCount > 2 || flag) {
						Cell checkCell = sheet.getRow(row.getRowNum()+1).getCell(0);
						if(checkCell != null && !checkCell.getStringCellValue().equalsIgnoreCase("total") && checkCell.getStringCellValue().matches(regex)) {
							headerRowNum = row.getRowNum()+1;
							flag = false;
							continue;
						}
						if(checkCell != null && checkCell.getStringCellValue().equalsIgnoreCase("total")) {
							flag = false;
							break;
						}
						if(checkCell != null && !checkCell.getStringCellValue().matches(regex)) {
							flag = true;
							continue;
						}
						if(checkCell == null) {
							continue;
						}
					}
					if(headerRowNum == row.getRowNum()) {
						flag = false;
						continue;
					}
					int cellCount = 0;
					for (Cell cell : row) {
						switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            break;
                        default:
                        	System.out.print(cell+ " -> " +rowCount+"-"+cellCount+ "\t");
						if(rowCount == 0 && cellCount == 0) {
							assetData = new AssetData();
							assetData.setAssetType(sheet.getRow(headerRowNum).getCell(3).getStringCellValue());
							assetData.setAssetNumber(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
						}
						if(rowCount == 0 && cellCount == 1) {
							assetData.setAssetName(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
						}
						if(rowCount == 1 && cellCount == 0) {
							assetData.setPurchaseDate(cell);
						}
						if(rowCount == 1 && cellCount == 1) {
							assetData.setPurchasePrice(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
						}
						if(rowCount == 1 && cellCount == 4) {
							assetData.setBookRate(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
						}
						if(rowCount == 1 && cellCount == 7) {
							assetData.setClosingBookRate(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
							assetData.setBookAccumulatedDepreciation(Double.parseDouble(assetData.getPurchasePrice().toString()) - 
									Double.parseDouble(assetData.getClosingBookRate().toString()));
						}
						cellCount++;
						}
					}
					flag = false;
					assetList.add(assetData);
					rowCount++;
					if(row.getCell(0) == null) {
						rowCount = 0;
					}
					System.out.println();
				}
				
			}
			writeCSVFile(assetList);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	private static void writeCSVFile(Set<AssetData> assetListSet) {
		String[][] data = new String [assetListSet.size()][31];
		int i = 0;
		for(AssetData a : assetListSet) {
			data[i][0] = a.getAssetName().toString();
			data[i][1] = a.getAssetNumber().toString();
			data[i][2] = a.getPurchaseDate().toString();
			data[i][3] = a.getPurchasePrice().toString();
			data[i][4] = a.getAssetType().toString();
			data[i][5] = "";
			data[i][6] = "";
			data[i][7] = "";
			data[i][8] = "";
			data[i][9] = "";
			data[i][10] = "";
			data[i][11] = "";
			data[i][12] = a.getPurchaseDate().toString();
			data[i][13] = "";
			data[i][14] = "";
			data[i][15] = "";
			data[i][16] = "Actual Days";
			data[i][17] = a.getBookRate().toString();
			data[i][18] = "";
			data[i][19] = a.getBookAccumulatedDepreciation().toString();
			data[i][20] = "";
			data[i][21] = "";
			data[i][22] = "";
			data[i][23] = "";
			data[i][24] = a.getPurchaseDate().toString();
			data[i][25] = "";
			data[i][26] = "";
			data[i][27] = "Actual Days";
			data[i][28] = a.getBookRate().toString();
			data[i][29] = "";
			data[i][30] = a.getBookAccumulatedDepreciation().toString();
			i++;
		}
		
		for (i = 0; i < data.length; i++) {
			for (int j = 0; j < data[i].length; j++) {
				if (data[i][j] != null && data[i][j].matches(".*\\(.*\\).*")) { // Checks for parentheses
					data[i][j] += " notice"; // Concatenates " notice" to the string
				}
			}
		}
 
 
		// Define the file path
        String filePath = "C://Users/test/Downloads/"+fileName.split(".xls")[0]+" - output.csv";

        // Fixed header for the CSV file
        String[] header = {"*AssetName","*AssetNumber","PurchaseDate","PurchasePrice","AssetType",
        		"Description","TrackingCategory1","TrackingOption1","TrackingCategory2","TrackingOption2",
        		"SerialNumber","WarrantyExpiry","Book_DepreciationStartDate","Book_CostLimit","Book_ResidualValue",
        		"Book_DepreciationMethod","Book_AveragingMethod","Book_Rate","Book_EffectiveLife","Book_OpeningBookAccumulatedDepreciation",
        		"Tax_DepreciationMethod","Tax_PoolName","Tax_PooledDate","Tax_PooledAmount","Tax_DepreciationStartDate",
        		"Tax_CostLimit","Tax_ResidualValue","Tax_AveragingMethod","Tax_Rate","Tax_EffectiveLife","Tax_OpeningAccumulatedDepreciation"};
        try {
        	File csvFile = new File(filePath);
            if (!csvFile.exists()) {
                if (csvFile.createNewFile()) {
                    System.out.println("CSV file created: " + csvFile.getAbsolutePath());
                } else {
                    System.err.println("Failed to create the CSV file.");
                    return;
                }
            }
            try (CSVWriter writer = new CSVWriter(new FileWriter(filePath))) {
                // Write the header
                writer.writeNext(header);

                // Write the data rows
                for (String[] row : data) {
                    writer.writeNext(row);
                }

                System.out.println("CSV file written successfully at " + filePath);

            } 
        }
        catch (IOException e) {
            System.err.println("Error writing to CSV file: " + e.getMessage());
        }
		}
	}
	
