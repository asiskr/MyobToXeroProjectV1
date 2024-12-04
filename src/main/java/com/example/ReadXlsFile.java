package com.example;
import java.util.List;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import com.opencsv.CSVWriter;

public class ReadXlsFile {
    @SuppressWarnings("deprecation")
    public static void main(String[] args) {
        SwingUtilities.invokeLater(ReadXlsFile::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("XLS to CSV Converter");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 300);
        frame.setLayout(new BorderLayout());

        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

        JLabel instructionLabel = new JLabel("Select one or more .xls files to process:");
        instructionLabel.setAlignmentX(Component.CENTER_ALIGNMENT);

        JButton selectFilesButton = new JButton("Select Files");
        selectFilesButton.setAlignmentX(Component.CENTER_ALIGNMENT);

        JLabel filePathLabel = new JLabel("No files selected");
        filePathLabel.setAlignmentX(Component.CENTER_ALIGNMENT);

        JButton processButton = new JButton("Process Files");
        processButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        processButton.setEnabled(false);

        // List to hold selected files
        List<File> selectedFiles = new ArrayList<>();

        selectFilesButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setMultiSelectionEnabled(true); // Allow multiple file selection
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files (*.xls)", "xls"));
            int returnValue = fileChooser.showOpenDialog(frame);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                selectedFiles.clear();
                selectedFiles.addAll(Arrays.asList(fileChooser.getSelectedFiles()));
                filePathLabel.setText("Selected: " + selectedFiles.size() + " files");
                processButton.setEnabled(true);
            }
        });

        processButton.addActionListener(e -> {
            if (!selectedFiles.isEmpty()) {
                // Process all files first
                List<Set<AssetData>> assetListSets = new ArrayList<>();
                for (File file : selectedFiles) {
                    try {
                        Set<AssetData> assetListSet = readXLSFile(file);
                        assetListSets.add(assetListSet); // Collect the asset data from each file
                    } catch (Exception ex) {
                        JOptionPane.showMessageDialog(frame, "Error processing file: " + file.getName() + "\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
                
                // Now ask the user to select the output directory
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setDialogTitle("Save Output CSV Files");
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int userChoice = fileChooser.showSaveDialog(frame);
                
                if (userChoice == JFileChooser.APPROVE_OPTION) {
                    File outputDirectory = fileChooser.getSelectedFile();
                    
                    // Now save each file's result as CSV
                    int i = 0;
                    for (Set<AssetData> assetListSet : assetListSets) {
                        File inputFile = selectedFiles.get(i++);
                        writeCSVFile(assetListSet, inputFile, outputDirectory);
                    }
                    
                    JOptionPane.showMessageDialog(frame, "Processing complete. CSV files saved to the selected directory.", "Success", JOptionPane.INFORMATION_MESSAGE);
                }
            }
        });

        panel.add(Box.createRigidArea(new Dimension(0, 10)));
        panel.add(instructionLabel);
        panel.add(Box.createRigidArea(new Dimension(0, 10)));
        panel.add(selectFilesButton);
        panel.add(Box.createRigidArea(new Dimension(0, 10)));
        panel.add(filePathLabel);
        panel.add(Box.createRigidArea(new Dimension(0, 10)));
        panel.add(processButton);

        frame.add(panel, BorderLayout.CENTER);
        frame.setVisible(true);
    }


    private static Set<AssetData> readXLSFile(File inputFile) {
		Set<AssetData> assetList = new HashSet<>();
		try (FileInputStream fis = new FileInputStream(inputFile); Workbook workbook = new HSSFWorkbook(fis)) {

			// Get the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			int empCellCount = 0;
			int rowCount = 0;

			// Iterate through each row
			AssetData assetData = null;
			int headerRowNum = 10;
			String regex = "^\\d{3}$|^[A-Z]+$";
			boolean flag = false;
			for (Row row : sheet) {
				if (row.getRowNum() > 10) {
					Cell empCell = row.getCell(0);
					empCellCount = empCell == null ? empCellCount+1 : 0;
					if(empCellCount > 2 || flag) {
						Cell checkCell = sheet.getRow(row.getRowNum()+1).getCell(0);
						if(checkCell != null && !checkCell.getStringCellValue().equalsIgnoreCase("Total") && checkCell.getStringCellValue().matches(regex)) {
							headerRowNum = row.getRowNum()+1;
							flag = false;
							continue;
						}
						if(checkCell != null && checkCell.getStringCellValue().equalsIgnoreCase("Total")) {
							flag = false;
							break;
						}
						if(checkCell != null && !checkCell.getStringCellValue().matches(regex)) {
							flag = true;
							continue;
						}
						if (checkCell == null || checkCell.getCellType() == Cell.CELL_TYPE_STRING) {
						    continue;  // Skip this iteration if cell is null or a string
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
							System.out.print(cell + " -> " + rowCount + "-" + cellCount + "\t");
							if (rowCount == 0 && cellCount == 0) {
							    assetData = new AssetData();
							    assetData.setAssetType(sheet.getRow(headerRowNum).getCell(3).getStringCellValue());
//							    System.out.println("Asest typoe" + assetData.getAssetType());
//							    System.out.println("Type of AssetType: " + assetData.getAssetType().getClass().getName());
							    
							    assetData.setAssetNumber(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
//							    System.out.println("Asest number" + assetData.getAssetNumber());
//							    System.out.println("Type of number: " + assetData.getAssetNumber().getClass().getName());
							    
							}

							if(rowCount == 0 && cellCount == 1) {
								assetData.setAssetName(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
							}
							if(rowCount == 1 && cellCount == 0) {
								assetData.setPurchaseDate(formatDate(cell.getDateCellValue()));  
							}
							if(rowCount == 1 && cellCount == 1) {
								assetData.setPurchasePrice(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
							}
							if(rowCount == 1 && cellCount == 4) {
								assetData.setBookRate(cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue());
							}
							if(rowCount == 1 && cellCount == 5) {
								String bookDepMethodValue = "";
								Object bookDepMethod = cell.getCellType() == Cell.CELL_TYPE_NUMERIC ? cell.getNumericCellValue() : cell.getStringCellValue();
								if(bookDepMethod.toString().equalsIgnoreCase("d")) {
									bookDepMethodValue = "Diminishing Value";
								}
								if(bookDepMethod.toString().equalsIgnoreCase("w")) {
									bookDepMethodValue = "Full Depreciation at purchase";
								}
								if(bookDepMethod.toString().equalsIgnoreCase("p")) {
									bookDepMethodValue = "Straight Line";
								}
								assetData.setDepnMethod(bookDepMethodValue);
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
		} catch (IOException e) {
			e.printStackTrace();
		}
		return assetList;
	}
    private static String formatDate(Object date) {
        if (date instanceof String) {
            String dateStr = (String) date;

            // Handle "dd-MMM-yyyy" format (e.g., "02-Oct-2009")
            if (dateStr.matches("\\d{2}-[a-zA-Z]{3}-\\d{4}")) {
                try {
                    // Parse "dd-MMM-yyyy" into a Date object
                    SimpleDateFormat inputFormat = new SimpleDateFormat("dd-MMM-yyyy");
                    Date parsedDate = inputFormat.parse(dateStr);

                    // Format into "M/d/yyyy"
                    SimpleDateFormat outputFormat = new SimpleDateFormat("M/d/yyyy");
                    return outputFormat.format(parsedDate);
                } catch (ParseException e) {
                    // Handle parsing error
                    System.out.println("Error parsing date: " + dateStr);
                    return "";
                }
            }

            // General "dd-mm-yyyy" or "dd-mm-yy" handling
            if (dateStr.matches("\\d{2}-\\d{2}-\\d{2,4}")) {
                String[] parts = dateStr.split("-");
                return Integer.parseInt(parts[1]) + "/" + Integer.parseInt(parts[0]) + "/" + parts[2];
            }

            // Replace '-' with '/' if present but no specific pattern matched
            if (dateStr.contains("-")) {
                return dateStr.replace("-", "/");
            }

            return dateStr;
        }

        return "";
    }
	private static void writeCSVFile(Set<AssetData> assetListSet ,File inputFile,File outputDirectory) {

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
			data[i][15] = a.getDepnMethod().toString();
			data[i][16] = "Actual Days";
			data[i][17] = a.getBookRate().toString();
			data[i][18] = "";
			data[i][19] = a.getBookAccumulatedDepreciation().toString();
			data[i][20] =  a.getDepnMethod().toString();
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
		            if (data[i][j] != null && data[i][j].matches(".*\\(.*\\).*")) {
		                data[i][j] += " notice"; // Append " notice"
		            }
		        }
		    }
		 String inputFileNameWithoutExtension = inputFile.getName().replaceFirst("[.][^.]+$", "");
		    String outputFilePath = outputDirectory.getAbsolutePath() + File.separator + inputFileNameWithoutExtension + ".csv";

		// Fixed header for the CSV file
		String[] header = {"*AssetName","*AssetNumber","PurchaseDate","PurchasePrice","AssetType",
				"Description","TrackingCategory1","TrackingOption1","TrackingCategory2","TrackingOption2",
				"SerialNumber","WarrantyExpiry","Book_DepreciationStartDate","Book_CostLimit","Book_ResidualValue",
				"Book_DepreciationMethod","Book_AveragingMethod","Book_Rate","Book_EffectiveLife","Book_OpeningBookAccumulatedDepreciation",
				"Tax_DepreciationMethod","Tax_PoolName","Tax_PooledDate","Tax_PooledAmount","Tax_DepreciationStartDate",
				"Tax_CostLimit","Tax_ResidualValue","Tax_AveragingMethod","Tax_Rate","Tax_EffectiveLife","Tax_OpeningAccumulatedDepreciation"};
		  // Write to CSV file
        try (CSVWriter writer = new CSVWriter(new FileWriter(outputFilePath))) {
            writer.writeNext(header);
            for (String[] row : data) {
                writer.writeNext(row);
            }
            JOptionPane.showMessageDialog(null, "CSV File Written Successfully: " + outputFilePath, 
                                          "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error writing CSV: " + e.getMessage(), 
                                          "Error", JOptionPane.ERROR_MESSAGE);
        }

	}
}

