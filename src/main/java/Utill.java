import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.HashMap;
import java.util.Map;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.Color;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utill {
	
	
	public static Map<String, String> erroresEnExcel = new HashMap<>();
	
	
	public static void escrituraDocumento(String excelFilePathOut, List<RegistroExcel> signalInfoList) {
		try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("SignalData");

            // Estilo para texto en negrita
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);

            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);
            
            
            // Cabecera
            Row head = sheet.createRow(0);
            head.createCell(0).setCellValue("GP Signal");
            head.createCell(1).setCellValue("Signal Name");
            head.createCell(2).setCellValue("Description");
            head.createCell(3).setCellValue("Unit");
            head.createCell(4).setCellValue("Min");
            head.createCell(5).setCellValue("Max");
            head.createCell(6).setCellValue("Scale");
            head.createCell(7).setCellValue("Offset");
            head.createCell(8).setCellValue("Read/Wirte");
            head.createCell(9).setCellValue("Storable");
            head.createCell(10).setCellValue("Storable Policy");
            head.createCell(11).setCellValue("Source");
            head.setRowStyle(boldStyle);
            
            int rowNum = 1;
            for (RegistroExcel signalInfo : signalInfoList) {
                
            	if(signalInfo.getSignalClass().equalsIgnoreCase("DUMMY") || signalInfo.getSignalClass().equalsIgnoreCase("PROTOCOL")) {
	            	
            		// estilos color celdas
                    // Crear un estilo de celda con color de fondo
                    CellStyle style = workbook.createCellStyle();
                    
                    Color colorAzul = new Color(0xC6, 0xD9, 0xF1); // Azul
                    Color colorVerde = new Color(0x92, 0xD0, 0x50); // verde
                    
                    if(signalInfo.getSignalClass().equalsIgnoreCase("PROTOCOL")) {
                    	style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex()); // Cambiar a tu color deseado
                    	//style.setFillForegroundColor(new XSSFColor(colorAzul)); 
                    }else {
                    	style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex()); // Cambiar a tu color deseado
                    	//style.setFillForegroundColor(new XSSFColor(colorVerde)); 
                    }
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            		
            		
            		//
            		Row row = sheet.createRow(rowNum++);
	                row.createCell(0).setCellValue(RegistroExcel.CELDA_VACIA);
	                
	                // se le concatena al inicio CALC o DATA dependiendo de la fuente (valor signalClass) 
	                // si es WEB, se descarta y no se escribe, 
	                // si es DUMMY se deberia escribir CACL, 
	                // si es PROTOCOL se deberia escribir DATA, tambieén deberia coincidir con el valor de source si es Protocol o Business
	                row.createCell(1).setCellValue(formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()));
	                if(signalInfo.getIsotrolSignalTag().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum), "El campo Signal Name no puede estar vacio, El registro no se incluira");
	                	rowNum--;
	                }
	                row.createCell(2).setCellValue(signalInfo.getIsotrolEngDescription());
	                if(signalInfo.getIsotrolEngDescription().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Description no puede estar vacio");
	                }
	                row.createCell(3).setCellValue(signalInfo.getUnits());
	                
	                row.createCell(4).setCellValue(signalInfo.getLowerRange());
	                if(signalInfo.getLowerRange().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Min no puede estar vacio");
	                }
	                row.createCell(5).setCellValue(signalInfo.getUpperRange());
	                if(signalInfo.getUpperRange().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Max no puede estar vacio");
	                }
	                row.createCell(6).setCellValue(signalInfo.getScale());
	                if(signalInfo.getScale().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Scale no puede estar vacio");
	                }
	                row.createCell(7).setCellValue(signalInfo.getOffset());
	                if(signalInfo.getOffset().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Offset no puede estar vacio");
	                }
	                row.createCell(8).setCellValue(signalInfo.getOperation());
	                
	                // Se debe escribir como Y o N el valor leido sera true o false ignoreCase
	                row.createCell(9).setCellValue(formatStorage(signalInfo.getStorage()));
	                if(signalInfo.getStorage().equals("")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Storable no puede estar vacio");
	                }
	                row.createCell(10).setCellValue(signalInfo.getStoragePolicy());
	                if(signalInfo.getStorage().equals("") && formatStorage(signalInfo.getStorage()).equalsIgnoreCase("y")) {
	                	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Storable Policy no puede estar vacio si el campo Storable esta en True");
	                }
	                
	                // Se debe escribir como 'Protocol' o 'Business' dependiendo de su valor en mayusculas
	                row.createCell(11).setCellValue(formatSource(signalInfo.getSource()));
	                
	                
	                
	                
	                
	                style.setBorderBottom(BorderStyle.THIN);
	                style.setBorderTop(BorderStyle.THIN);
	                style.setBorderRight(BorderStyle.THIN);
	                style.setBorderLeft(BorderStyle.THIN);
    
	                row.setRowStyle(style);
            	}               
            	if(signalInfo.getSignalClass().equalsIgnoreCase("WEB")) {
            		erroresEnExcel.put(String.valueOf(rowNum)+ " WEB", "Se omite la inclusión de este registro por ser de calse WEB");
            	}
            	
            	
            	
            }
            
            
            // Establecer ancho de la columna basado en el contenido más largo
            for (int colNum = 0; colNum < signalInfoList.size(); colNum++) {
                sheet.autoSizeColumn(colNum);
            }
            System.out.println("Se hán escrito: "+signalInfoList.size());
            // Guardar el archivo Excel
            try (FileOutputStream outputStream = new FileOutputStream(excelFilePathOut)) {
                workbook.write(outputStream);
                System.out.println("El archivo Excel ha sido creado satisfactoriamente.");
            }
            
        } catch (IOException e) {
            e.printStackTrace();
        }
	}

   private static String formatSource(String storage) {
    	
    	String result = "";
    	if(storage.equalsIgnoreCase("Business")) {
    		result = "Business";
    	}
    	if(storage.equalsIgnoreCase("Protocol")) {
    		result = "Protocol";
    	}
    	return result;
    }
    
   private static String formatStorage(String storage) {
    	
    	String result = "";
    	if(storage.equalsIgnoreCase("true")) {
    		result = "Y";
    	}
    	if(storage.equalsIgnoreCase("false")) {
    		result = "N";
    	}
    	return result;
    }
   
    private static String formatSignal(String signalName, String signalClass) {
    	
    	String result = "";
    	if(signalClass.equalsIgnoreCase("PROTOCOL")) {
    		result = "DATA.".concat(signalName);
    	}
    	if(signalClass.equalsIgnoreCase("DUMMY")) {
    		result = "CALC.".concat(signalName);
    	}
    	return result;
    }
    
    
	public static  List<RegistroExcel> obtencionDatos(String excelFilePath, String sheetName) {
		
		List<RegistroExcel> signalInfoList = new ArrayList<RegistroExcel>();
		try {
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);

            

            
            
            int lastRowIndex = -1;
            int columnIndex = 0; // Columna "A"

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getCellType() != CellType.BLANK) {
                    lastRowIndex = row.getRowNum();
                }
            }

           

            //int lastRowIndex = sheet.getLastRowNum();
            for (int rowIndex = 1; rowIndex <= lastRowIndex; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    String signalClass = getCellValueAsString(row.getCell(0));
                    String isotrolSignalTag = getCellValueAsString(row.getCell(2));
                    String isotrolEngDescription = getCellValueAsString(row.getCell(3));
                    String units = getCellValueAsString(row.getCell(5));
                    String scale = getCellValueAsString(row.getCell(6));
                    String offset = getCellValueAsString(row.getCell(7));
                    String lowerRange = getCellValueAsString(row.getCell(8));
                    String upperRange = getCellValueAsString(row.getCell(9));
                    String operation = getCellValueAsString(row.getCell(10));
                    String source = getCellValueAsString(row.getCell(11));
                    String storage = getCellValueAsString(row.getCell(12));
                    String storagePolicy = getCellValueAsString(row.getCell(13));

                    RegistroExcel signalInfo = new RegistroExcel(signalClass, isotrolSignalTag, isotrolEngDescription,
                            units, scale, offset, lowerRange, upperRange, operation, source, storage, storagePolicy);

                    signalInfoList.add(signalInfo);
                }
            }

            inputStream.close();

            // Imprimir la lista de objetos SignalInfo
          //  for (RegistroExcel signalInfo : signalInfoList) {
//                System.out.println(signalInfo);
  //          }
            System.out.println("Se hán leido: "+signalInfoList.size());
            workbook.close();
            

        } catch (IOException e) {
            e.printStackTrace();
        }
		return signalInfoList;
	}
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}
