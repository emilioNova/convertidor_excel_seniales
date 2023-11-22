import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.HashMap;
import java.util.Map;
import java.io.BufferedWriter;

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
import java.util.TreeMap;
public class Utill {
	
	private static final String CABECERA_DOC = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
			+ "<Template>\r\n"
			+ "	<Attributes>\r\n"
			+ "		<Attribute Name=\"Group Name\" Key=\"$GroupName$\"/>\r\n"
			+ "	</Attributes>\r\n"
			+ "	<Body>\r\n"
			+ "		<![CDATA[<DummyGroup GroupName=\"$GroupName$\">\r\n"
			+ "			<DummyItems>\r\n";
	private static final String PIE_DOC = "</DummyItems>\r\n"
			+ "		</DummyGroup>]]>\r\n"
			+ "	</Body>\r\n"
			+ "</Template>\r\n";
	
	/* EJEMPLO ITEM del xml
	 
		<DummyItem DummyItemName="DATA.VoltageRN">
			<SignalGenerated Name="DATA.VoltageRN" Description="AC Voltage Phase RN" Protocol="DummyProtocol" Unit="V" Bus="Bus" Device="Device" Address="Address" Offset="0" Scale="0.1" RangeLowerLimit="0" RangeUpperLimit="2000" Storable="False" Operation="R" Source="Protocol"/>
		</DummyItem>
	 
<DummyItem DummyItemName="DATA.VoltageRN">
	<SignalGenerated 
	Name="DATA.VoltageRN" 
	Description="AC 
	Voltage Phase RN" 
	Protocol="DummyProtocol" 
	Unit="V" 
	Bus="Bus" 
	Device="Device" 
	Address="Address" 
	Offset="0" 
	Scale="0.1" 
	RangeLowerLimit="0" 
	RangeUpperLimit="2000" 
	Storable="False" 
	Operation="R" 
	Source="Protocol"/>
</DummyItem>
	 
	 */
	
	
	private String itemXmlss ="<DummyItem DummyItemName=\"DATA.VoltageRN\">\r\n"
			+ "	<SignalGenerated \r\n"
			+ "	Name=\"DATA.VoltageRN\" \r\n"
			+ "	Description=\"AC \r\n"
			+ "	Voltage Phase RN\" \r\n"
			+ "	Protocol=\"DummyProtocol\" \r\n"
			+ "	Unit=\"V\" \r\n"
			+ "	Bus=\"Bus\" \r\n"
			+ "	Device=\"Device\" \r\n"
			+ "	Address=\"Address\" \r\n"
			+ "	Offset=\"0\" \r\n"
			+ "	Scale=\"0.1\" \r\n"
			+ "	RangeLowerLimit=\"0\" \r\n"
			+ "	RangeUpperLimit=\"2000\" \r\n"
			+ "	Storable=\"False\" \r\n"
			+ "	Operation=\"R\" \r\n"
			+ "	Source=\"Protocol\"/>\r\n"
			+ "</DummyItem>";
	
	public static Map<String, String> erroresEnExcel = new TreeMap<>();
	
	
	public static void escrituraDocumentoXml(String excelFilePathOut, List<RegistroExcel> signalInfoList) {
		
		// Se escribe la cabecera del documento
		
		final String CELDA_VACIA ="";
		String signalClass=""; 			
		String isotrolSignalTag="";		
		String isotrolEngDescription="";	
		String units="";		
		String scale="";		
		String offset="";		
		String lowerRange="";
		String upperRange="";
		String operation="";
		String source="";	
		String storage="";
		String storagePolicy="";

	       
		 try (BufferedWriter writer = new BufferedWriter(new FileWriter(excelFilePathOut))) {	
	            int rowNum = 1;	
	            String itemXml ="";
	            
	            writer.write(CABECERA_DOC);
	            
			for(RegistroExcel signalInfo : signalInfoList) {
				
				if(signalInfo.getSignalClass().equalsIgnoreCase("DUMMY") || signalInfo.getSignalClass().equalsIgnoreCase("PROTOCOL")) {
							
					isotrolSignalTag=formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass());
				    if(signalInfo.getIsotrolSignalTag().equals("")) {
				    	isotrolSignalTag="ValorNoMapeadoNoDefinido_"+rowNum+"_isotrolSignalTag";
				    	erroresEnExcel.put(String.valueOf(rowNum), "El campo Signal Name no puede estar vacio, El registro no se incluira");
				    }
					
					isotrolEngDescription=signalInfo.getIsotrolEngDescription();
				    if(signalInfo.getIsotrolEngDescription().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Description no puede estar vacio");
				    	isotrolEngDescription="ValorNoMapeadoNoDefinido_"+rowNum+"_isotrolEngDescription";
				    }
				    
					units=signalInfo.getUnits();
					
					scale=signalInfo.getScale();
				    if(signalInfo.getScale().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Scale no puede estar vacio");
				    	scale="ValorNoMapeadoNoDefinido_"+rowNum+"_scale";
				    }
					
					offset=signalInfo.getOffset();	
				    if(signalInfo.getOffset().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Offset no puede estar vacio");
				    	offset="ValorNoMapeadoNoDefinido_"+rowNum+"_offset";
				    }
					lowerRange=signalInfo.getLowerRange();
				    if(signalInfo.getLowerRange().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Min no puede estar vacio");
				    	lowerRange="ValorNoMapeadoNoDefinido_"+rowNum+"_lowerRange";
				    }
					upperRange=signalInfo.getUpperRange();
				    if(signalInfo.getUpperRange().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Max no puede estar vacio");
				    	upperRange="ValorNoMapeadoNoDefinido_"+rowNum+"_upperRange";
				    }
					
					storage=formatStorage(signalInfo.getStorage());
				    if(signalInfo.getStorage().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Storable no puede estar vacio");
				    	storage="ValorNoMapeadoNoDefinido_"+rowNum+"_storage";
				    }
					
					operation=signalInfo.getOperation();
					
					
					
					source=formatSource(signalInfo.getSource());	
				    if(signalInfo.getStorage().equals("")) {
				    	erroresEnExcel.put(String.valueOf(rowNum)+" referente a la señal: "+formatSignal(signalInfo.getIsotrolSignalTag(), signalInfo.getSignalClass()), "El campo Source no puede estar vacio");
				    }				
					
					
					itemXml ="				<DummyItem DummyItemName=\""+isotrolSignalTag+">\r\n"
							+ "					<SignalGenerated "
							+ "Name=\""+isotrolSignalTag+"\" "
							+ "Description=\""+isotrolEngDescription+"\" "
							+ "Protocol=\"DummyProtocol\" "
							+ "Unit=\""+units.replace("°C", "\\u00B0C")+"\" n"
							+ "Bus=\"Bus\" "
							+ "Device=\"Device\" "
							+ "Address=\"Address\" "
							+ "Offset=\""+offset+"\" "
							+ "Scale=\""+scale+"\" "
							+ "RangeLowerLimit=\""+lowerRange+"\" "
							+ "RangeUpperLimit=\""+upperRange+"\" "
							+ "Storable=\""+storage+"\" "
							+ "Operation=\""+operation+"\" "
							+ "Source=\""+source+"\"/>\r\n"
							+ "				</DummyItem>\r\n";
					
			           
					writer.write(itemXml);
				}
				if(signalInfo.getSignalClass().equalsIgnoreCase("WEB")) {
            		erroresEnExcel.put(String.valueOf(rowNum)+ " WEB", "Se omite la inclusión de este registro por ser de clase WEB");
				}

		            rowNum++;
			}
		
			
           
            System.out.println("Texto escrito en el archivo '" + excelFilePathOut + "' correctamente.");
     
			
		// Se escribe el pie del documento
            writer.write(PIE_DOC);
		
		
		
		
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
