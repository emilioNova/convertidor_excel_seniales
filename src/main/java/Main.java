
import java.util.List;
import java.util.Map;
import java.util.Scanner;


public class Main {

	// Ruta absoluta del fichero de documentacion del elemento (nombre y extensión incluidos)
	static String inputFilePath = "";

	// Ruta absoluta del fichero de salida (nombre y extensión incluidos)
	static String outputFilePath = "";
	
	// Nombre de la hoda donde están los datos para las señales 
	static String sheetName = "";
	
	
    public static void main(String[] args) {
/*
    	inputFilePath = "C:\\Users\\emilio-desarrollo\\Desktop\\pruebas_excel\\Business_model_INV_GAMESA_PV_4300_SERIES_3X_v03.xlsx";
    	outputFilePath = "C:\\Users\\emilio-desarrollo\\Desktop\\pruebas_excel\\ggggggggggggggggg.xlsx";
        sheetName = "Data";
  */  	
    	// pide las rutas por consola
    	ingresaRutasParam();
    	
        List<RegistroExcel> signalInfoList = Utill.obtencionDatos(inputFilePath, sheetName); 
        Utill.escrituraDocumento(outputFilePath, signalInfoList);
        System.out.println("Resultado de las validaciones: ");
        // Iterar sobre el Map
        for (Map.Entry<String, String> entry : Utill.erroresEnExcel.entrySet()) {
            System.out.println("En el registro Nº: " + entry.getKey() + ", Se ha producido el siguiente error: " + entry.getValue());
        }
    }

	private static void ingresaRutasParam() {
    	
        //System.out.println("inputFilePath al inicio:  "+inputFilePath);
        Scanner scanner = new Scanner(System.in);
        
        System.out.println("Ingresa I para info, enter para continuar");
        if(scanner.nextLine().equalsIgnoreCase("i")) {
        	System.out.println("El Excel de entrada debera tener las columnas de donde extraer la información con el siguiente orden\r\n"
        			+ "Signal Class -> Columna A\r\n"
        			+ "Columna vacia -> Columna B\r\n"
        			+ "Isotrol Signal Tag -> Columna C\r\n"
        			+ "Isotrol ENG description * -> Columna D\r\n"
        			+ "Isotrol SPA description -> Columna E\r\n"
        			+ "Units -> Columna F\r\n"
        			+ "Scale -> Columna G\r\n"
        			+ "Offset -> Columna H\r\n"
        			+ "Lower range  -> Columna I\r\n"
        			+ "Upper range -> Columna J\r\n"
        			+ "Operation -> Columna K\r\n"
        			+ "Source -> Columna L\r\n"
        			+ "Storage -> Columna M\r\n"
        			+ "Storage Policy -> Columna N");
 
        	System.out.println("");
        	System.out.println("Se debe introduci en nombre de la hoja, normalmente suele ser 'Data'");
        }
        
        System.out.println("Ingresa la ruta absoluta del excel (C:\\micarpeta\\mifichero.xlsx):");
        inputFilePath = scanner.nextLine();
        //System.out.println("inputFilePath despues escaner:  "+inputFilePath);
        
        System.out.println("Ingresa el nombre de la hoja del Excel de donde se leeran los datos (DATA):");
        sheetName = scanner.nextLine();
        
        System.out.println("Ingresa la ruta absoluta de salida del excel modificado (C:\\micarpeta\\mificheroDeSalida.xlsx):");
        outputFilePath = scanner.nextLine();
       /* 
        System.out.println("inputFilePath despues escaner:  "+inputFilePath);
        System.out.println("sheetName despues escaner:  "+sheetName);
        System.out.println("outputFilePath despues escaner:  "+outputFilePath);
        
        */
        scanner.close();
	}


}