
public class RegistroExcel {
	
	public static final String CELDA_VACIA ="";		// Orden 0 (GP Signal)
    private String signalClass; 					// Orden Se usa para identificar el tipo de señal
    private String isotrolSignalTag;				//* Orden 1 -> (Signal Name) se le concatena al inicio CALC o DATA dependiendo de la fuente (valor signalClass) si es WEB, se descarta y no se escribe, si es DUMMY se deberia escribir CACL, y si es PROTOCOL se deberia escribir DATA, tambieén deberia coincidir con el valor de source si es Protocol o Business
    private String isotrolEngDescription;			//* Orden 2 -> (Description)
    private String units;							// Orden 3 -> (Unit)
    private String scale;							//* Orden 6 -> (Scale)
    private String offset;							//* Orden 7 -> (Offset)
    private String lowerRange;						// Orden 4 -> (Min)
    private String upperRange;						// Orden 5 -> (Max)
    private String operation;						// Orden 8 -> (Read/Wirte) 
    private String source;							// Orden 11 -> (Source) Se debe escribir como 'Protocol' o 'Business' dependiendo de su valor en mayusculas
    private String storage;							// Orden 9 -> (Storable) Se debe escribir como Y o N el valor leido sera true o false ignoreCase
    private String storagePolicy;					// Orden 10 -> (Storable Policy)

    public RegistroExcel(String signalClass, String isotrolSignalTag, String isotrolEngDescription,
                      String units, String scale, String offset, String lowerRange, String upperRange,
                      String operation, String source, String storage, String storagePolicy) {
        this.signalClass = signalClass;
        this.isotrolSignalTag = isotrolSignalTag;
        this.isotrolEngDescription = isotrolEngDescription;
        this.units = units;
        this.scale = scale;
        this.offset = offset;
        this.lowerRange = lowerRange;
        this.upperRange = upperRange;
        this.operation = operation;
        this.source = source;
        this.storage = storage;
        this.storagePolicy = storagePolicy;
    }

    // Getters y setters opcionales

    @Override
    public String toString() {
        return "SignalInfo{" +
                "signalClass='" + signalClass + '\'' +
                ", isotrolSignalTag='" + isotrolSignalTag + '\'' +
                ", isotrolEngDescription='" + isotrolEngDescription + '\'' +
                ", units='" + units + '\'' +
                ", scale='" + scale + '\'' +
                ", offset='" + offset + '\'' +
                ", lowerRange='" + lowerRange + '\'' +
                ", upperRange='" + upperRange + '\'' +
                ", operation='" + operation + '\'' +
                ", source='" + source + '\'' +
                ", storage='" + storage + '\'' +
                ", storagePolicy='" + storagePolicy + '\'' +
                '}';
    }

	public String getSignalClass() {
		return signalClass;
	}

	public void setSignalClass(String signalClass) {
		this.signalClass = signalClass;
	}

	public String getIsotrolSignalTag() {
		return isotrolSignalTag;
	}

	public void setIsotrolSignalTag(String isotrolSignalTag) {
		this.isotrolSignalTag = isotrolSignalTag;
	}

	public String getIsotrolEngDescription() {
		return isotrolEngDescription;
	}

	public void setIsotrolEngDescription(String isotrolEngDescription) {
		this.isotrolEngDescription = isotrolEngDescription;
	}

	public String getUnits() {
		return units;
	}

	public void setUnits(String units) {
		this.units = units;
	}

	public String getScale() {
		return scale;
	}

	public void setScale(String scale) {
		this.scale = scale;
	}

	public String getOffset() {
		return offset;
	}

	public void setOffset(String offset) {
		this.offset = offset;
	}

	public String getLowerRange() {
		return lowerRange;
	}

	public void setLowerRange(String lowerRange) {
		this.lowerRange = lowerRange;
	}

	public String getUpperRange() {
		return upperRange;
	}

	public void setUpperRange(String upperRange) {
		this.upperRange = upperRange;
	}

	public String getOperation() {
		return operation;
	}

	public void setOperation(String operation) {
		this.operation = operation;
	}

	public String getSource() {
		return source;
	}

	public void setSource(String source) {
		this.source = source;
	}

	public String getStorage() {
		return storage;
	}

	public void setStorage(String storage) {
		this.storage = storage;
	}

	public String getStoragePolicy() {
		return storagePolicy;
	}

	public void setStoragePolicy(String storagePolicy) {
		this.storagePolicy = storagePolicy;
	}
	
	
	/*
	private String Signal_Class;
	private String Isotrol_Signal_Tag;
	private String Isotrol_ENG_description;
	private String Units;
	private String Scale;
	private String Offset;
	private String Lower_range;
	private String Upper_range;
	private String Operation;
	private String Source;
	private String Storage;
	private String Storage_Policy;
	
	public String getSignal_Class() {
		return Signal_Class;
	}
	public void setSignal_Class(String signal_Class) {
		Signal_Class = signal_Class;
	}
	public String getIsotrol_Signal_Tag() {
		return Isotrol_Signal_Tag;
	}
	public void setIsotrol_Signal_Tag(String isotrol_Signal_Tag) {
		Isotrol_Signal_Tag = isotrol_Signal_Tag;
	}
	public String getIsotrol_ENG_description() {
		return Isotrol_ENG_description;
	}
	public void setIsotrol_ENG_description(String isotrol_ENG_description) {
		Isotrol_ENG_description = isotrol_ENG_description;
	}
	public String getUnits() {
		return Units;
	}
	public void setUnits(String units) {
		Units = units;
	}
	public String getScale() {
		return Scale;
	}
	public void setScale(String scale) {
		Scale = scale;
	}
	public String getOffset() {
		return Offset;
	}
	public void setOffset(String offset) {
		Offset = offset;
	}
	public String getLower_range() {
		return Lower_range;
	}
	public void setLower_range(String lower_range) {
		Lower_range = lower_range;
	}
	public String getUpper_range() {
		return Upper_range;
	}
	public void setUpper_range(String upper_range) {
		Upper_range = upper_range;
	}
	public String getOperation() {
		return Operation;
	}
	public void setOperation(String operation) {
		Operation = operation;
	}
	public String getSource() {
		return Source;
	}
	public void setSource(String source) {
		Source = source;
	}
	public String getStorage() {
		return Storage;
	}
	public void setStorage(String storage) {
		Storage = storage;
	}
	public String getStorage_Policy() {
		return Storage_Policy;
	}
	public void setStorage_Policy(String storage_Policy) {
		Storage_Policy = storage_Policy;
	}
	*/
}
