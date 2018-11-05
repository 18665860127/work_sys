package com.jiang.work_sys.util.excel;

public enum ExcelEnum {
	EXCEL_XLS("xls"), EXCEL_XLSX("xlsx"), BASE_TEMPLATE("sheet_rows_cell");
	private final String type;

	ExcelEnum(String type) {
		this.type = type;
	}

	public String getType() {
		return type;
	}
}
