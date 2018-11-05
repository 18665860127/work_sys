package com.jiang.work_sys.util.excel;

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExportTemplate {
	void exportExcel(Workbook wb, List<List<List<String>>> sheetList);
}
