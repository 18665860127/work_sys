package com.jiang.work_sys.util.excel.TemplateImpl;

import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.jiang.work_sys.util.excel.ExportTemplate;


public class BaseTemplateImpl implements ExportTemplate {

	@Override
	public void exportExcel(Workbook wb, List<List<List<String>>> sheetList) {

		for (int s = 0; s < sheetList.size(); s++) {
			List<List<String>> rows = sheetList.get(s);
			Sheet sheet = wb.createSheet();
			for (int r = 0; r < rows.size(); r++) {
				List<String> cells = rows.get(r);
				Row row = sheet.createRow(r);
				for (int c = 0; c < cells.size(); c++) {
					String cell = cells.get(c);
					row.createCell(c, CellType.STRING).setCellValue(cell);
				}
			}
		}
	}

}
