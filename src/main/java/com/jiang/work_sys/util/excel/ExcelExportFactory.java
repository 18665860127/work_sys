package com.jiang.work_sys.util.excel;

import java.util.HashMap;
import java.util.Map;

import com.jiang.work_sys.util.excel.TemplateImpl.BaseTemplateImpl;


public class ExcelExportFactory {

	private static Map<ExcelEnum, ExportTemplate> excelExportMap = new HashMap<>();
	static {
		excelExportMap.put(ExcelEnum.BASE_TEMPLATE, new BaseTemplateImpl());
	}

	public static ExportTemplate getExportTemplate(ExcelEnum excelEnum) {
		return excelExportMap.get(excelEnum);
	}
}
