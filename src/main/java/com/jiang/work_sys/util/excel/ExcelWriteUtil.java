package com.jiang.work_sys.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

public class ExcelWriteUtil {

	private static DecimalFormat decimalFormat = new DecimalFormat("0.##");

	public static void payBankCard(Map<String, List<Map<String, String>>> basePersonInfo,
			List<Map<String, String>> payRecord, OutputStream ostream) throws IOException {
		Workbook wb = null;
		try {
			// 定义一个Excel表格
			wb = getWorkbok(ExcelEnum.EXCEL_XLSX);
			Sheet zaizhi = wb.createSheet("在职工资");
			double totalZaiZhi = 0d;
			payBankCardHead(zaizhi);
			Sheet lizhi = wb.createSheet("离职工资");
			double totalLiZhi = 0d;
			payBankCardHead(lizhi);
			Sheet fengxian = wb.createSheet("风险金");
			double totalFengXian = 0d;
			payBankCardHead(fengxian);
			// ---------------------------------------------输出对应的模板
			for (int i = 0; i < payRecord.size(); i++) {
				Map<String, String> pay = payRecord.get(i);
				Row createRow = null;
				String name = pay.get("name");
				String status_t = pay.get("status_t");
				String pay_money = decimalFormat.format(Double.parseDouble(pay.get("pay_money")));
				String fengxian_money = pay.get("fengxian_money");
				// String status = "";
				String card_no = "";
				String bank = "";
				String card_from = "";
				List<Map<String, String>> list = basePersonInfo.get(name);
				if (list != null && list.size() != 0) {
					Map<String, String> personInfo = list.get(0);
					// status = personInfo.get("status");
					card_no = personInfo.get("card_no");
					bank = personInfo.get("bank");
					card_from = personInfo.get("card_from");
				}
				if (!StringUtils.isEmpty(fengxian_money)) {
					int lastRowNum = fengxian.getPhysicalNumberOfRows();
					// lastRowNum++;
					createRow = fengxian.createRow(lastRowNum);
					totalFengXian += Double.parseDouble(pay_money);
				} else if (status_t.contains("离职")) {
					int lastRowNum = lizhi.getPhysicalNumberOfRows();
					// lastRowNum++;
					createRow = lizhi.createRow(lastRowNum);
					totalLiZhi += Double.parseDouble(pay_money);
				} else {
					int lastRowNum = zaizhi.getPhysicalNumberOfRows();
					// lastRowNum++;
					createRow = zaizhi.createRow(lastRowNum);
					totalZaiZhi += Double.parseDouble(pay_money);
				}
				if (createRow != null) {
					int index = createRow.getRowNum();
					createRow.createCell(0, CellType.STRING).setCellValue(index);
					createRow.createCell(1, CellType.STRING).setCellValue(card_no);
					createRow.createCell(2, CellType.STRING).setCellValue(name);
					createRow.createCell(3, CellType.STRING).setCellValue(pay_money);
					if (bank.contains("建")) {
						createRow.createCell(4, CellType.STRING).setCellValue("");
					} else {
						createRow.createCell(4, CellType.STRING).setCellValue("01");
					}
					createRow.createCell(5, CellType.STRING).setCellValue(bank);
					createRow.createCell(6, CellType.STRING).setCellValue(card_from);
					createRow.createCell(7, CellType.STRING).setCellValue("");
					if (list != null && list.size() > 1) {
						createRow.createCell(8, CellType.STRING).setCellValue("花名册存在重复数据，请手工确认信息是否正确！！！");
					}
				}
			}

			payBankCardFooter(zaizhi, totalZaiZhi);
			payBankCardFooter(lizhi, totalLiZhi);
			payBankCardFooter(fengxian, totalFengXian);

			// --------------------------------------------end
			// 输出流,下载时候的位置
			wb.write(ostream);
		} finally {
			if (ostream != null) {
				try {
					ostream.flush();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
				try {
					ostream.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private static void payBankCardFooter(Sheet sheet, double total) {
		Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
		{
			Cell cell = row.createCell(2, CellType.STRING);
			cell.setCellValue("合计");
		}
		{
			Cell cell = row.createCell(3, CellType.STRING);
			cell.setCellValue(decimalFormat.format(total));
		}
	}

	private static void payBankCardHead(Sheet sheet) {
		Row row = sheet.createRow(0);
		{
			Cell cell = row.createCell(0, CellType.STRING);
			cell.setCellValue("序号");
		}
		{
			Cell cell = row.createCell(1, CellType.STRING);
			cell.setCellValue("账号");
		}
		{
			Cell cell = row.createCell(2, CellType.STRING);
			cell.setCellValue("户名");
		}
		{
			Cell cell = row.createCell(3, CellType.STRING);
			cell.setCellValue("金额");
		}
		{
			Cell cell = row.createCell(4, CellType.STRING);
			cell.setCellValue("跨行标识（选填 建行填0 他行填1）");
		}
		{
			Cell cell = row.createCell(5, CellType.STRING);
			cell.setCellValue("行名");
		}
		{
			Cell cell = row.createCell(6, CellType.STRING);
			cell.setCellValue("开户支行");
		}
		{
			Cell cell = row.createCell(7, CellType.STRING);
			cell.setCellValue("备注");
		}

	}

	/**
	 * 工资条
	 * 
	 * @param basePersonInfo
	 * @param payRecord
	 * @param ostream
	 * @throws IOException
	 */
	public static void payroll(Map<String, List<Map<String, String>>> basePersonInfo,
			List<Map<String, String>> payRecord, OutputStream ostream) throws IOException {
		Workbook wb = null;
		try {
			// 定义一个Excel表格
			wb = getWorkbok(ExcelEnum.EXCEL_XLSX);
			Sheet sheet = wb.createSheet("工资条");
			// -------记录行数index
			int rowIndex = 0;
			for (int i = 0; i < payRecord.size(); i++) {
				Map<String, String> pay = payRecord.get(i);
				// ----数据
				String name = pay.get("name");
				String status_t = pay.get("status_t");
				String plate_number = pay.get("plate_number");
				String month = pay.get("month");
				String day_of_month = pay.get("day_of_month");
				String day_of_money = pay.get("day_of_money");
				String month_of_money = pay.get("month_of_money");
				String subsidy = pay.get("subsidy");
				String merit_pay = pay.get("merit_pay");
				String oil_wear_money = pay.get("oil_wear_money");
				String merit_fine = pay.get("merit_fine");
				String payable_money = pay.get("payable_money");
				String peccancy_fine = pay.get("peccancy_fine");
				String fengxian_money = pay.get("fengxian_money");
				String fengxian_money_fine = pay.get("fengxian_money_fine");
				String accident_insurance = pay.get("accident_insurance");
				String social_security = pay.get("social_security");
				String borrow_money = pay.get("borrow_money");
				String oil_wear_turn = pay.get("oil_wear_turn");
				String work_cloths_money = pay.get("work_cloths_money");
				String last_month = pay.get("last_month");
				String pit = pay.get("pit");
				String pay_money = decimalFormat.format(Double.parseDouble(pay.get("pay_money")));
				String card_no = "";
				String bank = "";
				String card_from = "";
				List<Map<String, String>> list = basePersonInfo.get(name);
				if (list != null && list.size() != 0) {
					Map<String, String> personInfo = list.get(0);
					// status = personInfo.get("status");
					card_no = personInfo.get("card_no");
					bank = personInfo.get("bank");
					card_from = personInfo.get("card_from");
				}
				// ----end
				Row row1 = sheet.createRow(sheet.getPhysicalNumberOfRows());
				Row row2 = sheet.createRow(sheet.getPhysicalNumberOfRows());
				Row row3 = sheet.createRow(sheet.getPhysicalNumberOfRows());
				Row row4 = sheet.createRow(sheet.getPhysicalNumberOfRows());
			}
			// --------------------------------------------end
			// 输出流,下载时候的位置
			wb.write(ostream);
		} finally {
			if (ostream != null) {
				try {
					ostream.flush();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
				try {
					ostream.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private static Workbook getWorkbok(ExcelEnum excelEnum) {
		if (ExcelEnum.EXCEL_XLS.equals(excelEnum)) {
			return new HSSFWorkbook();
		} else {
			return new XSSFWorkbook();
		}
	}

}
