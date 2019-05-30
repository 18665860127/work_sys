package com.jiang.work_sys.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

public class ExcelWriteUtil {

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
				String pay_money = StringUtils.isEmpty(pay.get("pay_money")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("pay_money")));
				String fengxian_money = StringUtils.isEmpty(pay.get("fengxian_money")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("fengxian_money")));
				
				
				if(StringUtils.isEmpty(pay_money)){
					System.out.println();
				}
				
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
			cell.setCellValue(roundKeepZero(total));
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
			CellStyle createCellStyle = wb.createCellStyle();
			Font createFont = wb.createFont();
			createFont.setFontName("微软雅黑");
			createFont.setBold(true);
			createFont.setFontHeightInPoints((short) 7);
			createCellStyle.setFont(createFont);
			createCellStyle.setBorderBottom(BorderStyle.THIN);
			createCellStyle.setBorderLeft(BorderStyle.THIN);
			createCellStyle.setBorderRight(BorderStyle.THIN);
			createCellStyle.setBorderTop(BorderStyle.THIN);
			createCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			createCellStyle.setAlignment(HorizontalAlignment.LEFT);
			createCellStyle.setWrapText(true);
			// createCellStyle.set
			// -------记录行数index
			int rowIndex = 0;
			for (int i = 0; i < payRecord.size(); i++) {
				Map<String, String> pay = payRecord.get(i);
				// ----数据
				String name = pay.get("name");
				if (name.equals("谢勇")) {
					System.out.println();
				}
				String status_t = pay.get("status_t");
				String plate_number = pay.get("plate_number");
				String month = pay.get("month");
				String day_of_month = pay.get("day_of_month");
				String day_of_money = StringUtils.isEmpty(pay.get("day_of_money")) ? ""
						: roundKeepTwo(Double.parseDouble(pay.get("day_of_money")));
				String month_of_money = pay.get("month_of_money");
				String subsidy = pay.get("subsidy");
				String merit_pay = pay.get("merit_pay");
				String oil_wear_money = StringUtils.isEmpty(pay.get("oil_wear_money")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("oil_wear_money")));
				String merit_fine = pay.get("merit_fine");
				String payable_money = StringUtils.isEmpty(pay.get("payable_money")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("payable_money")));
				String peccancy_fine = pay.get("peccancy_fine");
				String fengxian_money = pay.get("fengxian_money");
				String fengxian_money_fine = pay.get("fengxian_money_fine");
				String accident_insurance = pay.get("accident_insurance");
				String social_security = StringUtils.isEmpty(pay.get("social_security")) ? ""
						: roundKeepOne(Double.parseDouble(pay.get("social_security")));
				String borrow_money = pay.get("borrow_money");
				String oil_wear_turn = pay.get("oil_wear_turn");
				String work_cloths_money = pay.get("work_cloths_money");
				String last_month = pay.get("last_month");
				String pit = StringUtils.isEmpty(pay.get("pit")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("pit")));
				String pay_money = StringUtils.isEmpty(pay.get("pay_money")) ? ""
						: roundKeepZero(Double.parseDouble(pay.get("pay_money")));
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
				// -----------设置数据
				String[] rowCell1 = new String[] { "", "姓名", "车牌", "月份", "上班天数", "每天底薪", "基本底薪", "主班/补贴", "绩效工资",
						"油耗扣奖", "绩效奖罚", "应付合计", "违章罚款", "辞职退风险金", "扣风险金", "意外险", "社会保险", "借支", "油耗转接", "扣工作服", "调整上月",
						"个人所得税", "实付工资" };
				String[] rowCell2 = new String[] { status_t, name, plate_number, month, day_of_month, day_of_money,
						month_of_money, subsidy, merit_pay, oil_wear_money, merit_fine, payable_money, peccancy_fine,
						fengxian_money, fengxian_money_fine, accident_insurance, social_security, borrow_money,
						oil_wear_turn, work_cloths_money, last_month, pit, pay_money };
				String[] rowCell3 = new String[] { "", "账号（必填）", "", "", "", "户名（必填）", "", "行名", "", "开户支行", "", "备注",
						"", "", "", "", "", "", "", "", "", "", "" };
				String[] rowCell4 = new String[] { "", card_no, "", "", "", name, "", bank, "", card_from, "", "", "",
						"", "", "", "", "", "", "", "", "", "" };
				setRowCells(row1, rowCell1, createCellStyle);
				setRowCells(row2, rowCell2, createCellStyle);
				setRowCells(row3, rowCell3, createCellStyle);
				setRowCells(row4, rowCell4, createCellStyle);
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 1, 4));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 5, 6));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 7, 8));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 9, 10));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 11, 13));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 14, 22));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 1, 4));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 5, 6));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 7, 8));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 9, 10));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 11, 13));
				sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 14, 22));
				rowIndex += 4;
			}
			// --------------------------------------------end
			// 输出流,下载时候的位置
			// ---------设置列宽
			float[] columnWith = new float[] { 0.92f, 3.72f, 7.22f, 1.22f, 2.85f, 3.6f, 3.35f, 2.97f, 3.85f, 3.6f,
					3.35f, 3.97f, 3.6f, 4.35f, 3.22f, 2.85f, 4.72f, 3.97f, 3.85f, 2.47f, 2.85f, 3.1f, 5.85f };
			for (int i = 0; i < columnWith.length; i++) {
				int width = (int) (261 * columnWith[i] + 184);
				sheet.setColumnWidth(i, width);
			}
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

	private static void setRowCells(Row row, String[] rowCell, CellStyle createCellStyle) {
		row.setHeightInPoints(35.1f);
		for (int j = 0; j < rowCell.length; j++) {
			String val = rowCell[j];
			Cell createCell = row.createCell(j);
			createCell.setCellType(CellType.STRING);
			createCell.setCellStyle(createCellStyle);
			createCell.setCellValue(val);
		}
	}

	private static Workbook getWorkbok(ExcelEnum excelEnum) {
		if (ExcelEnum.EXCEL_XLS.equals(excelEnum)) {
			return new HSSFWorkbook();
		} else {
			return new XSSFWorkbook();
		}
	}

	public static String roundKeepOne(double param) {
		if (param < 0) {
			param = 0 - param;
			return "-" + ((double) Math.round(param * 10) / 10);
		} else {
			return ((double) Math.round(param * 10) / 10) + "";
		}
	}

	public static String roundKeepTwo(double param) {
		if (param < 0) {
			param = 0 - param;
			return "-" + ((double) Math.round(param * 100) / 100);
		} else {
			return ((double) Math.round(param * 100) / 100) + "";
		}
	}

	public static String roundKeepZero(double param) {
		if (param < 0) {
			param = 0 - param;
			return "-" + (Math.round(param)) + "";
		} else {
			return (Math.round(param)) + "";
		}
	}

	public static void main(String[] args) {
		System.out.println(roundKeepOne(205.35));
	}
}
