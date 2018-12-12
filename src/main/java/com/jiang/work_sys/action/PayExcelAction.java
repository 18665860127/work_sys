package com.jiang.work_sys.action;

import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.jiang.work_sys.util.excel.ExcelEnum;
import com.jiang.work_sys.util.excel.ExcelReadUtil;
import com.jiang.work_sys.util.excel.ExcelWriteUtil;

@Controller("payExcelAction")
@RequestMapping("/a")
public class PayExcelAction {

	@RequestMapping("a")
	public String uploadPage() {
		return "a/uploadPage";
	}

	private static ExecutorService threadPool = Executors.newFixedThreadPool(5);

	private final static long maxFileSize = 10485760l;// 10MB

	private final static String nameFilesPath = "d:/tempExcelFile/nameFiles/";
	private final static String payFilesPath = "d:/tempExcelFile/payFiles/";

	// private final static String nameFilesPath =
	// "/usr/local/tempExcelFile/nameFiles/";
	// private final static String payFilesPath =
	// "/usr/local/tempExcelFile/payFiles/";

	@RequestMapping("b")
	public void uploadPayExcelChange(HttpServletResponse rep, @RequestParam("nameFile") MultipartFile nameFile,
			@RequestParam("payFile") MultipartFile payFile, @RequestParam("exportType") String exportType)
			throws Exception {
		{
			if (nameFile == null || payFile == null) {
				return;
			}
			String originalFilename = nameFile.getOriginalFilename();
			String originalFilename2 = payFile.getOriginalFilename();
			if ((!originalFilename.contains(ExcelEnum.EXCEL_XLS.getType()))
					&& (!originalFilename.contains(ExcelEnum.EXCEL_XLSX.getType()))) {
				return;
			}
			if ((!originalFilename2.contains(ExcelEnum.EXCEL_XLS.getType()))
					&& (!originalFilename2.contains(ExcelEnum.EXCEL_XLSX.getType()))) {
				return;
			}
			long size = nameFile.getSize();
			long size2 = payFile.getSize();
			if (size > maxFileSize || size2 > maxFileSize) {
				return;
			}
		}
		CountDownLatch latch = new CountDownLatch(2);
		rep.setHeader("content-type", "application/octet-stream");
		rep.setContentType("application/octet-stream; charset=utf-8");
		// 3.设置content-disposition响应头控制浏览器以下载的形式打开文件
		long nowTime = System.currentTimeMillis();
		File newNameFile = transferTo(nameFile, nameFilesPath, nowTime);
		File newPayFile = transferTo(payFile, payFilesPath, nowTime);
		// ----读花名册
		Map<String, List<Map<String, String>>> basePersonInfo = new HashMap<>();
		threadPool.execute(new Runnable() {
			@Override
			public void run() {
				try {
					readNewNameFile(newNameFile, basePersonInfo);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					latch.countDown();
				}
			}
		});
		// ----读工资
		List<Map<String, String>> payRecord = new ArrayList<>();
		threadPool.execute(new Runnable() {
			@Override
			public void run() {
				try {
					readNewPayFile(newPayFile, payRecord);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					latch.countDown();
				}
			}
		});
		latch.await();
		if ("0".equals(exportType)) {
			rep.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode("司机银行卡.xlsx", "UTF-8")); //
			ExcelWriteUtil.payBankCard(basePersonInfo, payRecord, rep.getOutputStream());
		}
		if ("1".equals(exportType)) {
			rep.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode("工资条.xlsx", "UTF-8")); //
			ExcelWriteUtil.payroll(basePersonInfo, payRecord, rep.getOutputStream());
		}
	}

	private List<Map<String, String>> readNewPayFile(File newPayFile, List<Map<String, String>> payRecord)
			throws InvalidFormatException, IOException {
		// List<Map<String, String>> payRecord = new ArrayList<>();
		List<List<String>> sheet = ExcelReadUtil.readExcel(newPayFile, 0, 3);
		for (int i = 0; i < sheet.size(); i++) {
			List<String> row = sheet.get(i);
			Map<String, String> pay = new HashMap<>();
			for (int j = 0; j < row.size(); j++) {
				String cell = row.get(j);
				// 是否离职
				if (j == 0) {
					pay.put("status_t", cell);
				}
				// 姓名。作为唯一key
				if (j == 2) {
					if (StringUtils.isEmpty(cell)) {
						break;
					}
					pay.put("name", cell);
					payRecord.add(pay);
				}
				// 车牌
				if (j == 3) {
					pay.put("plate_number", cell);
				}
				// 月份
				if (j == 5) {
					pay.put("month", cell);
				}
				// 上班天数
				if (j == 6) {
					pay.put("day_of_month", cell);
				}
				// 日薪
				if (j == 7) {
					pay.put("day_of_money", cell);
				}
				// 月薪
				if (j == 8) {
					pay.put("month_of_money", cell);
				}
				// 主班/补贴
				if (j == 9) {
					pay.put("subsidy", cell);
				}
				// 绩效工资
				if (j == 10) {
					pay.put("merit_pay", cell);
				}
				// 油耗扣奖
				if (j == 11) {
					pay.put("oil_wear_money", cell);
				}
				// 绩效奖罚
				if (j == 12) {
					pay.put("merit_fine", cell);
				}
				// 应付合计
				if (j == 13) {
					pay.put("payable_money", cell);
				}
				// 违章罚款
				if (j == 14) {
					pay.put("peccancy_fine", cell);
				}
				// 辞退风险金
				if (j == 15) {
					if (!StringUtils.isEmpty(cell)) {
						pay.put("fengxian_money", cell);
					}
				}
				// 扣风险金
				if (j == 16) {
					pay.put("fengxian_money_fine", cell);
				}
				// 意外险
				if (j == 17) {
					pay.put("accident_insurance", cell);
				}
				// 社保
				if (j == 18) {
					pay.put("social_security", cell);
				}
				// 借支
				if (j == 19) {
					pay.put("borrow_money", cell);
				}
				// 油耗转接
				if (j == 20) {
					pay.put("oil_wear_turn", cell);
				}
				// 扣工作服
				if (j == 21) {
					pay.put("work_cloths_money", cell);
				}
				// 调整上月
				if (j == 22) {
					pay.put("last_month", cell);
				}
				// 个人所得税
				if (j == 23) {
					pay.put("pit", cell);
				}
				// 实付工资
				if (j == 24) {
					pay.put("pay_money", cell);
				}
			}
		}
		return payRecord;
	}

	private Map<String, List<Map<String, String>>> readNewNameFile(File newNameFile,
			Map<String, List<Map<String, String>>> person) throws InvalidFormatException, IOException {
		// Map<String, List<Map<String, String>>> person = new HashMap<>();
		List<List<List<String>>> readExcel = ExcelReadUtil.readExcel(newNameFile, new int[] { 2, 2 });
		for (int j = 0; j < readExcel.size(); j++) {
			List<List<String>> sheet = readExcel.get(j);
			// -----离职
			if (j == 0) {
				for (List<String> row : sheet) {
					Map<String, String> pseronInof = new HashMap<>();
					for (int i = 0; i < row.size(); i++) {
						String cell = row.get(i);
						if (i == 1) {
							if (StringUtils.isEmpty(cell)) {
								break;
							}
							List<Map<String, String>> list = person.get(cell);
							if (list == null) {
								list = new ArrayList<>();
								person.put(cell, list);
							}
							pseronInof.put("name", cell);
							list.add(pseronInof);
						}
						if (i == 3) {
							pseronInof.put("status", cell);
						}
						if (i == 11) {
							pseronInof.put("card_no", cell);
						}
						if (i == 13) {
							pseronInof.put("bank", cell);
						}
						if (i == 14) {
							pseronInof.put("card_from", cell);
						}
					}
				}
			}

			// -----在职
			if (j == 1) {
				for (List<String> row : sheet) {
					Map<String, String> pseronInof = new HashMap<>();
					for (int i = 0; i < row.size(); i++) {
						String cell = row.get(i);
						if (i == 4) {
							if (StringUtils.isEmpty(cell)) {
								break;
							}
							List<Map<String, String>> list = person.get(cell);
							if (list == null) {
								list = new ArrayList<>();
								person.put(cell, list);
							}
							pseronInof.put("name", cell);
							list.add(pseronInof);
						}
						if (i == 6) {
							pseronInof.put("status", cell);
						}
						if (i == 14) {
							pseronInof.put("card_no", cell);
						}
						if (i == 16) {
							pseronInof.put("bank", cell);
						}
						if (i == 17) {
							pseronInof.put("card_from", cell);
						}
					}
				}
			}
		}
		return person;
	}

	private File transferTo(MultipartFile sourceFile, String path, long nowTime)
			throws IllegalStateException, IOException {
		if (sourceFile == null) {
			return null;
		}
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
		String filePath = path + File.separator + dateFormat.format(new Date()) + File.separator;
		File neFile = new File(filePath);
		if (!neFile.exists()) {
			neFile.mkdirs();
		}
		neFile = new File(filePath + nowTime + "_" + sourceFile.getOriginalFilename());
		sourceFile.transferTo(neFile);
		return neFile;
	}

}
