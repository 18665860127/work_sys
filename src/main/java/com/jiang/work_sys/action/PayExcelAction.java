package com.jiang.work_sys.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSON;
import com.jiang.work_sys.util.excel.ExcelUtil;

@Controller("payExcelAction")
@RequestMapping("/a")
public class PayExcelAction {

	@RequestMapping("a")
	public String uploadPage() {
		return "a/uploadPage";
	}

	private final static String fileUrl = "d:/tempExcelFile/";

	@RequestMapping("b")
	public void uploadPayExcelChange(HttpServletRequest req, HttpServletResponse rep,
			@RequestParam("nameFile") MultipartFile  nameFile, @RequestParam("payFile") MultipartFile  payFile) throws IOException {
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
		File newNameFile = new File(fileUrl + dateFormat.format(new Date()) + nameFile.getName());
		if (!newNameFile.getParentFile().exists()) {
			newNameFile.mkdirs();
		}
		nameFile.transferTo(newNameFile);
		System.out.println();
	}
}
