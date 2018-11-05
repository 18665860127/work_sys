package com.jiang.work_sys.action;

import java.io.File;
import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

@Controller("payExcelAction")
@RequestMapping("/a")
public class PayExcelAction {

	@RequestMapping("a")
	public String uploadPage() {
		return "a/uploadPage";
	}

	@RequestMapping("b")
	public void uploadPayExcelChange(HttpServletRequest req, HttpServletResponse rep, @RequestParam("file") File file)
			throws IOException {
		System.out.println("aaaa");
		rep.getWriter().write("d12a");
		rep.getWriter().flush();
		rep.getWriter().close();
	}
}
