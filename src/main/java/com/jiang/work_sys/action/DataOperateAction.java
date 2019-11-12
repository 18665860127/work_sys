package com.jiang.work_sys.action;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Random;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;

import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@CrossOrigin(origins = "*", maxAge = 3600)
@Controller("dataOperate")
@RequestMapping("/dataOperate")
public class DataOperateAction {

	static List<Map<String, String>> studentList = Collections.synchronizedList(new ArrayList<>());
	static Map<String, Map<String, String>> studentMap = Collections.synchronizedMap(new HashMap<>());
	static {
		for (int i = 0; i < 10; i++) {
			Map<String, String> student = new HashMap<>();
			student.put("studentId", UUID.randomUUID().toString().replaceAll("-", ""));
			student.put("studentName", "学生" + (i + 1));
			student.put("studentAge", new Random().nextInt(50) + "");
			studentList.add(student);
			studentMap.put(student.get("studentId"), student);
		}
	}

	@RequestMapping("add")
	@ResponseBody
	public Map<String, Object> add(HttpServletRequest req) {
		Map<String, Object> resultMap = new HashMap<String, Object>();
		try {
			Map<String, String[]> parameterMap = req.getParameterMap();
			String studentName = parameterMap.get("studentName")[0];
			String studentAge = parameterMap.get("studentAge")[0];
			try {
				Long.valueOf(studentAge);
			} catch (Exception e) {
				resultMap.put("isSuccess", "0");
				resultMap.put("msg", "更新失败，年龄只能为数字");
				return resultMap;
			}
			Map<String, String> student = new HashMap<>();
			student.put("studentId", UUID.randomUUID().toString().replaceAll("-", ""));
			student.put("studentName", studentName);
			student.put("studentAge", studentAge);
			studentList.add(student);
			studentMap.put(student.get("studentId"), student);
			resultMap.put("isSuccess", "1");
			resultMap.put("msg", "新增成功");
			return resultMap;
		} catch (Exception e) {
			resultMap.put("isSuccess", "0");
			resultMap.put("msg", "更新失败，系统错误");
			return resultMap;
		}
	}

	@RequestMapping("del")
	@ResponseBody
	public Map<String, Object> del(HttpServletRequest req) {
		Map<String, Object> resultMap = new HashMap<String, Object>();
		Map<String, String[]> parameterMap = req.getParameterMap();
		String studentId = parameterMap.get("studentId")[0];
		if (Objects.isNull(studentMap.get(studentId))) {
			resultMap.put("isSuccess", "0");
			resultMap.put("msg", "删除失败，学生不存在");
			return resultMap;
		}
		studentMap.remove(studentId);
		Iterator<Map<String, String>> iterator = studentList.iterator();
		while (iterator.hasNext()) {
			Map<String, String> student = iterator.next();
			if (studentId.equals(student.get("studentId"))) {
				iterator.remove();
			}
		}
		resultMap.put("isSuccess", "1");
		resultMap.put("msg", "删除成功");
		return resultMap;
	}

	@RequestMapping("edit")
	@ResponseBody
	public Map<String, Object> edit(HttpServletRequest req) {
		Map<String, Object> resultMap = new HashMap<String, Object>();
		try {
			Map<String, String[]> parameterMap = req.getParameterMap();
			String studentName = parameterMap.get("studentName")[0];
			String studentAge = parameterMap.get("studentAge")[0];
			String studentId = parameterMap.get("studentId")[0];
			Map<String, String> student = studentMap.get(studentId);
			if (Objects.isNull(student)) {
				resultMap.put("isSuccess", "0");
				resultMap.put("msg", "修改失败，学生不存在");
				return resultMap;
			}
			if (!StringUtils.isEmpty(studentAge)) {
				try {
					Long.valueOf(studentAge);
				} catch (Exception e) {
					resultMap.put("isSuccess", "0");
					resultMap.put("msg", "更新失败，年龄只能为数字");
					return resultMap;
				}
				student.put("studentAge", studentAge);
			}
			if (!StringUtils.isEmpty(studentName)) {
				student.put("studentName", studentName);
			}
			resultMap.put("isSuccess", "1");
			resultMap.put("msg", " 成功");
			return resultMap;
		} catch (Exception e) {
			resultMap.put("isSuccess", "0");
			resultMap.put("msg", "更新失败，系统错误");
			return resultMap;
		}
	}

	@RequestMapping("query")
	@ResponseBody
	public Map<String, Object> query(HttpServletRequest req) {
		Map<String, Object> resultMap = new HashMap<String, Object>();
		resultMap.put("students", studentList);
		return resultMap;
	}
}
