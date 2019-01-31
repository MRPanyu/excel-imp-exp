package com.github.mrpanyu.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.Test;

public class ExcelImportExportToolTest {

	/** 生成导入模板示例 */
	@Test
	public void testGenerateTemplate() throws Exception {
		// 生成两个模型对象（分别对应两个sheet页）的Excel导入模板文件
		byte[] data = ExcelImportExportTools.impTemplate(DemoUserExcelModel.class, DemoUserExperienceExcelModel.class);
		// 输出到本地文件
		writeToFile("test-import-template.xlsx", data);
	}

	/** 导入示例 */
	@Test
	public void testImport() throws Exception {
		// 读取Excel文件
		InputStream in = ExcelImportExportToolTest.class.getResourceAsStream("test-import.xlsx");
		try {
			// 根据模型定义导入Excel
			List<List<Object>> models = ExcelImportExportTools.imp(in, DemoUserExcelModel.class,
					DemoUserExperienceExcelModel.class);
			List<Object> sheet1Models = models.get(0);
			List<Object> sheet2Models = models.get(1);
			for (Object model : sheet1Models) {
				DemoUserExcelModel user = (DemoUserExcelModel) model;
				System.out.println(user);
				// hasError表示导入过程本身是否这行解析有错，包括解析本身的问题（如数字/日期无法解析，代码翻译错误等），以及@NotBlank等校验标注的检查
				System.out.println("hasError=" + user.hasError());
				// hasFieldError表示某个导入字段是否存在问题
				System.out.println("hasFieldErrorOnUserName=" + user.hasFieldError("userName"));
				// 也可以自己增加错误，供之后导出错误模板
				if ("100004".equals(user.getUserCode())) {
					// 如果具体某个导入字段存在问题，进行addFieldError，导出错误信息时会标红
					user.addFieldError("idcardNo", "身份证不正确");
					// 与字段无关的问题，addOtherError，在最后一列会显示该信息
					user.addOtherError("该工号的人员已存在，不能重复导入");
				}
			}
			for (Object model : sheet2Models) {
				DemoUserExperienceExcelModel userExperience = (DemoUserExperienceExcelModel) model;
				System.out.println(userExperience);
			}
			// 还是通过Excel文件，导出错误信息
			// 导出的模型对象中设置过错误信息的，不论是导入时自动产生的还是后续人工添加的，导出后都有提现
			byte[] errorExportData = ExcelImportExportTools.exp(models, DemoUserExcelModel.class,
					DemoUserExperienceExcelModel.class);
			writeToFile("test-error.xlsx", errorExportData);
		} finally {
			in.close();
		}
	}

	/** 导出示例 */
	@Test
	public void testExport() throws Exception {
		// 要导出的对象
		List<List<Object>> models = new ArrayList<List<Object>>();
		List<Object> sheet1Models = new ArrayList<Object>();
		List<Object> sheet2Models = new ArrayList<Object>();
		models.add(sheet1Models);
		models.add(sheet2Models);

		DemoUserExcelModel user = new DemoUserExcelModel();
		user.setUserCode("100005");
		user.setUserName("孙明");
		user.setIdcardNo("120102199205200003");
		user.setMobile("13172727272");
		user.setBirthday(parseDate("1992-06-20"));
		user.setAge(27);
		user.setGender("0"); // 注意对象中是代码值，不是名称
		user.setJobType("02");
		user.setHomeProvince("120000");
		user.setHomeCity("120100");
		sheet1Models.add(user);

		DemoUserExperienceExcelModel userExperience = new DemoUserExperienceExcelModel();
		userExperience.setUserCode("100005");
		userExperience.setUserName("孙明");
		userExperience.setCompany("天津国际交易中心");
		userExperience.setStartDate(parseDate("2012-12-25"));
		userExperience.setExperience("实习工作");
		sheet2Models.add(userExperience);

		// 导出
		byte[] exportData = ExcelImportExportTools.exp(models, DemoUserExcelModel.class,
				DemoUserExperienceExcelModel.class);
		writeToFile("test-export.xlsx", exportData);
	}

	private void writeToFile(String fileName, byte[] data) throws IOException {
		FileOutputStream fout = new FileOutputStream(fileName);
		try {
			fout.write(data);
		} finally {
			fout.close();
		}
	}

	private Date parseDate(String str) throws ParseException {
		return new SimpleDateFormat("yyyy-MM-dd").parse(str);
	}

}
