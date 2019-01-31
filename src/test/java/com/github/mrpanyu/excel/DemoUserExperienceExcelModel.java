package com.github.mrpanyu.excel;

import java.util.Date;

import javax.validation.constraints.NotNull;

import org.hibernate.validator.constraints.NotBlank;

@ExcelSheet(name = "工作经历")
@SuppressWarnings("serial")
public class DemoUserExperienceExcelModel extends ExcelModelBase {

	@ExcelColumn(name = "用户工号", notes = "不同用户的工号不能重复", width = 20)
	@NotBlank(message = "用户工号不能为空")
	private String userCode;
	@ExcelColumn(name = "用户姓名", width = 10)
	@NotBlank(message = "用户姓名不能为空")
	private String userName;
	@ExcelColumn(name = "开始日期", notes = "请使用yyyy-MM-dd格式", width = 15, dateFormat = "yyyy-MM-dd")
	@NotNull(message = "开始日期不能为空")
	private Date startDate;
	@ExcelColumn(name = "结束日期", notes = "请使用yyyy-MM-dd格式", width = 15, dateFormat = "yyyy-MM-dd")
	private Date endDate;
	@ExcelColumn(name = "所在公司", width = 20)
	@NotBlank(message = "所在公司不能为空")
	private String company;
	@ExcelColumn(name = "经历简介", width = 50)
	private String experience;

	public String getUserCode() {
		return userCode;
	}

	public void setUserCode(String userCode) {
		this.userCode = userCode;
	}

	public String getUserName() {
		return userName;
	}

	public void setUserName(String userName) {
		this.userName = userName;
	}

	public Date getStartDate() {
		return startDate;
	}

	public void setStartDate(Date startDate) {
		this.startDate = startDate;
	}

	public Date getEndDate() {
		return endDate;
	}

	public void setEndDate(Date endDate) {
		this.endDate = endDate;
	}

	public String getCompany() {
		return company;
	}

	public void setCompany(String company) {
		this.company = company;
	}

	public String getExperience() {
		return experience;
	}

	public void setExperience(String experience) {
		this.experience = experience;
	}

	@Override
	public String toString() {
		return "DemoUserExperienceExcelModel [userCode=" + userCode + ", userName=" + userName + ", startDate="
				+ startDate + ", endDate=" + endDate + ", company=" + company + ", experience=" + experience + "]";
	}

}
