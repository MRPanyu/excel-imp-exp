package com.github.mrpanyu.excel;

import java.util.Date;

import org.hibernate.validator.constraints.NotBlank;

@ExcelSheet(name = "用户信息")
@SuppressWarnings("serial")
public class DemoUserExcelModel extends ExcelModelBase {

	@ExcelColumn(name = "用户工号", notes = "不同用户的工号不能重复", width = 20)
	@NotBlank(message = "用户工号不能为空")
	private String userCode;
	@ExcelColumn(name = "用户姓名", width = 10)
	@NotBlank(message = "用户姓名不能为空")
	private String userName;
	@ExcelColumn(name = "身份证号", width = 20)
	@NotBlank(message = "身份证号不能为空")
	private String idcardNo;
	@ExcelColumn(name = "手机号", width = 15)
	@NotBlank(message = "手机号不能为空")
	private String mobile;
	@ExcelColumn(name = "年龄", width = 10)
	private int age;
	@ExcelColumn(name = "出生日期", notes = "用户出生日期，请使用yyyy-MM-dd格式", width = 15, dateFormat = "yyyy-MM-dd")
	private Date birthday;
	@ExcelColumn(name = "性别", width = 6, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "gender")
	private String gender;
	@ExcelColumn(name = "职业类型", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "jobType")
	private String jobType;
	@ExcelColumn(name = "省份", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "province")
	private String homeProvince;
	@ExcelColumn(name = "城市", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "city", selectionRefField = "homeProvince")
	private String homeCity;

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

	public String getIdcardNo() {
		return idcardNo;
	}

	public void setIdcardNo(String idcardNo) {
		this.idcardNo = idcardNo;
	}

	public String getMobile() {
		return mobile;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	public String getJobType() {
		return jobType;
	}

	public void setJobType(String jobType) {
		this.jobType = jobType;
	}

	public String getHomeProvince() {
		return homeProvince;
	}

	public void setHomeProvince(String homeProvince) {
		this.homeProvince = homeProvince;
	}

	public String getHomeCity() {
		return homeCity;
	}

	public void setHomeCity(String homeCity) {
		this.homeCity = homeCity;
	}

	@Override
	public String toString() {
		return "DemoUserExcelModel [userCode=" + userCode + ", userName=" + userName + ", idcardNo=" + idcardNo
				+ ", mobile=" + mobile + ", age=" + age + ", birthday=" + birthday + ", gender=" + gender + ", jobType="
				+ jobType + ", homeProvince=" + homeProvince + ", homeCity=" + homeCity + "]";
	}

}
