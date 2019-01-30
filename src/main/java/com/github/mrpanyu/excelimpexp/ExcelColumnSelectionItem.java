package com.github.mrpanyu.excelimpexp;

import java.io.Serializable;

/**
 * Excel模型某个列的可选项值
 */
@SuppressWarnings("serial")
public class ExcelColumnSelectionItem implements Serializable {

	/** 实际值 */
	private String value;
	/** 显示名称 */
	private String name;

	/** 联动上级实际值，用于级联选择的场景 */
	private String refValue;
	/** 联动上级显示名称，用于级联选择的场景 */
	private String refName;

	public ExcelColumnSelectionItem() {
	}

	public ExcelColumnSelectionItem(String value, String name) {
		this.value = value;
		this.name = name;
	}

	public ExcelColumnSelectionItem(String value, String name, String refValue, String refName) {
		this.value = value;
		this.name = name;
		this.refValue = refValue;
		this.refName = refName;
	}

	protected String getValue() {
		return value;
	}

	protected void setValue(String value) {
		this.value = value;
	}

	protected String getName() {
		return name;
	}

	protected void setName(String name) {
		this.name = name;
	}

	protected String getRefValue() {
		return refValue;
	}

	protected void setRefValue(String refValue) {
		this.refValue = refValue;
	}

	protected String getRefName() {
		return refName;
	}

	protected void setRefName(String refName) {
		this.refName = refName;
	}

}
