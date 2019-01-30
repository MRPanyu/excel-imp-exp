package com.github.mrpanyu.excelimpexp;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel模型对象的基类，包含一个特殊的错误信息字段，如果该错误信息有值，导出的时候最后会有一个错误列显示错误信息
 */
@SuppressWarnings("serial")
public class ExcelModelBase implements Serializable {

	/** 其他错误信息 */
	private List<String> otherErrors = new ArrayList<String>();
	/** 字段相关的错误信息 */
	private Map<String, List<String>> fieldErrorMap = new LinkedHashMap<String, List<String>>();
	/** 导入时读取的原始值 */
	private Map<String, String> originalValueMap = new HashMap<String, String>();

	/** 获取所有错误信息 */
	public List<String> getAllErrors() {
		List<String> allErrors = new ArrayList<String>();
		for (List<String> fieldErrors : fieldErrorMap.values()) {
			allErrors.addAll(fieldErrors);
		}
		allErrors.addAll(otherErrors);
		return allErrors;
	}

	/** 增加字段相关错误 */
	public void addFieldError(String fieldName, String error) {
		List<String> fieldErrors = fieldErrorMap.get(fieldName);
		if (fieldErrors == null) {
			fieldErrors = new ArrayList<String>();
			fieldErrorMap.put(fieldName, fieldErrors);
		}
		fieldErrors.add(error);
	}

	/** 增加一条错误信息 */
	public void addOtherError(String error) {
		this.otherErrors.add(error);
	}

	/** 某字段是否有错误 */
	public boolean hasFieldError(String fieldName) {
		return fieldErrorMap.containsKey(fieldName);
	}

	/** 是否有错误 */
	public boolean hasError() {
		return !this.fieldErrorMap.isEmpty() || !this.otherErrors.isEmpty();
	}

	/** 设置Excel导入时的原始值 */
	public void setOriginalValue(String key, String value) {
		originalValueMap.put(key, value);
	}

	/** 获取Excel导入时的原始值 */
	public String getOriginalValue(String key) {
		return originalValueMap.get(key);
	}

}
