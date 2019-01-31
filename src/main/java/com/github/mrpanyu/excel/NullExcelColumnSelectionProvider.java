package com.github.mrpanyu.excel;

import java.util.List;

/**
 * 默认的列表提供者，不提供可选列表
 */
public class NullExcelColumnSelectionProvider implements ExcelColumnSelectionProvider {

	@Override
	public List<ExcelColumnSelectionItem> selectionItems(String type) {
		return null;
	}

}
