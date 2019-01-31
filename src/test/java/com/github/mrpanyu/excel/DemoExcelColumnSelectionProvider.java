package com.github.mrpanyu.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * 示例用的一个下拉框选项提供类。实际使用过程中一般会从数据库等获取下拉选项。
 */
public class DemoExcelColumnSelectionProvider implements ExcelColumnSelectionProvider {

	@Override
	public List<ExcelColumnSelectionItem> selectionItems(String type) {
		List<ExcelColumnSelectionItem> items = new ArrayList<ExcelColumnSelectionItem>();
		if ("gender".equals(type)) {
			items.add(new ExcelColumnSelectionItem("0", "女"));
			items.add(new ExcelColumnSelectionItem("1", "男"));
		} else if ("jobType".equals(type)) {
			items.add(new ExcelColumnSelectionItem("01", "管理人员"));
			items.add(new ExcelColumnSelectionItem("02", "现场工人"));
			items.add(new ExcelColumnSelectionItem("03", "后勤人员"));
			items.add(new ExcelColumnSelectionItem("99", "其他人员"));
		} else if ("province".equals(type)) {
			items.add(new ExcelColumnSelectionItem("110000", "北京市"));
			items.add(new ExcelColumnSelectionItem("120000", "天津市"));
			items.add(new ExcelColumnSelectionItem("130000", "河北省"));
			items.add(new ExcelColumnSelectionItem("140000", "山西省"));
		} else if ("city".equals(type)) { // 城市要与省份联动，因此要包含引用的省份信息
			items.add(new ExcelColumnSelectionItem("110100", "北京市", "110000", "北京市"));
			items.add(new ExcelColumnSelectionItem("120100", "天津市", "120000", "天津市"));
			items.add(new ExcelColumnSelectionItem("130100", "石家庄市", "130000", "河北省"));
			items.add(new ExcelColumnSelectionItem("130200", "唐山市", "130000", "河北省"));
			items.add(new ExcelColumnSelectionItem("130300", "秦皇岛市", "130000", "河北省"));
			items.add(new ExcelColumnSelectionItem("140100", "太原市", "140000", "山西省"));
			items.add(new ExcelColumnSelectionItem("140200", "大同市", "140000", "山西省"));
			items.add(new ExcelColumnSelectionItem("140300", "阳泉市", "140000", "山西省"));
		}
		return items;
	}

}
