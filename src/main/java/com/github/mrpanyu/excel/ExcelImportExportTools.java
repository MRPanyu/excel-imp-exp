package com.github.mrpanyu.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.validator.HibernateValidator;

/**
 * Excel导入导出工具
 */
public class ExcelImportExportTools {

	/** 导入模板填入多少行 */
	private static final int TEMPLATE_DATA_ROWS = 100;

	private static Map<Class<?>, ExcelColumnSelectionProvider> selectionProviderMap = new HashMap<Class<?>, ExcelColumnSelectionProvider>();

	private static ValidatorFactory validatorFactory = Validation.byProvider(HibernateValidator.class).configure()
			.buildValidatorFactory();

	/**
	 * 生成导入模板文件
	 * 
	 * @param modelClasses 导入模型信息，每个模型针对一个sheet页
	 * @return 导入模板文件内容
	 */
	public static byte[] impTemplate(Class<?>... modelClasses) {
		// 导入模板实际就是导出若干个空对象
		List<List<Object>> data = new ArrayList<List<Object>>(modelClasses.length);
		for (int i = 0; i < modelClasses.length; i++) {
			List<Object> sheetData = Arrays.asList(new Object[TEMPLATE_DATA_ROWS]);
			data.add(sheetData);
		}
		return exp(data, modelClasses);
	}
	
	/**
	 * 将导入的Excel转换为模型对象
	 * 
	 * @param excelInput    Excel文件输入流
	 * @param modelClasses 导入模型信息，每个模型针对一个sheet页，如果有某个sheet页无需导入，可以传一个null值表示跳过
	 * @return 转换后的模型对象，外侧List每个元素针对一个sheet页（无需导入的也会有个null值），内侧元素的List表示每个sheet页转换成的模型数据
	 */
	public static List<List<Object>> imp(InputStream excelInput, Class<?>... modelClasses) {
		try {
			List<List<Object>> result = new ArrayList<List<Object>>();
			Workbook wb = new XSSFWorkbook(excelInput);
			for (int i = 0; i < modelClasses.length; i++) {
				if (modelClasses[i] == null) {
					result.add(null);
				} else {
					Sheet sheet = wb.getSheetAt(i);
					List<Object> list = importSheet(wb, sheet, modelClasses[i]);
					result.add(list);
				}
			}
			return result;
		} catch (RuntimeException e) {
			throw e;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 将导入的Excel转换为模型对象
	 * 
	 * @param excelFile    Excel文件内容
	 * @param modelClasses 导入模型信息，每个模型针对一个sheet页，如果有某个sheet页无需导入，可以传一个null值表示跳过
	 * @return 转换后的模型对象，外侧List每个元素针对一个sheet页（无需导入的也会有个null值），内侧元素的List表示每个sheet页转换成的模型数据
	 */
	public static List<List<Object>> imp(byte[] excelFile, Class<?>... modelClasses) {
		return imp(new ByteArrayInputStream(excelFile), modelClasses);
	}

	/**
	 * 将模型对象导出成Excel文件
	 * 
	 * @param data         模型对象，外侧List每个元素针对一个sheet页，内侧元素的List表示每个sheet页中的模型数据
	 * @param modelClasses 导出的模型信息，每个模型针对一个sheet页
	 * @return 导出的Excel文件内容
	 */
	public static byte[] exp(List<List<Object>> data, Class<?>... modelClasses) {
		try {
			if (data == null) {
				throw new IllegalArgumentException("data不能为null");
			}
			if (modelClasses == null) {
				throw new IllegalArgumentException("modelClasses不能为null");
			}
			if (data.size() != modelClasses.length) {
				throw new IllegalArgumentException("modelClasses个数必须与data对应");
			}
			Workbook wb = new XSSFWorkbook();
			for (Class<?> modelClass : modelClasses) {
				makeExportSheet(wb, modelClass);
			}
			int i = 0;
			for (Class<?> modelClass : modelClasses) {
				exportSheet(wb, wb.getSheetAt(i), modelClass, data.get(i));
				i++;
			}
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			wb.write(baos);
			return baos.toByteArray();
		} catch (RuntimeException e) {
			throw e;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/** 创建导出sheet页 */
	private static Sheet makeExportSheet(Workbook wb, Class<?> modelClass) throws Exception {
		ExcelSheet sheetInfo = modelClass.getAnnotation(ExcelSheet.class);
		if (sheetInfo == null) {
			throw new IllegalArgumentException("Excel模型类" + modelClass.getName() + "未包含@ExcelSheet标注信息");
		}
		String sheetName = sheetInfo.name();
		Sheet sheet = wb.createSheet(sheetName);
		return sheet;
	}

	/** 导入模板单sheet处理 */
	private static void exportSheet(Workbook wb, Sheet sheet, Class<?> modelClass, List<Object> data) throws Exception {
		List<ExcelColumnInfo> columnInfoList = getColumnInfos(modelClass);
		// 样式
		CellStyle[] headerCellStyles = new CellStyle[columnInfoList.size()];
		CellStyle[] dataCellStyles = new CellStyle[columnInfoList.size()];
		CellStyle[] errorRowNormalCellStyles = new CellStyle[columnInfoList.size()];
		CellStyle[] errorRowErrorCellStyles = new CellStyle[columnInfoList.size()];
		CellStyle errorMessageCellStyle = getErrorMessageCellStyle(wb);
		int columnNum = 0;
		for (ExcelColumnInfo columnInfo : columnInfoList) {
			headerCellStyles[columnNum] = getHeaderCellStyle(wb, columnInfo);
			dataCellStyles[columnNum] = getCellStyle(wb, columnInfo);
			errorRowNormalCellStyles[columnNum] = getErrorRowNormalCellStyle(wb, columnInfo);
			errorRowErrorCellStyles[columnNum] = getErrorRowErrorCellStyle(wb, columnInfo);
			columnNum++;
		}

		// 表头
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		CreationHelper creationHelper = wb.getCreationHelper();
		Row headerRow = sheet.createRow(0);
		columnNum = 0;
		for (ExcelColumnInfo columnInfo : columnInfoList) {
			Cell cell = headerRow.createCell(columnNum, CellType.STRING);
			cell.setCellValue(columnInfo.annotation.name());
			cell.setCellStyle(headerCellStyles[columnNum]);
			if (Utils.isNotBlank(columnInfo.annotation.notes())) {
				ClientAnchor anchor = creationHelper.createClientAnchor();
				anchor.setCol1(cell.getColumnIndex());
				anchor.setCol2(cell.getColumnIndex() + 3);
				anchor.setRow1(cell.getRowIndex() + 1);
				anchor.setRow2(cell.getRowIndex() + 4);
				Comment comment = drawing.createCellComment(anchor);
				comment.setString(creationHelper.createRichTextString(columnInfo.annotation.notes()));
				comment.setVisible(false);
				cell.setCellComment(comment);
			}
			columnNum++;
		}

		// 数据行
		for (int i = 0; i < data.size(); i++) {
			Object item = data.get(i);
			boolean hasError = item != null && item instanceof ExcelModelBase && ((ExcelModelBase) item).hasError();
			Row row = sheet.createRow(i + 1);
			columnNum = 0;
			for (ExcelColumnInfo columnInfo : columnInfoList) {
				String fieldName = columnInfo.field.getName();
				Class<?> fieldType = columnInfo.field.getType();
				Cell cell = null;
				if (hasError && ((ExcelModelBase) item).hasFieldError(columnInfo.field.getName())) {
					cell = row.createCell(columnNum, CellType.STRING);
					cell.setCellStyle(errorRowErrorCellStyles[columnNum]);
					cell.setCellValue(((ExcelModelBase) item).getOriginalValue(fieldName));
				} else {
					if (Number.class.isAssignableFrom(fieldType) || Date.class.isAssignableFrom(fieldType)) {
						cell = row.createCell(columnNum, CellType.NUMERIC);
					} else {
						cell = row.createCell(columnNum, CellType.STRING);
					}
					if (hasError) {
						cell.setCellStyle(errorRowNormalCellStyles[columnNum]);
					} else {
						cell.setCellStyle(dataCellStyles[columnNum]);
					}
					// 设值
					if (item != null) {
						Object cellValue = columnInfo.field.get(item);
						if (cellValue != null) {
							if (cellValue instanceof Date) {
								cell.setCellValue((Date) cellValue);
							} else if (cellValue instanceof Number) {
								cell.setCellValue(((Number) cellValue).doubleValue());
							} else {
								List<ExcelColumnSelectionItem> selectionItems = columnInfo.getSelectionItems();
								if (selectionItems != null) {
									for (ExcelColumnSelectionItem selectionItem : selectionItems) {
										if (Utils.equals(selectionItem.getValue(), cellValue.toString())) {
											cellValue = selectionItem.getName();
											break;
										}
									}
								}
								cell.setCellValue(cellValue.toString());
							}
						}
					}
				}
				columnNum++;
			}
			if (hasError) {
				ExcelModelBase modelBase = (ExcelModelBase) item;
				Cell cell = row.createCell(columnNum, CellType.STRING);
				cell.setCellStyle(errorMessageCellStyle);
				String allErrorMessages = Utils.join(modelBase.getAllErrors(), ";");
				cell.setCellValue(allErrorMessages);
			}
		}

		// 可选项
		columnNum = 0;
		for (ExcelColumnInfo columnInfo : columnInfoList) {
			if (columnInfo.annotation.selectionProvider() != null) {
				ExcelColumnSelectionProvider provider = getExcelColumnSelectionProvider(
						columnInfo.annotation.selectionProvider());
				List<ExcelColumnSelectionItem> items = provider.selectionItems(columnInfo.annotation.selectionType());
				if (items != null) {
					if (Utils.isBlank(columnInfo.annotation.selectionRefField())) {
						// 非级联下拉
						String baseName = createRefSheetWithNames(wb, columnInfo);
						DataValidationConstraint dvc = sheet.getDataValidationHelper()
								.createFormulaListConstraint(baseName);
						CellRangeAddressList range = new CellRangeAddressList(1, data.size(), columnNum, columnNum);
						sheet.addValidationData(sheet.getDataValidationHelper().createValidation(dvc, range));
					} else {
						// 级联查询，需要创建参照sheet
						String baseName = createRefSheetWithNames(wb, columnInfo);
						int refColumnIndex = -1;
						for (int i = 0; i < columnInfoList.size(); i++) {
							if (Utils.equals(columnInfo.annotation.selectionRefField(),
									columnInfoList.get(i).field.getName())) {
								refColumnIndex = i;
								break;
							}
						}
						if (refColumnIndex < 0) {
							throw new RuntimeException(
									"selectionRefField指定的字段" + columnInfo.annotation.selectionRefField() + "不存在");
						}
						String refColumnName = numberToColumnHead(refColumnIndex + 1);
						String formula = "INDIRECT(CONCATENATE(\"" + baseName + "_\",$" + refColumnName + "2))";
						DataValidationConstraint dvc = sheet.getDataValidationHelper()
								.createFormulaListConstraint(formula);
						CellRangeAddressList range = new CellRangeAddressList(1, data.size(), columnNum, columnNum);
						sheet.addValidationData(sheet.getDataValidationHelper().createValidation(dvc, range));
					}
				}
			}
			columnNum++;
		}

		// 宽度设置
		columnNum = 0;
		for (ExcelColumnInfo columnInfo : columnInfoList) {
			if (columnInfo.annotation.width() > 0) {
				sheet.setColumnWidth(columnNum, columnInfo.annotation.width() * 256);
			} else {
				sheet.autoSizeColumn(columnNum);
			}
			columnNum++;
		}
		// 最后一列（错误信息）
		sheet.setColumnWidth(columnNum, 25600);
	}

	/** 导入单sheet处理 */
	private static List<Object> importSheet(Workbook wb, Sheet sheet, Class<?> modelClass) throws Exception {
		List<Object> result = new ArrayList<Object>();
		List<ExcelColumnInfo> columnInfos = getColumnInfos(modelClass);
		for (int r = 1; r <= sheet.getLastRowNum(); r++) {
			Row row = sheet.getRow(r);
			if (row != null) {
				Object model = modelClass.newInstance();
				boolean isAllNull = true;
				for (int c = 0; c < columnInfos.size(); c++) {
					Cell cell = row.getCell(c);
					ExcelColumnInfo columnInfo = columnInfos.get(c);
					if (readCellValueToModel(wb, cell, columnInfo, model)) {
						isAllNull = false;
					}
				}
				if (!isAllNull) {
					// 集成Validator校验
					if (model instanceof ExcelModelBase) {
						ExcelModelBase modelBase = (ExcelModelBase) model;
						Validator validator = validatorFactory.getValidator();
						Set<ConstraintViolation<Object>> violations = validator.validate(model);
						for (ConstraintViolation<Object> violation : violations) {
							String fieldName = violation.getPropertyPath().iterator().next().getName();
							String error = violation.getMessage();
							// 如果已经有解析错误了，先不增加额外的错误信息
							if (!modelBase.hasFieldError(fieldName)) {
								modelBase.addFieldError(fieldName, error);
							}
						}
					}

					result.add(model);
				}
			}
		}
		return result;
	}

	/** 获取模型对象所有标注属性 */
	private static List<ExcelColumnInfo> getColumnInfos(Class<?> modelClass) throws Exception {
		List<ExcelColumnInfo> infoList = new ArrayList<ExcelColumnInfo>();
		Field[] fields = modelClass.getDeclaredFields();
		for (Field field : fields) {
			ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
			if (annotation != null) {
				ExcelColumnInfo info = new ExcelColumnInfo();
				field.setAccessible(true);
				info.annotation = annotation;
				info.field = field;
				infoList.add(info);
			}
		}
		return infoList;
	}

	/** 导入模板及导出列头单元格样式 */
	private static CellStyle getHeaderCellStyle(Workbook wb, ExcelColumnInfo columnInfo) {
		CellStyle style = wb.createCellStyle();
		if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.CENTER) {
			style.setAlignment(HorizontalAlignment.CENTER);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.LEFT) {
			style.setAlignment(HorizontalAlignment.LEFT);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.RIGHT) {
			style.setAlignment(HorizontalAlignment.RIGHT);
		}
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = wb.createFont();
		font.setBold(true);
		style.setFont(font);
		return style;
	}

	/** 导入模板及导出数据单元格样式 */
	@SuppressWarnings("unchecked")
	private static CellStyle getCellStyle(Workbook wb, ExcelColumnInfo columnInfo) {
		CellStyle style = wb.createCellStyle();
		if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.CENTER) {
			style.setAlignment(HorizontalAlignment.CENTER);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.LEFT) {
			style.setAlignment(HorizontalAlignment.LEFT);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.RIGHT) {
			style.setAlignment(HorizontalAlignment.RIGHT);
		}
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		if (Utils.isNotBlank(columnInfo.annotation.dateFormat())) {
			// 自定义日期格式
			style.setDataFormat(
					wb.getCreationHelper().createDataFormat().getFormat(columnInfo.annotation.dateFormat()));
		} else if (String.class.equals(columnInfo.field.getType())) {
			// 文本
			style.setDataFormat((short) BuiltinFormats.getBuiltinFormat("@"));
		} else if (Date.class.equals(columnInfo.field.getType())) {
			// 日期
			style.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd"));
		} else if (Arrays.asList(Integer.class, Integer.TYPE, Long.class, Long.TYPE)
				.contains(columnInfo.field.getType())) {
			style.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0"));
		} else if (Arrays.asList(Double.class, Double.TYPE, Float.class, Float.TYPE)
				.contains(columnInfo.field.getType())) {
			style.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00"));
		} else {
			// 常规
			style.setDataFormat((short) 0);
		}
		return style;
	}

	/** 导出时错误行正常单元格样式 */
	private static CellStyle getErrorRowNormalCellStyle(Workbook wb, ExcelColumnInfo columnInfo) {
		CellStyle style = getCellStyle(wb, columnInfo);
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index);
		style.setFillPattern(FillPatternType.FINE_DOTS);
		return style;
	}

	/** 导出时错误行错误单元格样式 */
	private static CellStyle getErrorRowErrorCellStyle(Workbook wb, ExcelColumnInfo columnInfo) {
		CellStyle style = wb.createCellStyle();
		if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.CENTER) {
			style.setAlignment(HorizontalAlignment.CENTER);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.LEFT) {
			style.setAlignment(HorizontalAlignment.LEFT);
		} else if (columnInfo.annotation.horizontalAlignment() == ExcelColumnHorizontalAlignment.RIGHT) {
			style.setAlignment(HorizontalAlignment.RIGHT);
		}
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		Font font = wb.createFont();
		font.setBold(true);
		font.setColor(IndexedColors.RED.getIndex());
		style.setFont(font);
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index);
		style.setFillPattern(FillPatternType.FINE_DOTS);
		return style;
	}

	/** 导出时显示错误信息单元格样式 */
	private static CellStyle getErrorMessageCellStyle(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		Font font = wb.createFont();
		font.setBold(true);
		font.setColor(IndexedColors.RED.getIndex());
		style.setFont(font);
		return style;
	}

	/** 获取列可选项提供者 */
	private static ExcelColumnSelectionProvider getExcelColumnSelectionProvider(
			Class<? extends ExcelColumnSelectionProvider> selectionProviderClass) {
		ExcelColumnSelectionProvider selectionProvider = selectionProviderMap.get(selectionProviderClass);
		if (selectionProvider == null) {
			synchronized (selectionProviderMap) {
				selectionProvider = selectionProviderMap.get(selectionProviderClass);
				if (selectionProvider == null) {
					// 优先从Spring容器中获取同类型的Bean
					/*
					ApplicationContext applicationContext = SpringContextUtils.getApplicationContext();
					if (applicationContext != null) {
						try {
							selectionProvider = applicationContext.getBean(selectionProviderClass);
						} catch (Exception e) {
							// ignore
						}
					}
					*/
					// 如果获取不到，直接用new
					if (selectionProvider == null) {
						try {
							selectionProvider = selectionProviderClass.newInstance();
						} catch (Exception e) {
							throw new RuntimeException(e);
						}
					}
					selectionProviderMap.put(selectionProviderClass, selectionProvider);
				}
			}
		}
		return selectionProvider;
	}

	/** 创建关联sheet页（隐藏）及相关的“名称”（Excel名称管理器中的那种） */
	private static String createRefSheetWithNames(Workbook wb, ExcelColumnInfo columnInfo) {
		String baseName = "REF_" + columnInfo.annotation.selectionProvider().getSimpleName() + "_"
				+ columnInfo.annotation.selectionType();
		if (wb.getName(baseName) == null) {
			// 创建REF sheet
			Sheet sheet = wb.getSheet("REF");
			if (sheet == null) {
				sheet = wb.createSheet("REF");
				wb.setSheetHidden(wb.getSheetIndex(sheet), true);
			}

			// 创建全集引用Name
			int rowNum = sheet.getLastRowNum() + 1;
			Row row = sheet.createRow(rowNum);
			int cellNum = 0;
			for (ExcelColumnSelectionItem item : columnInfo.getSelectionItems()) {
				Cell cell = row.createCell(cellNum, CellType.STRING);
				cell.setCellValue(item.getName());
				cellNum++;
			}
			Name wbName = wb.createName();
			wbName.setNameName(baseName);
			wbName.setRefersToFormula(
					"REF!" + "$A$" + (rowNum + 1) + ":$" + numberToColumnHead(cellNum) + "$" + (rowNum + 1));
			rowNum++;

			// 创建级联引用的Name
			Map<String, List<ExcelColumnSelectionItem>> map = new LinkedHashMap<String, List<ExcelColumnSelectionItem>>();
			for (ExcelColumnSelectionItem selectionItem : columnInfo.getSelectionItems()) {
				String refName = selectionItem.getRefName();
				if (Utils.isNotBlank(refName)) {
					List<ExcelColumnSelectionItem> list = map.get(refName);
					if (list == null) {
						list = new ArrayList<ExcelColumnSelectionItem>();
						map.put(refName, list);
					}
					list.add(selectionItem);
				}
			}
			for (Map.Entry<String, List<ExcelColumnSelectionItem>> entry : map.entrySet()) {
				String refName = entry.getKey();
				List<ExcelColumnSelectionItem> list = entry.getValue();
				row = sheet.createRow(rowNum);
				cellNum = 0;
				for (ExcelColumnSelectionItem item : list) {
					Cell cell = row.createCell(cellNum, CellType.STRING);
					cell.setCellValue(item.getName());
					cellNum++;
				}
				wbName = wb.createName();
				wbName.setNameName(baseName + "_" + refName);
				wbName.setRefersToFormula(
						"REF!" + "$A$" + (rowNum + 1) + ":$" + numberToColumnHead(cellNum) + "$" + (rowNum + 1));
				rowNum++;
			}
		}
		return baseName;
	}

	/** 数字转换为Excel列头格式，如1转换为A，2转换为B，27转换为AA等 */
	private static String numberToColumnHead(int num) {
		if (num <= 26) {
			return String.valueOf((char) ('A' + (num - 1)));
		} else {
			int h = (num - 1) / 26;
			int l = (num - 1) % 26 + 1;
			return numberToColumnHead(h) + numberToColumnHead(l);
		}
	}

	/** 读取单元格的值，返回是否读到非空值 */
	private static boolean readCellValueToModel(Workbook wb, Cell cell, ExcelColumnInfo columnInfo, Object model) {
		Class<?> fieldType = columnInfo.field.getType();
		ExcelModelBase modelBase = new ExcelModelBase();
		if (model instanceof ExcelModelBase) {
			modelBase = (ExcelModelBase) model;
		}
		String strValue = Utils.trimToEmpty(getCellValueAsString(wb, cell));
		modelBase.setOriginalValue(columnInfo.field.getName(), strValue);
		Object value = null;
		if (Date.class.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				String dateFormat = columnInfo.annotation.dateFormat();
				if (Utils.isBlank(dateFormat)) {
					dateFormat = "yyyy-MM-dd";
				}
				try {
					value = Utils.parseDate(strValue, dateFormat);
				} catch (ParseException e) {
					modelBase.addFieldError(columnInfo.field.getName(),
							columnInfo.annotation.name() + "日期格式无法解析，应该为" + dateFormat + "格式");
				}
			}
		} else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				try {
					value = Integer.valueOf(strValue, 10);
				} catch (Exception e) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "数字格式无法解析");
				}
			}
		} else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				try {
					value = Long.valueOf(strValue, 10);
				} catch (Exception e) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "数字格式无法解析");
				}
			}
		} else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				try {
					value = Float.valueOf(strValue);
				} catch (Exception e) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "数字格式无法解析");
				}
			}
		} else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				try {
					value = Double.valueOf(strValue);
				} catch (Exception e) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "数字格式无法解析");
				}
			}
		} else if (BigDecimal.class.equals(fieldType)) {
			if (Utils.isNotBlank(strValue)) {
				try {
					value = new BigDecimal(strValue);
				} catch (Exception e) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "数字格式无法解析");
				}
			}
		} else {
			value = strValue;
		}
		if (Utils.isNotBlank(strValue)) {
			List<ExcelColumnSelectionItem> selectionItems = columnInfo.getSelectionItems();
			if (selectionItems != null) {
				String realValue = null;
				for (ExcelColumnSelectionItem selectionItem : selectionItems) {
					if (selectionItem.getName().equals(value)) {
						realValue = selectionItem.getValue();
					}
				}
				if (realValue == null) {
					modelBase.addFieldError(columnInfo.field.getName(), columnInfo.annotation.name() + "值不在可选范围内");
					value = null;
				} else {
					value = realValue;
				}
			}
		}
		if (value != null) {
			try {
				columnInfo.field.set(model, value);
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}
		return Utils.isNotBlank(strValue);
	}

	/** 获取单元格字符串值 */
	private static String getCellValueAsString(Workbook wb, Cell cell) {
		if (cell == null) {
			return null;
		}
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
		DataFormatter dataFormatter = new DataFormatter();
		return dataFormatter.formatCellValue(cell, evaluator);
	}

	private static class ExcelColumnInfo {
		ExcelColumn annotation;
		Field field;

		boolean selectionItemsInitialized;
		List<ExcelColumnSelectionItem> selectionItems;

		public List<ExcelColumnSelectionItem> getSelectionItems() {
			if (!selectionItemsInitialized) {
				try {
					ExcelColumnSelectionProvider selectionProvider = getExcelColumnSelectionProvider(
							annotation.selectionProvider());
					selectionItems = selectionProvider.selectionItems(annotation.selectionType());
					selectionItemsInitialized = true;
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
			return selectionItems;
		}
	}

	private ExcelImportExportTools() {
	}

}
