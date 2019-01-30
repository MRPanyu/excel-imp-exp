package com.github.mrpanyu.excelimpexp;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Retention(RUNTIME)
@Target({ FIELD, METHOD })
public @interface ExcelColumn {

	/**
	 * 列名称，显示在列头上
	 */
	String name() default "";

	/**
	 * 备注信息，显示在列头上
	 */
	String notes() default "";

	/**
	 * 列宽（大致按英文字数计算），等于0时为自动宽度
	 */
	int width() default 0;

	/**
	 * 水平对齐方式
	 */
	ExcelColumnHorizontalAlignment horizontalAlignment() default ExcelColumnHorizontalAlignment.CENTER;

	/**
	 * 提供可选项值的类，如果有值，这个列会有下拉列表
	 */
	Class<? extends ExcelColumnSelectionProvider> selectionProvider() default NullExcelColumnSelectionProvider.class;

	/**
	 * 可选项类型，会作为参数传给{{@link #selectionProvider()}指定的类
	 */
	String selectionType() default "";

	/**
	 * 当可选项是级联选择的时候，指定上级的字段名称。
	 */
	String selectionRefField() default "";

	/**
	 * 日期格式
	 */
	String dateFormat() default "";

}
