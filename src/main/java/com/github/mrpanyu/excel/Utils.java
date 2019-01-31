package com.github.mrpanyu.excel;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

class Utils {

	public static boolean isBlank(String str) {
		return str == null || str.trim().length() == 0;
	}

	public static boolean isNotBlank(String str) {
		return !isBlank(str);
	}

	public static boolean equals(Object obj1, Object obj2) {
		if (obj1 == null && obj2 == null) {
			return true;
		} else if (obj1 != null) {
			return obj1.equals(obj2);
		} else {
			return false;
		}
	}

	public static String join(Iterable<String> strs, String separator) {
		StringBuilder sb = new StringBuilder();
		for (String str : strs) {
			if (sb.length() == 0) {
				sb.append(str);
			} else {
				sb.append(separator).append(str);
			}
		}
		return sb.toString();
	}

	public static String trimToEmpty(String str) {
		return str == null ? "" : str.trim();
	}

	public static Date parseDate(String dateStr, String format) throws ParseException {
		return new SimpleDateFormat(format).parse(dateStr);
	}

}
