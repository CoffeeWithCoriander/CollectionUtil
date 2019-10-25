package test;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class DateUtil {

	 //定义常用变量-
	public static final String DATE_FORMAT_FULL = "yyyy-MM-dd HH:mm:ss";
	public static final String DATE_FORMAT_YMD = "yyyy-MM-dd";
	public static final String DATE_FORMAT_HMS = "HH:mm:ss";
	public static final String DATE_FORMAT_HM = "HH:mm";
	public static final String DATE_FORMAT_YMDHM = "yyyy-MM-dd HH:mm";
	public static final String DATE_FORMAT_YMDHMS = "yyyyMMddHHmmss";
	public static final long ONE_DAY_MILLS = 3600000 * 24;
	public static final int WEEK_DAYS = 7;
	private static final int DATE_LENGHT = DATE_FORMAT_YMDHM.length();

	//1.日期转换为制定格式字符串
	public static String formatDateToString(Date time, String format) {
	    if (time == null) {
			return null;
		}
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		return sdf.format(time);
	}

	//2.字符串转换成制定格式日期
	public static Date formatStringToDate(String date, String format) {
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		try {
			return sdf.parse(date);
		} catch (ParseException ex) {
			return null;
		}
	}

	//3.判断一个日期是否属于两个时段内
	public static boolean isTimeInRange(Date time, Date[] timeRange) {
		return (!time.before(timeRange[0]) && !time.after(timeRange[1]));
	}

	//4.从完整的时间截取精确到分的时间
	public static String getDateToMinute(String fullDateStr) {
		return fullDateStr == null ? null
		: (fullDateStr.length() >= DATE_LENGHT ? fullDateStr.substring(0, DATE_LENGHT) : fullDateStr);
	}

	public static int getWeeksOfWeekYear(final int year) {
		Calendar cal = Calendar.getInstance();
		cal.setFirstDayOfWeek(Calendar.MONDAY);
		cal.setMinimalDaysInFirstWeek(WEEK_DAYS);
		cal.set(Calendar.YEAR, year);
		return cal.getWeeksInWeekYear();
	}

	//5.获取指定年份的第几周的第几天对应的日期yyyy-MM-dd，指定周几算一周的第一天（firstDayOfWeek）
	public static String getDateForDayOfWeek(int year, int weekOfYear, int dayOfWeek, int firstDayOfWeek) {
		Calendar cal = Calendar.getInstance();
		cal.setFirstDayOfWeek(firstDayOfWeek);
		cal.set(Calendar.DAY_OF_WEEK, dayOfWeek);
		cal.setMinimalDaysInFirstWeek(WEEK_DAYS);
		cal.set(Calendar.YEAR, year);
		cal.set(Calendar.WEEK_OF_YEAR, weekOfYear);
		return formatDateToString(cal.getTime(), DATE_FORMAT_YMD);
	}

	//7.获取指定日期的星期几
	public static int getWeekOfDate(String datetime) {
		Calendar cal = Calendar.getInstance();
		cal.setFirstDayOfWeek(Calendar.MONDAY);
		cal.setMinimalDaysInFirstWeek(WEEK_DAYS);
		Date date = formatStringToDate(datetime, DATE_FORMAT_YMD);
		cal.setTime(date);
		return cal.get(Calendar.DAY_OF_WEEK);
	}

	//8.计算某年周内的所有日期(从周一开始 为每周的第一天)
	@SuppressWarnings("rawtypes")
	public static List getWeekDays(int yearNum, int weekNum) {
		return getWeekDays(yearNum, weekNum, Calendar.MONDAY);
	}

	//9.计算某年某周内的所有日期(七天)
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static List getWeekDays(int year, int weekOfYear, int firstDayOfWeek) {
		List dates = new ArrayList<>();
		int dayOfWeek = firstDayOfWeek;
		for (int i = 0; i < WEEK_DAYS; i++) {
			dates.add(getDateForDayOfWeek(year, weekOfYear, dayOfWeek++, firstDayOfWeek));
		}
		return dates;
	}


	//10.计算两个日期的天数（startDate 开始日期字串、endDate 结束日期字串）
	public static int getDaysBetween(String startDate, String endDate) {
		int dayGap = 0;
		if (startDate != null && startDate.length() > 0 && endDate != null && endDate.length() > 0) {
			Date end = formatStringToDate(endDate, DATE_FORMAT_YMD);
			Date start = formatStringToDate(startDate, DATE_FORMAT_YMD);
			dayGap = getDaysBetween(start, end);
		}
		return dayGap;
	}

	private static int getDaysBetween(Date startDate, Date endDate) {
		return (int) ((endDate.getTime() - startDate.getTime()) / ONE_DAY_MILLS);
	}

	//11.计算两个日期之间的天数差
	public static int getDaysGapOfDates(Date startDate, Date endDate) {
		int date = 0;
		if (startDate != null && endDate != null) {
			date = getDaysBetween(startDate, endDate);
		}
		return date;
	}

	//12.计算两个日期之间的年份差距
	public static int getYearGapOfDates(Date firstDate, Date secondDate) {
		if (firstDate == null || secondDate == null) {
			return 0;
		}
		Calendar helpCalendar = Calendar.getInstance();
		helpCalendar.setTime(firstDate);
		int firstYear = helpCalendar.get(Calendar.YEAR);
		helpCalendar.setTime(secondDate);
		int secondYear = helpCalendar.get(Calendar.YEAR);
		return secondYear - firstYear;
	}

	//13.计算两个日期之间的月份差距
	public static int getMonthGapOfDates(Date firstDate, Date secondDate) {
		if (firstDate == null || secondDate == null) {
			return 0;
		}
		return (int) ((secondDate.getTime() - firstDate.getTime()) / ONE_DAY_MILLS / 30);
	}

	//14.计算是否包含当前日期
	@SuppressWarnings("rawtypes")
	public static boolean isContainCurrent(List dates) {
		boolean flag = false;
		SimpleDateFormat fmt = new SimpleDateFormat(DATE_FORMAT_YMD);
		Date date = new Date();
		String dateStr = fmt.format(date);
		for (int i = 0; i < dates.size(); i++) {
			if (dateStr.equals(dates.get(i))) {
				flag = true;
			}
		}
		return flag;
	}

	//15.从date开始计算time天后的日期
	public static String getCalculateDateToString(String startDate, int time) {
		String resultDate = null;
		if (startDate != null && startDate.length() > 0) {
			Date date = formatStringToDate(startDate, DATE_FORMAT_YMD);
			Calendar c = Calendar.getInstance();
			c.setTime(date);
			c.set(Calendar.DAY_OF_YEAR, time);
			date = c.getTime();
			resultDate = formatDateToString(date, DATE_FORMAT_YMD);
		}
		return resultDate;
	}


	//16.根据时间点获取时间区间
	public static List<String[]> getTimePointsByHour(int[] hours) {
		List<String[]> hourPoints = new ArrayList<>();
		String sbStart = ":00:00";
		String sbEnd = ":59:59";
		for (int i = 0; i < hours.length; i++) {
			String[] times = new String[2];
			times[0] = hours[i] + sbStart;
			times[1] = (hours[(i + 1 + hours.length) % hours.length] - 1) + sbEnd;
			hourPoints.add(times);
		}
		return hourPoints;
	}

	//17.根据指定的日期，增加或者减少天数
	public static Date addDays(Date date, int amount) {
		return add(date, Calendar.DAY_OF_MONTH, amount);
	}
	public static Date add(Date date, int calendarField, int amount) {
		if (date == null) {
			throw new IllegalArgumentException("The date must not be null");
		}
		Calendar c = Calendar.getInstance();
		c.setTime(date);
		c.add(calendarField, amount);
		return c.getTime();
	}

	//18.获取当前日期的最大日期 时间2014-12-21 23:59:59
	public static Date getCurDateWithMaxTime() {
		Calendar c = Calendar.getInstance();
		c.set(Calendar.HOUR_OF_DAY, 23);
		c.set(Calendar.MINUTE, 59);
		c.set(Calendar.SECOND, 59);
		return c.getTime();
	}

	//19.获取当前日期的最小日期时间 2014-12-21 00:00:00
	public static Date getCurDateWithMinTime() {
		Calendar c = Calendar.getInstance();
		c.set(Calendar.HOUR_OF_DAY, 0);
		c.set(Calendar.MINUTE, 0);
		c.set(Calendar.SECOND, 0);
		c.set(Calendar.MILLISECOND, 0);
		return c.getTime();
	}

}
