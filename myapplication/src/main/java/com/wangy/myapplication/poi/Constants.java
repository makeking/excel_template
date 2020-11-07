package com.wangy.myapplication.poi;
/**
* @author wzy
* @date 2017年7月6日 下午2:22:30
*/
public class Constants {
	public static final String NEW_POI_FOREACH_START_REGEXP = "<poi:foreach\\s+type=\"(\\w+)\"\\s+list=\"(\\w+)\"\\s+var=\"(\\w+)\">";
	// 处理表格
	public static final String NEW_POI_KEY_REGEXP = "\\$\\{(\\w+)\\.(\\w+)\\}";
	// 处理图片
	public static final String POI_IMG_KEY_REGEXP = "\\$\\((\\w+).(\\w+)\\)";
//	public static final String NEW_POI_CHART = "<poi:foreach\\s+type=\"(\\w+)\"\\s+list=\"(\\w+)\"\\s+row=\"(\\w+)\"\\s+cell=\"(\\w+)\">";
//	public static final String NEWs_POI_FOREACH_START_REGEXP = "<poi:foreach\\s+type=\"(\\w+)\">";
//	public static final String POI_FOREACH_START_REGEXP = "<poi:foreach\\s+list=\"(\\w+)\">";
//	public static final String POI_FOREACH_END_REGEXP = "</poi:foreach>";
//	public static final String POI_PROPERTY_START_REGEXP = "<poi:property\\s+item=\"(\\w+)\">";
//	public static final String POI_PROPERTY_END_REGEXP = "</poi:property>";
//	// 处理类似${key}的串
//	public static final String POI_KEY_REGEXP = "\\$\\{(\\w+)\\}";
////	public static final String POI_KEY_REGEXP = "\\$\\{([^.]+?)\\}";
//	// 处理类似${vo.key}的串
//	public static final String POI_VO_DOT_KEY_REGEXP = "\\$(\\{)(\\w+\\.)(.+?\\})";
}
