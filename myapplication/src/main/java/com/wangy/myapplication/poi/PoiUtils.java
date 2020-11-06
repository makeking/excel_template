package com.wangy.myapplication.poi;

import android.annotation.SuppressLint;
import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LineChartSeries;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.charts.XSSFCategoryAxis;
import org.apache.poi.xssf.usermodel.charts.XSSFChartAxis;
import org.apache.poi.xssf.usermodel.charts.XSSFChartLegend;
import org.apache.poi.xssf.usermodel.charts.XSSFValueAxis;
import org.apache.xmlbeans.XmlString;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.lang.Integer.parseInt;

/**
 * todo 缺少对合并单元格的操作
 */
public class PoiUtils {
    public PoiUtils() {
    }

    private static XSSFWorkbook hssfWorkbook;

    public static PoiUtils init(String oldFile) throws Exception {
        if (oldFile.isEmpty()) {
            throw new NullPointerException("请将模板填充到工具类中");
        } else if (oldFile.endsWith(".xls") | oldFile.endsWith(".xlsx")) {

            File oldfile = new File(oldFile);
            if (oldfile.exists()) {
                // 解析模板

                hssfWorkbook = new XSSFWorkbook(new FileInputStream(oldfile));

                return new PoiUtils();
            } else {
                throw new FileNotFoundException("未找到文件");
            }
        } else {
            throw new NumberFormatException("请填充正确的数据类型");
        }
    }

    int lastRowNum1 = 0;
    int insertRowNum = 0;
    int endNum = 0;
    int nums = 0;
    int mergeRow = 1;
    int mergeColumn = 1;
    XSSFCellStyle rowStyle = null;

    public PoiUtils writeTab(Map<String, Object[]> data, Class clazz, String check, Map<String, ArrayList<String>> imageMap, String
            imagss) {
        for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
            //获取对应的sheet
            XSSFSheet sheet = hssfWorkbook.getSheetAt(i);
            // 判断 srccell 是否合并单元格
            List<CellRangeAddress> rangeList = new ArrayList<>();
            //获得一个 sheet 中合并单元格的数量
            int sheetmergerCount = sheet.getNumMergedRegions();
            //遍历所有的合并单元格
            for (int j = 0; j < sheetmergerCount; j++) {
                //获得合并单元格保存进list中
                CellRangeAddress ca = sheet.getMergedRegion(j);
                rangeList.add(ca);
            }
            Log.e("tag", "合并单元格有  " + rangeList.size() + "个");
            int lastRowNum = sheet.getLastRowNum();
            Log.e("tag", "lastRowNum " + lastRowNum);
            // 1. 获取所有的折线图
            XSSFDrawing drawingPatriarch = sheet.getDrawingPatriarch();
            List<XSSFChart> charts = drawingPatriarch.getCharts();
            // todo 先模拟第一个折线图
            XSSFChart xssfChart = charts.get(0);
            CTChart ctChart = xssfChart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();
            CTLineSer[] serArray = plotArea.getLineChartArray(0).getSerArray();
            // todo Android use this method can throw excetion
//            List<CTLineSer> serList = lineChartArray.getSerList();

            String f = serArray[0].getVal().getNumRef().getF();
            Log.e("tag", "第一列列数区间为 " + f);
            // 获取行数区间
            String val = serArray[0].getCat().getNumRef().getF();
            Log.e("tag", "formatCode " + val);


//            //sheet1表示操作的sheet
//            XSSFDrawing drawingPatriarch = sheet1.getDrawingPatriarch();
//            List<XSSFChart> charts = drawingPatriarch.getCharts();
//            XSSFChart xssfChart = charts.get(0);
//            CTChart ctChart = xssfChart.getCTChart();
//            CTPlotArea plotArea = ctChart.getPlotArea();
//            //dataRangeMaxValue 代表取值范围单元格行数最大值
//            Integer dataRangeMaxValue = 9;
//            List<CTRadarSer> serList = plotArea.getRadarChartArray(0).getSerList();
//            //serList1.get(0).getVal().getNumRef().getF() -> 'TestChart1'!$E$4:$E$9
//            //serList1.get(1).getVal().getNumRef().getF() -> 'TestChart1'!$F$4:$F$9
//            //设置第一列取数区间
//            serList.get(0).getVal().getNumRef().setF("'TestChart1'!$E$4:$E$" + dataRangeMaxValue);
//            //设置第二列取数区间
//            serList.get(1).getVal().getNumRef().setF("'TestChart1'!$F$4:$F$" + dataRangeMaxValue);
//            //serList1.get(1).getCat().getNumRef().getF() -> 'TestChart1'!$D$7:$D$9
//            //注意:如果是一列标题对应多列数据,serList中都必须设置以下取数区域
//            //设置标题取数区间
//            serList.get(0).getCat().getNumRef().setF("'TestChart1'!$D$4:$D$" + dataRangeMaxValue);
//            serList.get(1).getCat().getNumRef().setF("'TestChart1'!$D$4:$D$" + dataRangeMaxValue);

//            // 图表图例
//            XSSFChartLegend legend = xssfChart.getOrCreateLegend();
//            // 设置标题的位置
//            legend.setPosition(LegendPosition.TOP);
//            // 获取图表数据区域
//
//            List<? extends XSSFChartAxis> axis = xssfChart.getAxis();
//            for (int j = 0; j < axis.size(); j++) {
//                XSSFChartAxis xssfChartAxis = axis.get(j);
//                // 最小值和最大值
//                double minimum = xssfChartAxis.getMinimum();
//                double maximum = xssfChartAxis.getMaximum();
//                // 设置值
//                AxisCrosses crosses = xssfChartAxis.getCrosses();
//                Class<AxisCrosses> declaringClass = crosses.getDeclaringClass();
//
//            }
            if (check == null && imagss == null) {
                lastRowNum1 = sheet.getLastRowNum();
            }
            if (check != null || imagss != null) {
                // 插入行
                //todo  判断插入的是否是合并单元格
//                try {
//                    XSSFRow row = sheet.getRow(insertRowNum);
//                    for (int j = 0; j < row.getLastCellNum(); j++) {
//                        XSSFCell cell = row.getCell(j);
//                        Map<String, Object> combineCell = PoiCopyUtils.isCombineCell(rangeList, cell);
//                        if ((Boolean) combineCell.get("flag")) {
//                            // 判断是否是合并单元格
//                            int mergedRow = (int) combineCell.get("mergedRow");
//                            int mergedCol = (int) combineCell.get("mergedCol");
//                            Log.e("tag", " mergedRow =" + mergedRow + ",mergedCol = " + mergedCol);
//                            if (mergedRow > mergeRow) {
//                                mergeRow = mergedRow;
//                            }
//
//                        }
//                    }
//                } catch (Exception e) {
//                    e.printStackTrace();
//                }
                if (check != null) {
                    Object[] students = data.get(check);
                    try {
                        assert students != null;
                        ArrayList arrayLists = (ArrayList) students[1];
                        if (arrayLists == null || arrayLists.isEmpty()) {
                            throw new NullPointerException("Don't find table data, please cheack you table data!");
                        } else {
                            //todo  mergeRow 合并行
                            lastRowNum1 += arrayLists.size();
                            nums = arrayLists.size() * mergeRow;
                            endNum = arrayLists.size() + insertRowNum;
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else if (imageMap != null) {
                    ArrayList<String> list = imageMap.get(imagss);
                    if (list != null && !list.isEmpty()) {
                        lastRowNum1 += list.size();
                        nums = list.size() * mergeRow;
                        endNum = list.size() + insertRowNum;
                    } else {
                        throw new NullPointerException("Image data is null ,please check you image data!");
                    }
                }
                int lastRowNum2 = sheet.getLastRowNum();

                // todo 合并行为 mergeRow
                sheet.shiftRows(insertRowNum + 1, lastRowNum2, nums - 1, true, true);
                XSSFRow row = sheet.getRow(insertRowNum);
                if (row != null) {
                    rowStyle = row.getRowStyle();
                }
                //  2 3
                //  4 5
                //  6 7
                //  8 9
                for (int j = insertRowNum + 1; j < endNum; j++) {
//                    if (mergeRow > 1) {
//                        int num = j * mergeRow - mergeRow;
//                        short height = sheet.getRow(insertRowNum).getHeight();
//                        XSSFRow row1 = sheet.createRow(num);
//                        row1.setHeight(height);
//                        if (rowStyle != null) {
//                            row1.setRowStyle(rowStyle);
//                        }
//                        NewPoiCopyUtils.copyRow(hssfWorkbook, sheet.getRow(insertRowNum * (j - insertRowNum)), row1, true);
////                        PoiCopyUtils.copyRow(hssfWorkbook, sheet.getRow(insertRowNum * (j - insertRowNum)), row1, true, rangeList, sheet);
//                    } else {
                    short height = sheet.getRow(insertRowNum).getHeight();
                    XSSFRow row1 = sheet.createRow(j);
                    row1.setHeight(height);
                    if (rowStyle != null) {
                        row1.setRowStyle(rowStyle);
                    }
                    PoiCopyUtils.copyRow(hssfWorkbook, sheet.getRow(insertRowNum), row1, true, rangeList, sheet);
//                    }
                }
                // todo 设置数据
                newCheckData(data, clazz, check, imageMap, imagss, sheet);

            }

            for (int j = 0; j <= lastRowNum1; j++) {
                // 获取第i行的数据
                XSSFRow row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                for (int k = 0; k < row.getLastCellNum(); k++) {
                    XSSFCell cell = row.getCell(k);
                    if (cell != null && cell.getCellComment() != null) {
                        String tag = cell.getCellComment().getString().toString();
                        Log.e("tag", "注解： 当前的行是" + j + " 当前的列是 " + k + "内容是" + tag);
                        // 1. 先匹配是不是chart 类型
//                        Matcher matchers = Pattern.compile(Constants.NEW_POI_FOREACH_START_REGEXP).matcher(tag);
//                        if (matchers.find()) {
//                            String newType = Objects.requireNonNull(matchers.group(1)).trim();
//                            if ("chart".equals(newType)) {
//                                // 设置数据
//                                // 判断list
//                                String dataLists = Objects.requireNonNull(matchers.group(2)).trim();
//                                // 判断row
//                                String rowName = Objects.requireNonNull(matchers.group(3)).trim();
//                                // 判断 cell
//                                String cellName = Objects.requireNonNull(matchers.group(4)).trim();
//                                for (String key : data.keySet()) {
//                                    if (dataLists.equals(key)) {
//                                        // 获取数据
//                                        Object[] dats = data.get(key);
//                                        assert dats != null;
//                                        int isList = (int) dats[0];
//                                        if (0 == isList) {
//                                        } else {
//                                            // 表示传递过来的集合
//                                            ArrayList list = (ArrayList) dats[1];
//                                            Class clazzs = (Class) dats[2];
//                                            Field[] fields = clazzs.getDeclaredFields();
//                                            for (Field field : fields) {
//                                                field.setAccessible(true); //设置属性为可访问
//                                                String name = field.getName();//获取属性的名称
//                                                if (rowName.equals(name)) {
//                                                    // 设置行
//
//                                                } else if (cellName.equals(name)) {
//                                                    // 设置列
//
//                                                }
//                                                field.setAccessible(false);
//
//                                            }
//                                        }
//                                    }
//                                }
//                            }
//                        }
                        Matcher ma = Pattern.compile(Constants.NEW_POI_FOREACH_START_REGEXP).matcher(tag);
                        // todo 注意不可以使用 groupCount(),要不然会进行分组
                        if (ma.find()) {
                            String type = Objects.requireNonNull(ma.group(1)).trim();
                            if ("tab".equals(type)) {
                                cell.removeCellComment();
                                String dataLists = Objects.requireNonNull(ma.group(2)).trim();
                                Log.e("tag", "tab group " + i + " 的内容是 " + dataLists);
                                // 记录当前的内容，记录当前的行数和列数
                                // key0 表示的是这个map集合的key
                                for (String key : data.keySet()) {
                                    if (dataLists.equals(key)) {
                                        // 表示获取到了数据
                                        // 获取data 的数量
                                        Object[] dats = data.get(key);
                                        assert dats != null;
                                        int isList = (int) dats[0];
                                        if (0 == isList) {
                                            // 是普通数据类型
                                            parseValue(dats[1], (Class) dats[2], cell);
                                        } else {
                                            // 表示传递过来的集合
                                            ArrayList list = (ArrayList) dats[1];
                                            if (!list.isEmpty()) {
                                                insertRowNum = j;
                                                // 跳出循环
                                                writeTab(data, (Class) dats[2], key, imageMap, null);
                                            }
                                        }
                                    }
                                }
                            }
                            if ("img".equals(type)) {
                                cell.removeCellComment();
                                //  表示是图片
                                String imgdDataLists = Objects.requireNonNull(ma.group(2)).trim();
                                Log.e("tag", "img group " + i + " 的内容是 " + imgdDataLists);
                                // 类型
                                assert imageMap != null;
                                for (String imags : imageMap.keySet()) {
                                    if (imgdDataLists.equals(imags)) {
                                        // 表示获取到了数据
                                        // 获取data 的数量
                                        ArrayList<String> list = imageMap.get(imags);
                                        assert list != null;
                                        if (!list.isEmpty()) {
                                            insertRowNum = j;
                                            writeTab(data, null, null, imageMap, imags);
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
        }
        return this;
    }

    private void newCheckData(Map<String, Object[]> data, Class clazz, String
            check, Map<String, ArrayList<String>> imageMap, String imagss, XSSFSheet sheet) {
        for (int j = 0; j <= lastRowNum1; j++) {
            // 获取第i行的数据
            XSSFRow row = sheet.getRow(j);
            if (row == null) {
                continue;
            }
            short lastCellNum = row.getLastCellNum();
            for (int k = 0; k < lastCellNum; k++) {
                XSSFCell cell = row.getCell(k);
                if (checkData(data, clazz, check, imageMap, imagss, sheet, j, (short) k, cell))
                    break;
            }
        }
    }

    private boolean checkData(Map<String, Object[]> data, Class clazz, String
            check, Map<String, ArrayList<String>> imageMap, String imagss, XSSFSheet sheet, int j,
                              short k, XSSFCell cell) {
        // 进行设置数据
        if (insertRowNum != 0 && endNum != 0) {
            if (j >= insertRowNum && j < endNum) {
                // 开始设置数据
                //获取实体类中所有属性的值
                if (check != null && clazz != null) {
                    setTableData(data, clazz, check, j, cell);
                }
                if (cell != null && imagss != null) {
                    if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                        // 进行正则匹配
                        String stringCellValue = cell.getStringCellValue().replace("\"", "").trim();
//                        Log.e("tag", "images 获取到的内容是 " + stringCellValue);
                        Matcher ma = Pattern.compile(Constants.POI_IMG_KEY_REGEXP).matcher(stringCellValue);
                        while ( ma.find() ) {
                            Log.e("tag", "images 匹配成功的内容是 " + stringCellValue);
                            int group1 = parseInt(Objects.requireNonNull(ma.group(2)));
                            cell.setCellValue("");
                            ArrayList<String> list = imageMap.get(imagss);
                            assert list != null;
                            for (int l = 0; l < list.size(); l++) {
                                if (group1 == l + 1) {
                                    FileInputStream is;
                                    try {
                                        // 设置图片
                                        //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
                                        XSSFDrawing patriarch = sheet.createDrawingPatriarch();
                                        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                                        is = new FileInputStream(new File(list.get(l)));
                                        byte[] b = new byte[1024 * 2];
                                        int len;
                                        while ( (len = is.read(b, 0, b.length)) != -1 ) {
                                            byteArrayOut.write(b, 0, len);
                                            byteArrayOut.flush();
                                        }
                                        XSSFClientAnchor xssfClientAnchor = new XSSFClientAnchor(0, 0, 0, 0, k, j, k + 1, j + 1);
                                        //anchor主要用于设置图片的属性
                                        /*
                                        dx1 dy1 起始单元格中的x,y坐标.

                                        dx2 dy2 结束单元格中的x,y坐标

                                        col1,row1 指定起始的单元格，下标从0开始

                                        col2,row2 指定结束的单元格 ，下标从0开始
                                        */
                                        //            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 21, 0, (short) 24, 2);
                                        // c1 表示的是右上角
                                        // c2 表示的是 5 表示距离左边
                                        xssfClientAnchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                                        //插入图片
                                        //todo 根据尾号判断类型  HSSFWorkbook.PICTURE_TYPE_PNG
                                        patriarch.createPicture(xssfClientAnchor, hssfWorkbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_PNG));
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                    break;
                                }
                            }
                            cell.removeCellComment();
                        }
                    }
                }
            } else if (j > endNum) {
                insertRowNum = 0;
                endNum = 0;
                return true;
            }
        }
        return false;
    }

    private void setTableData(Map<String, Object[]> data, Class clazz, String check,
                              int j, XSSFCell cell) {
        if (cell != null) {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true); //设置属性为可访问
                String name = field.getName();//获取属性的名称
                CellType cellTypeEnum = cell.getCellTypeEnum();
                String stringCellValue;
                // todo 这里使用 STRING 用来判断，原因有两种： 1. 使用通配符配置的类型一定是String 类型 2. cell 不可以设置两次
                if (cellTypeEnum.equals(CellType.STRING)) {
                    stringCellValue = cell.getStringCellValue().replace("\"", "").trim();
                    Matcher ma = Pattern.compile(Constants.NEW_POI_KEY_REGEXP).matcher(stringCellValue);
                    // todo 注意不可以使用 groupCount(),要不然会进行分组
                    if (ma.find()) {
                        String key0 = Objects.requireNonNull(ma.group(2)).trim();
                        // 对于值进行判断
                        if (name.equals(key0)) {
                            // 设置值
                            try {
                                Object[] students = data.get(check);
                                assert students != null;
                                ArrayList arrayLists = (ArrayList) students[1];
                                Class<?> type = field.getType();
                                Object object = field.get(arrayLists.get(j - insertRowNum));
                                assert object != null;
                                parseValue(object, type, cell);
                            } catch (IllegalAccessException e) {
                                e.printStackTrace();
                            }
                        }

                    }
                    if (cell.getCellComment() != null) {
                        cell.removeCellComment();
                    }

                    field.setAccessible(false);
                }

            }
        }

    }

    /**
     * 将字符串转化为指定类型的对象
     *
     * @param str  ----要转化的字符串
     * @param type ----目标对象类型
     * @param cell cell 的内容
     */
    private static void parseValue(Object str, Class type, XSSFCell cell) {
        Date obj = null;
        String className = type.getName();
        String s = str.toString();
        //excel中的数字解析之后可能末尾会有.0，需要去除
        if (s.endsWith(".0")) s = s.substring(0, s.length() - 2);
        switch (className) {
            case "java.lang.Integer":  //Integer
            case "int":  //int
                setNumberType(cell);
                cell.setCellValue(parseInt(s));
                break;
            case "java.lang.String":  //String
                cell.setCellType(CellType.STRING);
                cell.setCellValue(s);
                break;
            case "java.lang.Double":  //Double
            case "double":  //double
                setNumberType(cell);
                cell.setCellValue(Double.parseDouble(s));
                break;
            case "java.lang.Float":  //Float
            case "float":  //float
                setNumberType(cell);
                cell.setCellValue(Float.parseFloat(s));
                break;
            case "java.util.Date":  //Date
                @SuppressLint("SimpleDateFormat") SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                try {
                    obj = sdf.parse(s);
                } catch (ParseException e) {
                    e.printStackTrace();
                }
                setNumberType(cell);
                cell.setCellValue(obj);
                break;
            case "long":  //long
            case "java.util.Long":  //Long
                setNumberType(cell);
                cell.setCellValue(Long.parseLong(s));
                break;
        }
    }

    private static void setNumberType(XSSFCell cell) {
        cell.setCellType(CellType.NUMERIC);
    }

    public void copy(String newFile) {
        FileOutputStream fos;
        try {
            fos = new FileOutputStream(new File(newFile));
            hssfWorkbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
