package com.poi.util;


import com.alibaba.fastjson.JSON;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

public class ExcelUtil {
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 将对象数组转换成excel
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @param alias    指定对象属性别名，生成列名和列顺序Map<"类属性名","列名">
     * @param headLine 表标题
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, LinkedHashMap<String, String> alias, String headLine, int fieldNum) throws Exception {
        //创建一个工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        //创建一个表
        XSSFSheet sheet = wb.createSheet();
        //创建第一行，作为表名
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue(headLine);
        //设置表头格式
        CellStyle cellStyle = setExcelHeadType(wb);
        cell.setCellStyle(cellStyle);
        //设置表头单元格合并
        CellRangeAddress cra = new CellRangeAddress(0, 0, 0, fieldNum);
        sheet.addMergedRegion(cra);
        // 使用RegionUtil类为合并后的单元格添加边框
        // 下边框
        RegionUtil.setBorderBottom(1, cra, sheet);
        // 左边框
        RegionUtil.setBorderLeft(1, cra, sheet);
        // 有边框
        RegionUtil.setBorderRight(1, cra, sheet);
        // 上边框
        RegionUtil.setBorderTop(1, cra, sheet);
        //设置表列名格式
        CellStyle excelCellHeadType = setExcelCellHeadType(wb);
        //在第一行插入列名
        insertColumnName(1, sheet, alias, excelCellHeadType);
        //设置数据单元格格式
        CellStyle excelCellType = setExcelCellType(wb);
        //从第2行开始插入数据
        insertColumnDate(2, pojoList, sheet, alias, excelCellType);
        //设置批注,先获取绘图对象
        XSSFDrawing p = sheet.createDrawingPatriarch();
        //获取批注对象(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2) 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
        XSSFComment comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 9, 7));
        //设置要添加批注的内容
        XSSFRichTextString rtf = new XSSFRichTextString("现身吧！！！小老弟");
        comment.setString(rtf);
        comment.setAuthor("小田");
        sheet.getRow(1).getCell(0).setCellComment(comment);
        //设置列宽自动调整,必须在所有值插入完毕执行
        Set<Entry<String, String>> entries = alias.entrySet();
        Integer column = 0;
        for (Entry<String, String> entry : entries) {
            //设置和pojo类个数相同的列的列宽自动调整
            sheet.autoSizeColumn(column++);
        }
        // 处理中文不能自动调整列宽的问题
        setSizeColumn(sheet, fieldNum);
        //输出表格文件
        try {
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            wb.close();
        }
    }

    /**
     * 自适应列宽，中文
     *
     * @param sheet
     * @param size
     */
    private static void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
    }

    /**
     * 设置列名单元格格式
     *
     * @param wb
     * @return
     */
    public static CellStyle setExcelCellType(XSSFWorkbook wb) {
        XSSFCellStyle cellStyle = wb.createCellStyle();
        //设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        XSSFFont font = wb.createFont();
        //设置字体名称
        font.setFontName("宋体");
        //设置字号
        font.setFontHeightInPoints((short) 11);
        //设置是否为斜体
        font.setItalic(false);
        //设置是否加粗
        font.setBold(false);
        //设置字体颜色
        font.setColor(IndexedColors.BLACK.index);
        //下边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        //左边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        // 上边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        //右边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 设置列名单元格格式
     *
     * @param wb
     * @return
     */
    public static CellStyle setExcelCellHeadType(XSSFWorkbook wb) {
        XSSFCellStyle cellStyle = wb.createCellStyle();
        //设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        XSSFFont font = wb.createFont();
        //设置字体名称
        font.setFontName("Times New Roman");
        //设置字号
        font.setFontHeightInPoints((short) 15);
        //设置是否为斜体
        font.setItalic(false);
        //设置是否加粗
        font.setBold(true);
        //设置字体颜色
        font.setColor(IndexedColors.DARK_RED.index);
        //下边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        //左边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        // 上边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        //右边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 设置表头单元格格式
     *
     * @return
     */
    public static CellStyle setExcelHeadType(XSSFWorkbook wb) {
        XSSFCellStyle cellStyle = wb.createCellStyle();
        //设置水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        XSSFFont font = wb.createFont();
        //设置字体名称
        font.setFontName("宋体");
        //设置字号
        font.setFontHeightInPoints((short) 20);
        //设置是否为斜体
        font.setItalic(false);
        //设置是否加粗
        font.setBold(true);
        //设置字体颜色
        font.setColor(IndexedColors.BLACK.index);
        //下边框
        cellStyle.setBorderBottom(BorderStyle.THIN);
        //左边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        // 上边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        //右边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 将对象数组转换成excel
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @param alias    指定对象属性别名，生成列名和列顺序
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, LinkedHashMap<String, String> alias, int fieldNum) throws Exception {
        //获取类名作为标题
        String headLine = "";
        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            Class<? extends Object> clazz = pojo.getClass();
            headLine = clazz.getName();
            pojo2Excel(pojoList, out, alias, headLine, fieldNum);
        }
    }

    /**
     * 将对象数组转换成excel,列名为对象属性名
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @param headLine 表标题
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, String headLine, int fieldNum) throws Exception {

        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            //反射得到所有属性
            Class<?> aClass = pojo.getClass();
            //获取类的属性作为列名
            LinkedHashMap<String, String> alias = getAlias(aClass);
            pojo2Excel(pojoList, out, alias, headLine, fieldNum);
        }
    }

    /**
     * 获得linkedhashmap
     *
     * @param clazz
     * @return
     */
    public static LinkedHashMap<String, String> getAlias(Class<?> clazz) {
        //获取类的属性作为列名
        LinkedHashMap<String, String> alias = new LinkedHashMap<String, String>();
        Field[] fields = clazz.getDeclaredFields();
        //创建长度等于属性长度的 数组
        String[] name = new String[fields.length];
        //暴力反射所有字段
        Field.setAccessible(fields, true);
        //循环遍历将所有字段放在map中
        for (int i = 0; i < name.length; i++) {
            name[i] = fields[i].getName();
            if (!name[i].equals("serialVersionUID")) {
                alias.put(isNull(name[i]).toString(), isNull(name[i]).toString());
            }
        }
        return alias;
    }

    /**
     * 将对象数组转换成excel，列名默认为对象属性名，标题为类名
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, int fieldNum) throws Exception {

        //获取类名作为标题
        String headLine = "";
        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            Class<? extends Object> clazz = pojo.getClass();
            headLine = clazz.getName();
            //获取类的属性作为列名
            LinkedHashMap<String, String> alias = getAlias(clazz);
            pojo2Excel(pojoList, out, alias, headLine, fieldNum);
        }
    }

    /**
     * 此方法作用是创建表头的列名
     *
     * @param alias  要创建的表的列名与实体类的属性名的映射集合
     * @param rowNum 指定行创建列名
     * @return
     */
    private static void insertColumnName(int rowNum, XSSFSheet sheet, Map<String, String> alias, CellStyle excelCellHeadType) {
        XSSFRow row = sheet.createRow(rowNum);
        //列的数量
        int columnCount = 0;
        //遍历映射集合
        Set<Entry<String, String>> entrySet = alias.entrySet();
        for (Entry<String, String> entry : entrySet) {
            //创建第一行的第columnCount个格子
            XSSFCell cell = row.createCell(columnCount++);
            //将此格子的值设置为alias中的键名
            cell.setCellValue(isNull(entry.getValue()).toString());
            cell.setCellStyle(excelCellHeadType);
        }
    }

    /**
     * 从指定行开始给表格中插入数据
     *
     * @param beginRowNum 开始行
     * @param models      对象数组
     * @param sheet       表
     * @param alias       列别名
     * @throws Exception
     */
    private static <T> void insertColumnDate(int beginRowNum, List<T> models, XSSFSheet sheet, Map<String, String> alias, CellStyle excelCellType) throws Exception {
        for (T model : models) {
            //创建新的一行，先创建在++
            XSSFRow rowTemp = sheet.createRow(beginRowNum++);
            logger.info("创建了第:{}行", beginRowNum);
            //获取列的迭代
            Set<Entry<String, String>> entrySet = alias.entrySet();
            //从第0个格子开始创建
            int columnNum = 0;
            for (Entry<String, String> entry : entrySet) {
                //获取当前属性的类型，如果是Date类型就进行格式化
                String property = BeanUtils.getProperty(model, entry.getKey());
                Field declaredField = model.getClass().getDeclaredField(entry.getKey());
                String typeName = declaredField.getGenericType().getTypeName();
                if ("java.util.Date".equals(typeName) && property != null) {
                    SimpleDateFormat format = new SimpleDateFormat("E MMM dd hh:mm:ss z yyyy", Locale.US);
                    Date parse = format.parse(property);
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                    property = simpleDateFormat.format(parse);
                }
                //创建一个格子
                XSSFCell cell = rowTemp.createCell(columnNum++);
                cell.setCellValue(property);
                cell.setCellStyle(excelCellType);
            }
        }
    }

    /**
     * 判断是否为空，若为空设为""
     */
    private static Object isNull(Object object) {
        if (object != null) {
            return object;
        } else {
            return "";
        }
    }

    /**
     * 将excel表转换成指定类型的对象数组
     *
     * @param clazz 类型
     * @param alias 列别名,格式要求：Map<"列名","类属性名">
     * @return
     * @throws IOException
     * @throws IllegalArgumentException
     * @throws SecurityException
     */
    public static <T> List<T> excel2Pojo(InputStream inputStream, Class<T> clazz, LinkedHashMap<String, String> alias) throws IOException {
        XSSFWorkbook xh = new XSSFWorkbook(inputStream);
        try {
            //获取到第一张sheet表
            XSSFSheet sheet = xh.getSheetAt(0);
            //生成属性和列对应关系的map，Map<类属性名，对应一行的第几列>
            Map<String, Integer> propertyMap = generateColumnPropertyMap(sheet, alias);
            //根据指定的映射关系进行转换
            List<T> pojoList = generateList(sheet, propertyMap, clazz);
            return pojoList;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        } finally {
            xh.close();
        }
    }

    /**
     * 生成一个属性-列的对应关系的map
     *
     * @param sheet 表
     * @param alias 别名
     * @return
     */
    private static Map<String, Integer> generateColumnPropertyMap(XSSFSheet sheet, LinkedHashMap<String, String> alias) {
        Map<String, Integer> propertyMap = new HashMap<>();
        //获取到表头后的第一行
        XSSFRow propertyRow = sheet.getRow(1);
        //获取到当前行的第一个单元格的列数
        short firstCellNum = propertyRow.getFirstCellNum();
        //获取到当前行的最后一个单元格的列数
        short lastCellNum = propertyRow.getLastCellNum();
        //对当前行的所有列进行遍历
        for (int i = firstCellNum; i < lastCellNum; i++) {
            //获取到当前的单元格
            Cell cell = propertyRow.getCell(i);
            //空就跳过，进行下次循环
            if (cell == null) {
                continue;
            }
            //列名（就是单元格的值）
            String cellValue = cell.getStringCellValue();
            //从自己设定的map中找出对应属性名
            String propertyName = alias.get(cellValue);
            //获取到当前属性在哪一列的map<属性名，第几列>
            propertyMap.put(propertyName, i);
        }
        return propertyMap;
    }

    /**
     * 根据指定关系将表数据转换成对象数组
     *
     * @param sheet       表
     * @param propertyMap 属性映射关系Map<"属性名",一行第几列>
     * @param clazz       类类型
     * @return
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private static <T> List<T> generateList(XSSFSheet sheet, Map<String, Integer> propertyMap, Class<T> clazz) throws InstantiationException, IllegalAccessException, InvocationTargetException {
        //对象数组
        List<T> pojoList = new ArrayList<>();
        for (Row row : sheet) {
            //跳过前两行标题和列名
            if (row.getRowNum() < 2) {
                continue;
            }
            //反射创建一个T对象
            T instance = clazz.newInstance();
            //对映射好关系的map进行遍历
            Set<Entry<String, Integer>> entrySet = propertyMap.entrySet();
            for (Entry<String, Integer> entry : entrySet) {
                //获取此行指定列的值,即为属性对应的值（map中的value存的就是这个属性在那个列），获取的值可能是空
                String property = null;
                try {
                    //处理数字格式问题，输入0，POI会当成double来处理成0.0，
                    Cell cell = row.getCell(entry.getValue());
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    property = cell.getStringCellValue();
                    if ("".equals(property)) {
                        property = null;
                    }
                    Field field = clazz.getDeclaredField(entry.getKey());
                    String typeName = field.getGenericType().getTypeName();
                    //1、日期类型的处理
                    if (property != null && "java.util.Date".equals(typeName)) {
                        //时间格式化
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                        Date parse = simpleDateFormat.parse(property);
                        //保存属性
                        BeanUtils.setProperty(instance, entry.getKey(), parse);
                        //跳出循环
                        continue;
                    }
                    if (property != null) {
                        //java工具类，把参数放在指定类的指定属性中（将map赋值给一个类使用popul方法）
                        BeanUtils.setProperty(instance, entry.getKey(), property);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            pojoList.add(instance);
        }
        return pojoList;
    }

    /**
     * 将excel表转换成指定类型的对象数组，列名即作为对象属性
     *
     * @param clazz 类型
     * @return
     * @throws IOException
     * @throws SecurityException
     * @throws IllegalArgumentException
     */
    public static <T> List<T> excel2Pojo(InputStream inputStream, Class<T> clazz) throws IllegalArgumentException, SecurityException, IOException {
        LinkedHashMap<String, String> alias = new LinkedHashMap<String, String>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            //序列化添加的变量，不需要，排除。
            if (!field.getName().equals("serialVersionUID")) {
                alias.put(field.getName(), field.getName());
            }
        }
        List<T> pojoList = excel2Pojo(inputStream, clazz, alias);
        return pojoList;
    }
}
