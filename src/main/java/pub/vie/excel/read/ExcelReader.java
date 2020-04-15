package pub.vie.excel.read;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pub.vie.excel.common.annotation.ExcelEntity;
import pub.vie.excel.common.annotation.ExcelField;
import pub.vie.excel.common.utils.CommonUtils;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.net.URL;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.*;

import static pub.vie.excel.common.utils.CommonUtils.arrayEmpty;
import static pub.vie.excel.common.utils.CommonUtils.isBlank;

/**
 * @Descrption :
 * @Author: zoujie
 * @Date: 2020-4-13
 */
public class ExcelReader<T> {

    private static Logger log = LoggerFactory.getLogger(ExcelReader.class);

    private static final String DEFAULT_DATE_FORMATE = "yyyy-MM-dd";

    public static InputStream getStreamOnClassPath(String virtualPath) {
        return ExcelReader.class.getClassLoader().getResourceAsStream(virtualPath);
    }

    public static String getRealPathOnClassPath(String virtualPath) {
        URL resource = getClassLoader().getResource(virtualPath);
        if (resource == null) {
            return null;
        }
        String realPath;
        try {
            realPath = URLDecoder.decode(resource.getPath(), StandardCharsets.UTF_8.toString());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }
        return realPath;
    }

    public List<T> read(InputStream inputStream, Class<T> dataClass) {
        return read(inputStream, 0, 0, -1, dataClass);
    }

    public List<T> read(InputStream inputStream,int skip, Class<T> dataClass) {
        return read(inputStream, 0, skip, -1, dataClass);
    }

    public List<T> read(InputStream inputStream, int sheetAt, int skip, int limitRow, Class<T> dataClass) {
        return baseRead(dataClass,inputStream,sheetAt,skip,limitRow,false);
    }

    public List<T> read(Class<T> dataClass) {
        ExcelReadProperties excelReadProperties = excelReadProperties(dataClass);
        if (excelReadProperties == null) {
            throw new IllegalArgumentException(dataClass.getName() + "不含有ExcelEntity注解信息，无法解析excel");
        }

        String classPathSource = excelReadProperties.classPathSource;
        if (isBlank(classPathSource)) {
            throw new IllegalArgumentException(dataClass.getName() + "该方法ExcelEntity注解配置的classPathSource不能为空");
        }

        int sheetAt = excelReadProperties.sheetAt;
        int skip = excelReadProperties.skip;
        int limitRow = excelReadProperties.limitRow;

        InputStream inputStream = getStreamOnClassPath(classPathSource);

        return baseRead(dataClass, inputStream, sheetAt, skip, limitRow,true);
    }

    public List<T> read(Class<T> dataClass, InputStream inputStream) {
        ExcelReadProperties excelReadProperties = excelReadProperties(dataClass);
        if (excelReadProperties == null) {
            throw new IllegalArgumentException(dataClass.getName() + "不含有ExcelEntity注解信息，无法解析excel");
        }

        int sheetAt = excelReadProperties.sheetAt;
        int skip = excelReadProperties.skip;
        int limitRow = excelReadProperties.limitRow;
        return baseRead(dataClass, inputStream, sheetAt, skip, limitRow,true);
    }

    private List<T> baseRead(Class<T> dataClass, InputStream inputStream, int sheetAt, int skip, int limitRow, boolean useAnnotation) {
        Workbook workbook = null;
        List<T> resList = null;
        if (inputStream == null) {
            log.error("excel inputStream 为 null");
            return null;
        }
        BufferedInputStream bufferedInputStream = convertInputStream(inputStream);

        try {
            workbook = WorkbookFactory.create(bufferedInputStream);
            Sheet sheet = workbook.getSheetAt(sheetAt);
            if (sheet == null) {
                log.info("获取工作簿sheetAt[{}]为空");
                return null;
            }

            int lastRowNum = sheet.getLastRowNum();
            if (skip > lastRowNum) {
                log.warn("sheet 总条数[{}] 跳过[{}]条不合法直接返回", lastRowNum, skip);
                return null;
            }

            resList = new ArrayList<>();

            if (skip < 0) {
                skip = 0;
            }

            //最大条数
            int end;
            if (limitRow <= 0 || skip + limitRow > lastRowNum) {
                end = lastRowNum;
            } else {
                end = skip + limitRow - 1;
            }

            for (int i = skip; i <= end; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    break;
                }

                if (useAnnotation) {
                    resList.add(handleRow(dataClass, row));
                } else {
                    resList.add(handleRow(row, dataClass));
                }

            }

            log.info("总条数[{}] 成功处理了{}条excel数据,跳过了{}条excel数据", end, resList.size(), skip);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            CommonUtils.close(workbook, bufferedInputStream);
        }

        return resList;
    }

    private static ClassLoader getClassLoader() {
        return ExcelReader.class.getClassLoader();
    }

    private T handleRow(Row row, Class<T> resDataClass) {
        T resData;
        try {
            resData = resDataClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
            return null;
        }

        short lastCellNum = row.getLastCellNum();
        if (lastCellNum < 0) {
            throw new IllegalArgumentException("数据列小于1");
        }


        Field[] declaredFields = resDataClass.getDeclaredFields();
        if (arrayEmpty(declaredFields)) {
            throw new IllegalArgumentException(resDataClass.getName() + "未定义任何属性匹配excel文件");
        }

        int limit = declaredFields.length > lastCellNum ? lastCellNum : declaredFields.length;

        List<Cell> cellList = getCellList(row, limit);
        handleRow(Arrays.asList(declaredFields), cellList, resData, limit);
        return resData;
    }

    private void handleRow(List<Field> fieldList, List<Cell> cellList, Object object, int limit) {
        if (CollectionUtils.isEmpty(cellList) || CollectionUtils.isEmpty(fieldList)) {
            return;
        }

        if (limit > cellList.size() || limit > fieldList.size()) {
            throw new IndexOutOfBoundsException("cellList.size = " + cellList.size() + "fieldList.size=" + fieldList.size() + "max index=" + limit);
        }

        for (int i = 0; i < limit; i++) {
            Field field = fieldList.get(i);
            Cell cell = cellList.get(i);
            try {
                convertAndSet(field, object, cell);
            } catch (IllegalAccessException e) {
                log.error(e.getLocalizedMessage(), e);
            }
        }
        log.debug("解析到内容并封装为对象:[{}]", object);
    }

    private T handleRow(Class<T> dataClass, Row row) {
        Field[] declaredFields = dataClass.getDeclaredFields();

        Map<Integer, Field> colIndexFieldMap = selectFieldsToMap(declaredFields);
        if (colIndexFieldMap == null) {
            return null;
        }

        T dataObject;
        try {
            dataObject = dataClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }

        short lastCellNum = row.getLastCellNum();
        for (Map.Entry entry : colIndexFieldMap.entrySet()) {

            Integer key = (Integer) entry.getKey();
            Field value = (Field) entry.getValue();

            if (key < lastCellNum) {
                try {
                    convertAndSet(value, dataObject, row.getCell(key));
                } catch (IllegalAccessException e) {
                    log.error(e.getLocalizedMessage(), e);
                }
            }
        }

        return dataObject;
    }

    private BufferedInputStream convertInputStream(InputStream inputStream) {
        if (!(inputStream instanceof BufferedInputStream)) {
            return new BufferedInputStream(inputStream);
        } else {
            return (BufferedInputStream) inputStream;
        }
    }

    private List<Cell> getCellList(Row row, int limit) {
        List<Cell> cellList = new ArrayList<>();
        for (int j = 0; j < limit; j++) {
            cellList.add(row.getCell(j));
        }
        return cellList;
    }

    private void convertAndSet(Field field, Object object, Cell cell) throws IllegalAccessException {
        CellType cellType = cell.getCellType();
        field.setAccessible(true);
        Class<?> type = field.getType();
        switch (cellType) {
            case _NONE:
                log.debug("row[{}] cell[{}] 单元格内容未知类型", cell.getRowIndex(), cell.getColumnIndex());
                break;
            case NUMERIC:
                Double numericCellValue = cell.getNumericCellValue();
                if (Double.TYPE.equals(type) || Double.class.equals(type)) {
                    field.set(object, numericCellValue);
                } else if (Date.class.equals(type)) {
                    field.set(object, new Date(numericCellValue.longValue()));
                }else if(String.class.equals(type)){
                    field.set(object, String.valueOf(numericCellValue));
                }else if(Integer.TYPE.equals(type) || Integer.class.equals(type)){
                    field.set(object, numericCellValue.intValue());
                }else if(Long.TYPE.equals(type) || Long.class.equals(type)){
                    field.set(object, numericCellValue.longValue());
                }
                log.debug("row[{}] cell[{}] 单元格内容[{}]为数字", cell.getRowIndex(), cell.getColumnIndex(), cell.getNumericCellValue());
                break;
            case STRING:
                String stringCellValue = cell.getStringCellValue();
                if (String.class.equals(type)) {
                    field.set(object, stringCellValue);

                } else if (Date.class.equals(type)) {
                    SimpleDateFormat dateFormat = getDateFormat(field);
                    try {
                        Date parse = dateFormat.parse(stringCellValue);
                        field.set(object, parse);
                    } catch (Exception e) {
                        log.error("row[{}] col[{}] 内容为非日期类字符串,不能按照日期格式转换为Date类型或ExcelField注解配置的时间格式有误，detail message:\n{}", cell.getRowIndex(), cell.getColumnIndex(), e.getLocalizedMessage());
                    }

                } else if (Boolean.TYPE.equals(type) || Boolean.class.equals(type)) {

                    try {
                        boolean booleanCellValue = Boolean.valueOf(stringCellValue);
                        field.set(object, booleanCellValue);
                    } catch (Exception e) {
                        log.error("row[{}] col[{}] 内容为非boolean类型字符串,不能转换为boolean类型，detail message:\n{}", cell.getRowIndex(), cell.getColumnIndex(), e.getLocalizedMessage());
                    }
                }

                log.debug("row[{}] cell[{}] 单元格内容[{}]为字符串", cell.getRowIndex(), cell.getColumnIndex(), stringCellValue);
                break;
            case FORMULA:
                if (String.class.equals(type)) {
                    String formulaCellValue = cell.getCellFormula();
                    field.set(object, formulaCellValue);
                }
                log.debug("row[{}] cell[{}] 单元格内容[{}]为公式字符串", cell.getRowIndex(), cell.getColumnIndex(), cell.getCellFormula());
                break;
            case BLANK:
                log.debug("row[{}] cell[{}] 单元格内容为空", cell.getRowIndex(), cell.getColumnIndex());
                break;
            case BOOLEAN:
                if (Boolean.TYPE.equals(type) || Boolean.class.equals(type)) {
                    boolean booleanCellValue = cell.getBooleanCellValue();
                    field.set(object, booleanCellValue);
                }
                log.debug("row[{}] cell[{}] 单元格内容[{}]为boolean", cell.getRowIndex(), cell.getColumnIndex(), cell.getBooleanCellValue());
                break;
            case ERROR:
                log.debug("row[{}] cell[{}] 单元格内容错误", cell.getRowIndex(), cell.getColumnIndex());
                break;
            default:
                break;
        }
    }

    private SimpleDateFormat getDateFormat(Field field) {
        ExcelField[] annotationsByType = field.getAnnotationsByType(ExcelField.class);
        SimpleDateFormat format;
        if (arrayEmpty(annotationsByType)) {
            format = new SimpleDateFormat(DEFAULT_DATE_FORMATE);
        } else {
            format = new SimpleDateFormat(annotationsByType[0].dataFormat());
        }
        return format;
    }

    private ExcelReadProperties excelReadProperties(Class<T> tClass) {
        ExcelEntity[] annotationsByType = tClass.getAnnotationsByType(ExcelEntity.class);
        if (arrayEmpty(annotationsByType)) {
            return null;
        }

        ExcelEntity excelEntity = annotationsByType[0];
        ExcelReadProperties properties = new ExcelReadProperties();
        properties.classPathSource = excelEntity.classPathSource();
        properties.limitRow = excelEntity.limitRow();
        properties.sheetAt = excelEntity.sheetAt();
        properties.skip = excelEntity.skip();

        return properties;
    }

    private Map<Integer, Field> selectFieldsToMap(Field[] fields) {
        if (CommonUtils.arrayEmpty(fields)) {
            return null;
        }
        Map<Integer, Field> selectMap = new HashMap<>();
        for (Field field : fields) {
            ExcelField[] annotationsByType = field.getAnnotationsByType(ExcelField.class);
            if (CommonUtils.arrayEmpty(annotationsByType)) {
                continue;
            }
            int i = annotationsByType[0].colIndex();
            if (i >= 0) {
                selectMap.put(annotationsByType[0].colIndex(), field);
            }
        }
        return selectMap;
    }

    private class ExcelReadProperties {
        public String classPathSource;

        public int sheetAt;

        public int skip;

        public int limitRow;
    }

}
