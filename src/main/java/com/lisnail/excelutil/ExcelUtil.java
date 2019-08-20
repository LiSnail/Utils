package com.lisnail.excelutil;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

/**
 * @author liyong
 * @date 2019.3.20
 */

public class ExcelUtil<T> {
    /**
     * 时间格式
     */
    private final static String TIME_PATTERN = "yyyy-MM-dd";
    /**
     * 默认列的宽度
     */
    private final static int COLUMN_WIDTH = 20;
    /**
     * 表头字体大小
     */
    private final static int TITLE_FONT_SIZE = 11;
    /**
     * 表头单元格
     **/
    private final static int HEADER_CELL = 0;
    /**
     * 内容单元格
     **/
    private final static int CONTENT_CELL = 1;
    /**
     * title 单元格
     */
    private final static int TITLE_CELL = 2;
    /**
     * 复杂表头单元格的默认宽度
     */
    private final static int COLUMN_COMPLEX_WIDTH = 6;

    private void creatCaption(int rowStart, HSSFSheet sheet, HSSFWorkbook workbook) {
        HSSFCellStyle captionStyle = setCellStyle(workbook, TITLE_CELL);
        HSSFRow row = sheet.createRow(rowStart - 3);
        HSSFCell cell1 = row.createCell(1);
        cell1.setCellValue("名称：");
        cell1.setCellStyle(captionStyle);
        cell1 = row.createCell(2);
        cell1.setCellValue("lisnail");
        cell1.setCellStyle(captionStyle);
        //cell1.setCellStyle(getCaptionStyle());
        row = sheet.createRow(rowStart - 2);
        HSSFCell cell2 = row.createCell(1);
        cell2.setCellValue("id：");
        cell2.setCellStyle(captionStyle);
        cell2 = row.createCell(2);
        cell2.setCellValue("123");
        cell2.setCellStyle(captionStyle);
    }

    /**
     * 设置样式
     **/
    private HSSFCellStyle setCellStyle(HSSFWorkbook workbook, int cellType) {
        HSSFCellStyle style = null;
        if (cellType == HEADER_CELL) {
            style = workbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            style.setBorderRight(HSSFCellStyle.BORDER_THIN);
            style.setBorderTop(HSSFCellStyle.BORDER_THIN);
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setTopBorderColor(HSSFColor.BLACK.index);
            style.setLeftBorderColor(HSSFColor.BLACK.index);
            style.setRightBorderColor(HSSFColor.BLACK.index);
            style.setBottomBorderColor(HSSFColor.BLACK.index);
            //垂直居中
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            //自动换行
            style.setWrapText(true);
            // 生成标题字体
            HSSFFont font = workbook.createFont();
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            font.setFontName("宋体");
            font.setColor(HSSFColor.WHITE.index);
            font.setFontHeightInPoints((short) TITLE_FONT_SIZE);
            // 把字体应用到当前的样式
            style.setFont(font);
        } else if (cellType == CONTENT_CELL) {
            style = workbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.WHITE.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            style.setBorderRight(HSSFCellStyle.BORDER_THIN);
            style.setBorderTop(HSSFCellStyle.BORDER_THIN);
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            //设置自动换行
            style.setWrapText(true);
            // 生成内容字体
            HSSFFont font = workbook.createFont();
            font.setFontName("宋体");
            font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
            style.setFont(font);
        } else if (cellType == TITLE_CELL) {
            style = workbook.createCellStyle();
            //设置单元格填充色
            style.setFillForegroundColor(HSSFColor.LIME.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            //居中对齐
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            HSSFFont font = workbook.createFont();
            font.setFontHeightInPoints((short) 12);
            //设置字体名字
            font.setFontName("宋体");
            style.setFont(font);
        }
        return style;
    }


    /**
     * 导出有复杂表头的excel表格
     *
     * @param rowStart    表格从第几行开始创建
     * @param sheetName   sheet名称
     * @param headers     表头
     * @param headNums    表头分布
     * @param dataSet     填充数据
     * @param headerWords 表格头部标题的字段名
     * @return
     */
    public InputStream excelContent(Integer rowStart, String sheetName, List<String[]> headers, List<String[]>
            headNums, Collection<T> dataSet, String[] headerWords) throws IOException {
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个sheet
        HSSFSheet sheet = workbook.createSheet(sheetName);

        //创建title
        creatCaption(rowStart, sheet, workbook);

        // 生成表头样式
        HSSFCellStyle titleStyle = setCellStyle(workbook, HEADER_CELL);
        // 设置表格默认列宽度
        sheet.setDefaultColumnWidth(COLUMN_COMPLEX_WIDTH);
        //制作表头
        String[] header;
        String[] headNum;
        int length = headers.size();
        for (int i = 0; i < length; i++) {
            header = headers.get(i);
            headNum = headNums.get(i);
            HSSFRow row = sheet.createRow(i + rowStart);
            HSSFCell cellHeader;
            for (int j = 0; j < header.length; j++) {
                sheet.setColumnWidth(j + 1, sheet.getColumnWidth(j + 1) * COLUMN_WIDTH / 10);
                cellHeader = row.createCell(j);
                cellHeader.setCellStyle(titleStyle);
                cellHeader.setCellValue(new HSSFRichTextString(header[j]));
            }
            // 动态合并单元格
            for (int j = 0; j < headNum.length; j++) {
                String[] temp = headNum[j].split(",");
                int startRow = Integer.parseInt(temp[0]) + rowStart;
                int overRow = Integer.parseInt(temp[1]) + rowStart;
                int startCol = Integer.parseInt(temp[2]);
                int overCol = Integer.parseInt(temp[3]);
                if (overRow - startRow > 0 || overCol - startCol > 0) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, overRow, startCol, overCol));
                }
            }
        }
        /*填充表格内容
        rowStart+headers.size() 表格开始行 + 表头所占行
        */
        fillContent(headers.get(0).length, rowStart + headers.size(), headerWords, dataSet, workbook, sheet);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        return new ByteArrayInputStream(bos.toByteArray());
    }

    /**
     * 填充sheet内容
     *
     * @param contentRowStart 表格的内容从哪一行开始填充
     * @param headerWords     填充数据的字段
     * @param dataSet         数据集
     * @param workbook        excelSheet上下文
     * @param sheet           sheet
     */
    private void fillContent(int column, int contentRowStart, String[] headerWords, Collection<T> dataSet, HSSFWorkbook workbook, HSSFSheet sheet) {
        // 内容样式
        HSSFCellStyle contentStyle = setCellStyle(workbook, CONTENT_CELL);
        Iterator<T> it = dataSet.iterator();
        int index = 0;
        T t;
        //字段名
        String fieldName;
        //get方法名称
        String getMethodName;
        //单元格
        HSSFCell cell;
        Class tCls;
        Method getMethod;
        Object value;
        HSSFRow row;

        while (it.hasNext()) {
            row = sheet.createRow(contentRowStart + index);
            t = it.next();
            /*如果传入T的类型是基本的数据类型，headerWords 为null或者size为0，如果是包装类型，
            那么就需要传入包装类型所对应的*/
            if (null != headerWords && headerWords.length > 0) {
                for (int i = 0; i < column; i++) {
                    cell = row.createCell(i);
                    cell.setCellStyle(contentStyle);
                    fieldName = headerWords[i];
                    getMethodName = "get" + fieldName.substring(0, 1).toUpperCase()
                            + fieldName.substring(1);
                    try {
                        tCls = t.getClass();
                        getMethod = tCls.getMethod(getMethodName, new Class[]{});
                        value = getMethod.invoke(t, new Object[]{});
                        formatCell(cell, value);
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    } catch (InvocationTargetException e) {
                        e.printStackTrace();
                    }
                }
            } else {
                for (int i = 0; i < column; i++) {
                    cell = row.createCell(i);
                    cell.setCellStyle(contentStyle);
                    if (i != column - 1) {
                        formatCell(cell, t);
                    }
                }
            }
            index++;
        }
    }

    /**
     * 单元格值统一格式化
     *
     * @param cell
     */
    private static <E> void formatCell(HSSFCell cell, E t) {
        HSSFRichTextString richString = null;
        SimpleDateFormat sdf = new SimpleDateFormat(TIME_PATTERN);
        String textValue = null;
        if (t instanceof Integer) {
            cell.setCellValue((Integer) t);
        } else if (t instanceof Float) {
            textValue = String.valueOf((Float) t);
            cell.setCellValue(textValue);
        } else if (t instanceof Double) {
            textValue = String.valueOf((Double) t);
            cell.setCellValue(textValue);
        } else if (t instanceof Long) {
            cell.setCellValue((Long) t);
        }
        if (t instanceof Boolean) {
            textValue = "是";
            if (!(Boolean) t) {
                textValue = "否";
            }
        } else if (t instanceof Date) {
            textValue = sdf.format((Date) t);
        } else {
            // 其它数据类型都当作字符串简单处理
            if (t != null) {
                textValue = t.toString();
            }
        }
        if (textValue != null) {
            richString = new HSSFRichTextString(textValue);
        }
        cell.setCellValue(richString);
    }


    /**
     * 导出excel
     *
     * @param inputStream 文件流
     * @param filePath    文件路径
     * @param fileName    导出文件的名称
     */

    public void export(InputStream inputStream, String filePath, String fileName) {
        FileOutputStream out = null;
        File file = new File(filePath);
        try {
            if (!file.exists()) {
                file.createNewFile();
            }
            String fileUrl = joinPath(filePath, fileName + ".xls");
            out = new FileOutputStream(new File(fileUrl));
            int c;
            while ((c = inputStream.read()) != -1) {
                out.write(c);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null)
                    out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    private static String joinPath(String url, String path) {
        String filePath;
        if (url.lastIndexOf(File.pathSeparator) == -1) {
            filePath = url + File.separator + path;
        } else {
            filePath = url + path;
        }
        return filePath;
    }

}

