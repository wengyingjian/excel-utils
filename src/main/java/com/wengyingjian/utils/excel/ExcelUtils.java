package com.wengyingjian.utils.excel;

import com.wengyingjian.utils.excel.annotations.ExcelCell;
import com.wengyingjian.utils.excel.processors.PageSheetProcessor;
import com.wengyingjian.utils.excel.processors.SheetProcessor;
import com.wengyingjian.utils.excel.processors.SingleSheetProcessor;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {

    /**
     * 导出excel，每个sheet的内容支持分页查询
     */
    public static void exportExcel(OutputStream stream, PageSheetProcessor... processors) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //一个sheet
        for (PageSheetProcessor processor : processors) {
            writeSheet(workbook, processor);
        }
        //写入流
        workbook.write(stream);
        workbook.close();
    }


    /**
     * 导出excel，每个sheet的内容不支持分页查询
     */
    public static <T> void exportExcel(OutputStream stream, SingleSheetProcessor... processors) throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //一个sheet
        for (SingleSheetProcessor processor : processors) {
            writeSheet(workbook, processor);
        }
        //写入流
        workbook.write(stream);
        workbook.close();
    }

    private static void writeSheet(Workbook workbook, SheetProcessor processor) {
        //createSheet
        Sheet sheet = createSheet(workbook, processor.getSheetName());
        // headLine
        int rowNum = writeHeadLine(sheet, processor.getHeadLine());
        //data
        writeBody(sheet, rowNum, processor);
    }

    private static <T> void writeBody(Sheet sheet, int rowNum, SheetProcessor processor) {
        int currentPage = 1;
        if (processor instanceof PageSheetProcessor) {
            PageSheetProcessor pageProcessor = (PageSheetProcessor) processor;
            int pageSize = pageProcessor.getPageSize();
            while (true) {
                //（分页）一页数据
                List<T> lines = pageProcessor.getDataByPage(currentPage++, pageSize);
                if (CollectionUtils.isEmpty(lines)) {
                    break;
                }
                for (T t : lines) {
                    //一行数据
                    Row row = sheet.createRow(rowNum++);
                    writeLine(row, t);
                }
            }
        }
        if (processor instanceof SingleSheetProcessor) {
            SingleSheetProcessor singleProcessor = (SingleSheetProcessor) processor;
            //（分页）一页数据
            List<T> lines = singleProcessor.getData();
            if (CollectionUtils.isEmpty(lines)) {
                return;
            }
            for (T t : lines) {
                //一行数据
                Row row = sheet.createRow(rowNum++);
                writeLine(row, t);
            }
        }
    }

    /**
     * 创建一个sheet页
     *
     * @param sheetName sheet名 可选
     */
    private static Sheet createSheet(Workbook workbook, String sheetName) {
        if (sheetName == null) {
            return workbook.createSheet();
        }
        return workbook.createSheet(sheetName);
    }

    /**
     * 写入表头（第一行）
     *
     * @return 0无表头，1写入第一行表头
     */
    private static int writeHeadLine(Sheet sheet, List<String> headLine) {
        if (headLine == null) {
            return 0;
        }
        Row row = sheet.createRow(0);
        for (int i = 0; i < headLine.size(); i++) {
            Cell cell = row.createCell(i);
            setCellValue(cell, headLine.get(i));
        }
        return 1;
    }

    private static <T> void writeLine(Row row, T t) {
        if (t instanceof List) {
            int cellIndex = 0;
            List line = (List) t;
            for (Object cell : line) {
                setCellValue(row.createCell(cellIndex++), cell);
            }
        } else {
            //get by annotation
            Class clazz = t.getClass();
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                ExcelCell anno = field.getAnnotation(ExcelCell.class);
                if (anno == null) {
                    continue;
                }
                int index = anno.index();
                try {
                    setCellValue(row.createCell(index), field.get(t));
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }


    private static void setCellValue(Cell cell, Object value) {
        if (value == null || "null".equals(value)) {
            cell.setCellValue("");
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) value));
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }

}
