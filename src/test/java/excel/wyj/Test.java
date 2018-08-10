package excel.wyj;

import com.wengyingjian.utils.excel.processors.PageSheetProcessor;
import com.wengyingjian.utils.excel.ExcelUtils;
import com.wengyingjian.utils.excel.processors.SingleSheetProcessor;
import excel.Model;

import java.io.*;
import java.util.*;

public class Test {


    @org.junit.Test
    public void testPageList() throws IOException {
        String sheetName = "sheet1";
        List<String> headLine = Arrays.asList("标题1", "标题2");

        File f = new File("testPageList.xlsx");
        OutputStream out = new FileOutputStream(f);
        ExcelUtils.exportExcel(out, new PageSheetProcessor(sheetName, headLine, 50) {
            @Override
            public List getDataByPage(int currentPage, int pageSize) {
                if (currentPage > 10) {
                    return Collections.emptyList();
                }
                List<List<String>> dataList = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    List<String> line = new ArrayList<>();
                    for (int j = 0; j < 10; j++) {
                        line.add("分页" + currentPage + ":" + j);
                    }
                    dataList.add(line);
                }
                return dataList;
            }
        });
        out.close();
    }

    @org.junit.Test
    public void testPageBean() throws IOException {
        String sheetName = "sheet1";
        List<String> headLine = Arrays.asList("标题1", "标题2");

        File f = new File("testPageBean.xlsx");
        OutputStream out = new FileOutputStream(f);
        ExcelUtils.exportExcel(out, new PageSheetProcessor(sheetName, headLine, 50) {
            @Override
            public List getDataByPage(int currentPage, int pageSize) {
                if (currentPage > 10) {
                    return Collections.emptyList();
                }
                List<Model> dataList = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    Model bean = new Model("a", 1, "c", new Date());
                    dataList.add(bean);
                }
                return dataList;
            }
        });
        out.close();
    }

    @org.junit.Test
    public void testSingleList() throws IOException {
        String sheetName = "sheet1";
        List<String> headLine = Arrays.asList("标题1", "标题2");

        File f = new File("testSingleList.xlsx");
        OutputStream out = new FileOutputStream(f);

        ExcelUtils.exportExcel(out, new SingleSheetProcessor(sheetName, headLine) {
            @Override
            public List getData() {
                List<List<String>> dataList = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    List<String> line = new ArrayList<>();
                    for (int j = 0; j < 10; j++) {
                        line.add(String.valueOf(j));
                    }
                    dataList.add(line);
                }
                return dataList;
            }
        });
    }

    @org.junit.Test
    public void testSingleBean() throws IOException {
        String sheetName = "sheet1";
        List<String> headLine = Arrays.asList("标题1", "标题2");

        File f = new File("testSingleBean.xlsx");
        OutputStream out = new FileOutputStream(f);

        ExcelUtils.exportExcel(out, new SingleSheetProcessor(sheetName, headLine) {
            @Override
            public List getData() {
                List<Model> dataList = new ArrayList<>();
                for (int i = 0; i < 10; i++) {
                    Model bean = new Model("a", 1, "c", new Date());
                    dataList.add(bean);
                }
                return dataList;
            }
        });
    }
}
