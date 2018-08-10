package com.wengyingjian.utils.excel.processors;

import java.util.List;

public abstract class PageSheetProcessor<T> implements SheetProcessor {

    private String sheetName;

    private List<String> headLine;

    private int pageSize;

    public PageSheetProcessor(String sheetName, List<String> headLine, int pageSize) {
        this.sheetName = sheetName;
        this.headLine = headLine;
        this.pageSize = pageSize;
    }

    public PageSheetProcessor(List<String> headLine, int pageSize) {
        this.headLine = headLine;
        this.pageSize = pageSize;
    }


    /**
     * 批量获取一页数据，以emptyList作为截止标志
     *
     * @param currentPage 当前分页，从1开始
     * @param pageSize    每页大小
     * @return 一页数据
     */
    public abstract List<T> getDataByPage(int currentPage, int pageSize);


    @Override
    public String getSheetName() {
        return sheetName;
    }

    public int getPageSize() {
        return pageSize;
    }

    @Override
    public List<String> getHeadLine() {
        return headLine;
    }

}
