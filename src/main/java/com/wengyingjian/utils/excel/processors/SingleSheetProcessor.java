package com.wengyingjian.utils.excel.processors;

import java.util.List;

public abstract class SingleSheetProcessor<T> implements SheetProcessor {

    private String sheetName;

    private List<String> headLine;


    public SingleSheetProcessor(String sheetName, List<String> headLine) {
        this.sheetName = sheetName;
        this.headLine = headLine;
    }

    public SingleSheetProcessor(List<String> headLine) {
        this.headLine = headLine;
    }

    abstract public List<T> getData();

    @Override
    public String getSheetName() {
        return sheetName;
    }

    @Override
    public List<String> getHeadLine() {
        return headLine;
    }
}
