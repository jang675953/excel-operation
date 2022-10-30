package com.klein.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.klein.easyexcel.domain.annotation.User;
import lombok.extern.slf4j.Slf4j;

import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@Slf4j
public class CustomReadListener implements ReadListener<User> {
    private AtomicInteger count = new AtomicInteger(0);

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        log.debug("read excel start, sheet: {}", context.readSheetHolder().getSheetName());
        log.debug("read head: {}", headMap.values().stream().map(ReadCellData::getStringValue).collect(Collectors.toList()));
    }

    @Override
    public void invoke(User o, AnalysisContext analysisContext) {
        log.debug("read data: {}", o);
        count.incrementAndGet();
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("read excel finished!");
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        log.debug("read excel failed!");
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
            log.error("第{}行，第{}列解析异常，数据为: {}", excelDataConvertException.getRowIndex(),
                    excelDataConvertException.getColumnIndex(), excelDataConvertException.getCellData());
        }
    }

    public int getCount() {
        return count.get();
    }
}
