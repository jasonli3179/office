package com.office.config;

import org.apache.ibatis.session.ResultContext;

import java.util.HashMap;
import java.util.Map;

/**
 * 用于转化结果集
 *
 * @author lc
 */
public class ResultHandler implements org.apache.ibatis.session.ResultHandler {

    @SuppressWarnings("rawtypes")
    private final Map mappedResults = new HashMap();

    @SuppressWarnings("unchecked")
    @Override
    public void handleResult(ResultContext context) {
        @SuppressWarnings("rawtypes")
        Map map = (Map) context.getResultObject();
        mappedResults.put(map.get("key"), map.get("value")); // xml配置里面的property的值，对应的列
    }

    @SuppressWarnings("rawtypes")
    public Map getMappedResults() {
        return mappedResults;
    }

}
