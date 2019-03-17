package com.office.domain;


import lombok.Data;

@Data
public class ColumnParam {

    /**
     * 字段注释
     */
    private String columnComment;

    /**
     * 字段名
     */
    private String columnName;

    public ColumnParam(String columnComment, String columnName) {
        this.columnComment = columnComment;
        this.columnName = columnName;
    }

    public String getColumnComment() {
        return columnComment;
    }

    public void setColumnComment(String columnComment) {
        this.columnComment = columnComment;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }
}
