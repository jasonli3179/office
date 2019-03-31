package com.office.mapper;

import com.office.domain.ColumnParam;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface AutoExcelMapper {

    /**
     * 获取指定表名的所有字段名和注释
     * @return
     */
    List<ColumnParam> getColumnAndComment(@Param("tableName") String tableName,@Param("schemaName") String schemaName);


}
