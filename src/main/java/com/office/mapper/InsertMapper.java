package com.office.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Map;

@Mapper
public interface InsertMapper {

	Integer insert(@Param("map") Map<String,Object> map,@Param("tableName") String tableName);






}