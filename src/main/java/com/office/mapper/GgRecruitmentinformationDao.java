package com.office.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Update;

import java.util.Map;

@Mapper
public interface GgRecruitmentinformationDao {

	Integer insert(@Param("map") Map<String,Object> map);

	@Update("update obtain_town_entry set phasetime=#{phasetime} where idcard=#{idcard}")
	Integer update(Map<String,Object> map);

	
}