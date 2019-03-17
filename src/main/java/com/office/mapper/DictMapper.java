package com.office.mapper;

import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

@Mapper
public interface DictMapper {

    @Select("select code from insured_dict where codeName=#{code} and insured='up'")
    String getDict(String code);

}
