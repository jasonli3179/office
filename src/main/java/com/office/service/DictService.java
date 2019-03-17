package com.office.service;

import com.office.mapper.DictMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class DictService {

    @Autowired
    private DictMapper dictMapper;

    public String getDict(String code){
        return dictMapper.getDict(code);
    }

}
