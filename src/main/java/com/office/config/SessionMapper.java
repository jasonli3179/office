package com.office.config;

import com.office.mapper.AutoExcelMapper;
import org.apache.ibatis.session.SqlSessionFactory;
import org.mybatis.spring.support.SqlSessionDaoSupport;
import org.springframework.stereotype.Repository;

import javax.annotation.Resource;
import java.util.Map;

/**
 * 用于session查询
 * @author lc
 */
@Repository
public class SessionMapper extends SqlSessionDaoSupport {

    @Resource
    public void setSqlSessionFactory(SqlSessionFactory sqlSessionFactory) {
        super.setSqlSessionFactory(sqlSessionFactory);
    }

    /**
     * @return
     */
    @SuppressWarnings("unchecked")
    public Map<String,String> getDict(){
        ResultHandler handler = new ResultHandler();
        //namespace : XxxMapper.xml 中配置的地址（XxxMapper.xml的qualified name）
        //.selectXxxxNum : XxxMapper.xml 中配置的方法名称
        //this.getSqlSession().select(namespace+".selectXxxxNum", handler);
        this.getSqlSession().select(AutoExcelMapper.class.getName()+".getDict", handler);
        Map<String, String> map = handler.getMappedResults();
        return map;
    }
}
