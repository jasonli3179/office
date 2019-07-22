package com.office.controller;

import com.office.domain.ColumnParam;
import com.office.mapper.AutoExcelMapper;
import com.office.mapper.InsertMapper;
import com.office.util.IdWorker;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel表格半自动操作
 */
@Controller
@RequestMapping("/Excel")
public class AutoExcelController {

    @Autowired
    private AutoExcelMapper autoExcelMapper;
    @Autowired
    private InsertMapper insertMapper;


    private SimpleDateFormat format=new SimpleDateFormat("yyyy-MM-dd");


    /**
     * Excel半自动导入数据库
     * （注：表格中不能有和数据库字段不匹配的列或多余的列）
     * 接口带出错回滚数据功能
     *
     * @param file      Excel文件
     * @param headBegin Excel表头起始行
     * @param sheetNum  sheet表索引
     * @param tableName 导入数据库表名（可多个逗号分隔）
     * @param schemaName 数据库名
     * @return
     * @throws Exception
     */
    @PostMapping("/autoImport")
    @ResponseBody
    @Transactional(rollbackFor = Exception.class)
    public String autoImport(MultipartFile file, Integer headBegin, Integer sheetNum, String tableName,String schemaName) throws Exception {

        //从excel读取数据总数
        int dataNum=0;

        //序号起始位置减一变下标
        headBegin--;
        sheetNum--;


        //获取指定表名的所有字段名和注释
        List<ColumnParam> cac = autoExcelMapper.getColumnAndComment("'"+tableName+"'","'"+schemaName+"'");

        //数据库字段名对应excel头顺序
        List<ColumnParam> columnNameList = new ArrayList<>(cac.size());

        Workbook wb = null;
        try {
            if (isExcel2007(file.getOriginalFilename())) {
                wb = new XSSFWorkbook(file.getInputStream());
            } else {
                wb = new HSSFWorkbook(file.getInputStream());
            }
            Sheet sheet = wb.getSheetAt(sheetNum);//获取一张表

            //获取表头行
            Row titleRow = sheet.getRow(headBegin);

            //excel有效数据列数
            int columnNum = 0;

            //确定excel头部类型字段对应的表字段顺序
            for (int x = 0; true; x++) {
                //过滤后的单元格内容
                String titleComment = this.stringFilter(String.valueOf(getValue(titleRow, x)));
                //判断若到行末尾，则终止
                if ("".equals(titleComment)) {
                    break;
                }
                //获取注释对应的字段名
                for (int y = 0; y < cac.size(); y++) {
                    ColumnParam param = cac.get(y);
                    //判断若表头注释和数据库注释相似
                    if (titleComment.contains(param.getColumnComment()) || param.getColumnComment().contains(titleComment)) {
//                    if (titleComment.equals(param.getColumnComment())) {
                        //顺序添加
                        columnNameList.add(new ColumnParam(titleComment, param.getColumnName()));
                        columnNum++;
                    }
                }
            }


            //获取Excel标记行
            Row markRow = sheet.getRow(headBegin - 1);

            //数据导入
            for (int i = headBegin + 1; true; i++) {

                //要插入的列和值
                Map<String, Object> insertMap = new HashMap<>();

                Row row = sheet.getRow(i);//获取索引为i的行

                //判断是否到最后一条
                if(row==null){
                    break;
                }

                //遍历一行数据
                for (int z = 0; z < columnNum; z++) {
                    //过滤后的单元格内容
                    String intention = this.stringFilter(String.valueOf(getValue(row, z)));

                    //若单元格值为空则跳过
                    if(intention==null || intention.equals("")){
                        continue;
                    }


                    //区域编码字典表转换sql
                    if ("a".equals(String.valueOf(getValue(markRow, z)))) {
                        intention = "(select townTownCode from agencies where townTownName like '%" + intention + "%' limit 1)";
                        insertMap.put(columnNameList.get(z).getColumnName(), intention);
                        continue;
                    }
//
//                    //判断对应的标记单元格是否有“zd”标识
//                    //insured_dict字典表
                    //根据数据库表列名和值查出对应码值
                    if ("zd".equals(String.valueOf(getValue(markRow, z)))) {
                        intention = "(select code from insured_dict where insured='report' and type='" + columnNameList.get(z).getColumnName() + "' and codeName like '%" + intention + "%' limit 1)";
                        insertMap.put(columnNameList.get(z).getColumnName(), intention);
                    } else {
                        insertMap.put(columnNameList.get(z).getColumnName(), "'" + intention + "'");
                    }
                }
                //插入一条数据
                //主键id
                Long id = IdWorker.getInstance().nextId();
                insertMap.put("id", id);
                insertMap.put("personnelcategory", "'1'");//添加字符串类型参数需要自行添加单引号
                insertMap.put("departcategory", "'1'");
                insertMap.put("phasetime", "'2019-07-16'");
                insertMapper.insert(insertMap,tableName);



                //成功导入一条数据加一
                dataNum++;

            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("ERROR---------"+dataNum);
        } finally {
            if (wb != null) {
                wb.close();
            }
        }

        return "BINGO---------"+dataNum;

    }

    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }


    //获取单元格数值
    public Object getValue(Row row, int i) {
        try {
            return row.getCell(i) == null ? "" : row.getCell(i).getStringCellValue();
        } catch (Exception e) {
            return row.getCell(i).getNumericCellValue();
        }
    }

    /**
     * 自定义过滤字符串
     * 去掉单元格值两边的空格以免影响判断
     *
     * @param str
     * @return
     */
    public String stringFilter(String str) {
        return str.trim();
    }


//    /**
//     * 半自动导出Excel
//     *
//     * @return
//     * @throws Exception
//     */
//    @PostMapping("/autoExport")
//    @ResponseBody
//    @Transactional(rollbackFor = Exception.class)
//    public void autoImport(HttpServletResponse response) throws Exception {
//        int sheetNum = 0;
//        int headNum = 0;
//        //多个表则逗号分隔
//        String tableName = "tableName";
//
//
//        //获取数据库字段名和汉化*
//        List<ColumnParam> cac = autoExcelMapper.getColumnAndComment(tableName);
//
//        //字段名excel对应顺序
//        List<ColumnParam> columnNameList = new ArrayList<>(cac.size());
//
//        //填充excel的数据*
//        List<Map<String, String>> list = new ArrayList<>();
//
//
//        InputStream in = this.getClass().getResourceAsStream("/templates/----------------.xls");
//        //流转MultipartFile
//        MultipartFile file = new MockMultipartFile("----------------.xls.xls", "----------------.xls.xls", "", in);
//
//        Workbook wb = null;
//        if (this.isExcel2007(file.getOriginalFilename())) {
//            wb = new XSSFWorkbook(file.getInputStream());
//        } else {
//            wb = new HSSFWorkbook(file.getInputStream());
//        }
//
//        Sheet sheet = wb.getSheetAt(sheetNum);
//
//        //获取表头行
//        Row titleRow = sheet.getRow(headNum);
//
//        //excel有效数据列数
//        int columnNum = 0;
//
//        //确定excel头部类型字段对应的表字段顺序
//        for (int x = 1; true; x++) {
//            //过滤后的单元格内容
//            String titleComment = this.stringFilter(String.valueOf(getValue(titleRow, x)));
//            //判断若到行末尾，则终止
//            if ("".equals(titleComment)) {
//                break;
//            }
//            //获取注释对应的字段名
//            for (int y = 0; y < cac.size(); y++) {
//                ColumnParam param = cac.get(y);
//                //判断若表头注释和数据库注释相似
//                if (titleComment.contains(param.getColumnComment()) || param.getColumnComment().contains(titleComment)) {
//                    //顺序添加
//                    columnNameList.add(new ColumnParam(titleComment, param.getColumnName()));
//                    columnNum++;
//                }
//            }
//        }
//        //序号
//        int index=1;
//
//        //遍历数据
//        for (int x = 0; x < list.size(); x++) {
//
//            //一条正文数据
//            Map<String, String> map = list.get(x);
//
//            //获取正文行
//            Row row = sheet.getRow(headNum + 1);
//
//            //序号单元格
//            Cell indexCell = row.createCell(0);
//            indexCell.setCellValue(index++);
//
//            //循环插入excel所列字段一行
//            for (int y = 1; y < columnNameList.size(); y++) {
//                Cell cell = row.createCell(y);
//                cell.setCellValue(map.get(columnNameList.get(y).getColumnName()));
//            }
//        }
//
//
//        OutputStream output = response.getOutputStream();
//        response.reset();
//        response.setHeader("Content-disposition", "attachment; filename=file.xls");
//        response.setContentType("application/msexcel");
//        wb.write(output);
//        output.close();
//    }

}
