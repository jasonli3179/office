package com.office.controller;

import com.office.mapper.InsertMapper;
import com.office.service.DictService;
import com.office.service.ExcelServiceXLS;
import com.office.service.ExcelServiceXLSX;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.text.SimpleDateFormat;

/**
 * Excel表格操作
 */
@Controller
@RequestMapping("/Excel")
public class ExcelController {


    @Autowired
    private DictService dictService;
    @Autowired
    private ExcelServiceXLS excelServiceXLS;
    @Autowired
    private ExcelServiceXLSX excelServiceXLSX;


    @Autowired
    private InsertMapper insertMapper;

    private SimpleDateFormat format=new SimpleDateFormat("yyyy-MM-dd");


    @RequestMapping(method = RequestMethod.POST, value = "/do")
    @ResponseBody
    public Object doSome(String name) throws Exception {
        String code = dictService.getDict(name);//码值
        //目标目录
        File dir = new File("C:\\Users\\mrli\\Desktop\\files");
        String[] fileNames = dir.list();
        InputStream in = null;
        //写回文件流
        FileOutputStream out = null;
        //遍历当前目录下所有Excel文件
        for (int i = 0; i < fileNames.length; i++) {
            //写入和写出的文件名
            String fileName = "C:\\Users\\mrli\\Desktop\\files\\" + fileNames[i];
            out = new FileOutputStream(fileName);
            //将文件读入
            in = new FileInputStream(new File(fileName));
            //创建工作簿
            //xlsx格式
            if (fileNames[i].substring(fileNames[i].length() - 1).equals("x")) {
                Workbook wb = WorkbookFactory.create(in);
                for (int x = 0; x < 4; x++) {
                    Sheet sheet = wb.getSheetAt(x);

                    excelServiceXLSX.move(sheet, 3, 21);
                    excelServiceXLSX.insertColumn(sheet, 0, name);
                    excelServiceXLSX.insert(sheet, 2, 0, "乡镇名称");

                    excelServiceXLSX.move(sheet, 3, 21);
                    excelServiceXLSX.insertColumn(sheet, 0, code);
                    excelServiceXLSX.insert(sheet, 2, 0, "乡镇代码");
                }
                //写回文件
                out = new FileOutputStream(fileName);
                wb.write(out);
                out.flush();
            } else {//xls格式
                HSSFWorkbook wb = new HSSFWorkbook(in);
                for (int x = 0; x < 4; x++) {
                    HSSFSheet sheet = wb.getSheetAt(x);

                    excelServiceXLS.move(sheet, 0, 21);
                    excelServiceXLS.insertColumn(sheet, 0, name);
                    excelServiceXLSX.insert(sheet, 2, 0, "乡镇名称");

                    excelServiceXLS.move(sheet, 0, 21);
                    excelServiceXLS.insertColumn(sheet, 0, code);
                    excelServiceXLSX.insert(sheet, 2, 0, "乡镇代码");
                }
                //写回文件
                wb.write(out);
                out.flush();
            }
            in.close();
            out.close();
        }


        return "success";
    }


    /**
     * 笔试成绩录入以及处理
     */
//    @PostMapping("/import")
//    @ResponseBody
//    @Transactional(rollbackFor = Exception.class)
//    public String import1(MultipartFile file,int begin,int end,int sheetNum) throws Exception {
//        Workbook wb = null;
//        try {
//            if (isExcel2007(file.getOriginalFilename())) {
//                wb = new XSSFWorkbook(file.getInputStream());
//            } else {
//                wb = new HSSFWorkbook(file.getInputStream());
//            }
//            Sheet sheet = wb.getSheetAt(sheetNum);//获取一张表
//
//            for (int i = begin; i <= end; i++) {
//
//                Map<String, Object> map = new HashMap<>();
//
//                Row row = sheet.getRow(i);//获取索引为i的行
//                //区域编码
//                String areacode = "130824014";//String.valueOf(getValue(row, 0));
//                map.put("areacode", areacode);
//
//                //姓名
//                String name = String.valueOf(getValue(row, 1));
//                map.put("name", name);
//
//                //身份证号
//                String idcard = String.valueOf(getValue(row, 5));
//                map.put("idcard", idcard);
//
//                //性别
//                String sex = String.valueOf(getValue(row, 2)).equals("男") ? "1" : "2";
//                map.put("sex", sex);
//
//                //民族
//                String nation = String.valueOf(getValue(row, 3));
//                if(!nation.contains("族")){
//                    nation+="族";
//                }
//                map.put("nation", nation);
//
//                //户籍性质
//                String householdtype = String.valueOf("城镇".equals(getValue(row, 9)) ? "1" : "2");
//                map.put("householdtype", householdtype);
//
//                //家庭住址
//                String homeaddress = String.valueOf(getValue(row, 16));
//                map.put("homeaddress", homeaddress);
//
//                //电话号码
//                String phoneship = String.valueOf(getValue(row, 17));
//                map.put("phoneship", phoneship);
//
//                //毕业时间
//                String collegestime = String.valueOf(getValue(row, 8));
//                map.put("collegestime", collegestime);
//
//                //文化程度
//                String education = String.valueOf(getValue(row, 6));
//                if(education!=" "){
//                    if(education.contains("中专")){
//                        education="中等专科";
//                    }else if(education.contains("初中")){
//                        education="初级中学";
//                    }else if(education.contains("高中")){
//                        education="职业高中";
//                    }else if(education.contains("小学")){
//                        education="小学";
//                    }else if(education.contains("大专")){
//                        education="大学专科";
//                    }else if(education.contains("中学")){
//                        education="普通中学";
//                    }else if(education.contains("本科")){
//                        education="大学本科";
//                    }
//                }
//                map.put("education", education);
//
//                //高校毕业生
//                String graduate = String.valueOf(getValue(row, 12));
//                map.put("graduate", "是".equals(graduate) ? "0" : "1");
//
//                //就业转失业 0是  1否
//                String transformationwork = String.valueOf(getValue(row, 11));
//                map.put("transformationwork", "是".equals(transformationwork) ? "0" : "1");
//
//                //原就业单位
//                String originaldepart = String.valueOf(getValue(row, 10));
//                map.put("originaldepart", originaldepart);
//
//                //毕业院校
//                String colleges = String.valueOf(getValue(row, 7));
//                map.put("colleges", colleges);
//
//                //就业困难人员   0是  1否
//                String difficultyman = String.valueOf(getValue(row, 13));
//                map.put("difficultyman", "是".equals(difficultyman) ? "0" : "1");
//
//                //健康状况 1健康 2是残疾
//                String healthy = String.valueOf(getValue(row, 14));
//                map.put("healthy", "健康".equals(healthy) ? "1" : "2");
//
//                //求职意向
//                String intention=String.valueOf(getValue(row, 15));
//                map.put("intention", intention);
//
//                //失业登记年月
////                String registertime=String.valueOf(getValue(row, 18));
//                map.put("registertime", "2019-02-01");
//
//                //户籍所在地 1本县 2外县
//                String householdhome = String.valueOf(getValue(row, 4));
//                map.put("householdhome", "本县".equals(householdhome) ? "1" : "2");
//
//                //人员类别1新增就业 2失业在就业 3失业
//                map.put("personnelcategory", "3");
//
//                Long id = IdWorker.getInstance().nextId();
//                map.put("id", id);
//
//                ggRecruitmentinformationDao.insert(map);
//
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        } finally {
//            if (wb != null) {
//                wb.close();
//            }
//        }
//
//        return "BINGO";
//
//    }

    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    public Object getValue(Row row, int i) {
        try {
            return row.getCell(i) == null ? " " : row.getCell(i).getStringCellValue();
        } catch (Exception e) {
            return row.getCell(i).getNumericCellValue();
        }
    }


}
