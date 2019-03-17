package com.office.controller;

import com.office.domain.ColumnParam;
import com.office.mapper.AutoExcelMapper;
import com.office.mapper.GgRecruitmentinformationDao;
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

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
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
    private GgRecruitmentinformationDao ggRecruitmentinformationDao;


    /**
     * Excel半自动导入数据库
     * （注：表格中不能有和数据库字段不匹配的列）
     *
     * @param file      Excel文件
     * @param headBegin Excel表头起始行
     * @param end       数据结束行
     * @param sheetNum  sheet表索引
     * @param tableName 导入数据库表名
     * @return
     * @throws Exception
     */
    @PostMapping("/autoImport")
    @ResponseBody
    @Transactional(rollbackFor = Exception.class)
    public String autoImport(MultipartFile file, Integer headBegin, Integer end, Integer sheetNum, String tableName) throws Exception {
        //序号起始位置减一变下标
        headBegin--;
        end--;
        sheetNum--;

        //获取指定表名的所有字段名和注释
        List<ColumnParam> cac = autoExcelMapper.getColumnAndComment("'" + tableName + "'");

        //字段名对应顺序
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
                        //顺序添加
                        columnNameList.add(new ColumnParam(titleComment, param.getColumnName()));
                        columnNum++;
                    }
                }
            }

            //获取Excel标记行
            Row markRow = sheet.getRow(headBegin - 1);

            //数据导入
            for (int i = headBegin + 1; i <= end; i++) {

                Map<String, Object> insertMap = new HashMap<>();

                Row row = sheet.getRow(i);//获取索引为i的行

                for (int z = 0; z < columnNum; z++) {
                    //过滤后的单元格内容
                    String intention = this.stringFilter(String.valueOf(getValue(row, z)));

                    //判断对应的标记单元格是否有“字典”标识
                    if ("zd".equals(String.valueOf(getValue(markRow, z)))) {
                        intention = "(select code from insured_dict where insured='report' and type='" + columnNameList.get(z).getColumnName() + "' and codeName like '%" + intention + "%' limit 1)";
                        insertMap.put(columnNameList.get(z).getColumnName(), intention);
                    } else {
                        insertMap.put(columnNameList.get(z).getColumnName(), "'" + intention + "'");
                    }
                }

                //主键id
                Long id = IdWorker.getInstance().nextId();
                insertMap.put("id", id);

                ggRecruitmentinformationDao.insert(insertMap);

            }
        } catch (Exception e) {
            e.printStackTrace();
            return "ERROR";
        } finally {
            if (wb != null) {
                wb.close();
            }
        }

        return "BINGO";

    }

    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }


    public Object getValue(Row row, int i) {
        try {
            return row.getCell(i) == null ? "" : row.getCell(i).getStringCellValue();
        } catch (Exception e) {
            return row.getCell(i).getNumericCellValue();
        }
    }

    /**
     * 自定义过滤字符串
     *
     * @param str
     * @return
     */
    public String stringFilter(String str) {
        return str.trim();
    }


    /**
     * 半自动导出Excel
     *
     * @return
     * @throws Exception
     */
    @PostMapping("/autoExport")
    @ResponseBody
    @Transactional(rollbackFor = Exception.class)
    public String autoImport(HttpServletResponse response) throws Exception {
        int sheetNum = 0;
        int headNum = 0;
        String tableName = "tableName";


        //获取数据库字段名和汉化
        List<ColumnParam> cac = autoExcelMapper.getColumnAndComment(tableName);

        //字段名对应顺序
        List<ColumnParam> columnNameList = new ArrayList<>(cac.size());

        //填充excel的数据
        List<Map<String, String>> list = new ArrayList<>();


        InputStream in = this.getClass().getResourceAsStream("/templates/----------------.xls");
        //流转MultipartFile
        MultipartFile file = new MockMultipartFile("----------------.xls.xls", "----------------.xls.xls", "", in);

        Workbook wb = null;
        if (this.isExcel2007(file.getOriginalFilename())) {
            wb = new XSSFWorkbook(file.getInputStream());
        } else {
            wb = new HSSFWorkbook(file.getInputStream());
        }

        Sheet sheet = wb.getSheetAt(sheetNum);

        //获取表头行
        Row titleRow = sheet.getRow(headNum);

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
                    //顺序添加
                    columnNameList.add(new ColumnParam(titleComment, param.getColumnName()));
                    columnNum++;
                }
            }
        }

        //遍历数据
        for (int x = 0; x < list.size(); x++) {

        }


        OutputStream output = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=file.xls");
        response.setContentType("application/msexcel");
        wb.write(output);
        output.close();
    }

}
