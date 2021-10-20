package com.yx.dwdb.controller;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.yx.dwdb.YxDwdbApplication;
import com.yxkj.common.core.config.SpringContextUtils;
import com.yxkj.common.core.util.other.StringUtils;
import com.yxkj.dwdb_common.mapper.DwdbBasicMapper;
import com.yxkj.dwdb_common.pojo.DwdbBasic;
import com.yxkj.dwdb_common.pojo.DwdbPhotoInputPojo;
import com.yxkj.dwdb_common.service.IDwdbBasicService;
import com.yxkj.dwdb_common.utils.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.transaction.annotation.Transactional;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = YxDwdbApplication.class)

public class DemoPictureInPut {

    @Autowired
    private DwdbBasicMapper dwdbBasicMapper;
    @Autowired
    private IDwdbBasicService basicService;

    public DwdbBasicMapper getDwdbBasicMapper() {
        return dwdbBasicMapper;
    }

    public void setDwdbBasicMapper(DwdbBasicMapper dwdbBasicMapper) {
        this.dwdbBasicMapper = dwdbBasicMapper;
    }

    public IDwdbBasicService getBasicService() {
        return basicService;
    }

    public void setBasicService(IDwdbBasicService basicService) {
        this.basicService = basicService;
    }


    public List<DwdbBasic> phtotImport(String filePath,String i) {
        System.out.println("遍历"+i+"表");
        List<DwdbBasic> dwdbBasicList =new ArrayList<DwdbBasic>();
        Map<String, PictureData> pictures = new HashMap<>();
        try {
            FileInputStream fio = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fio);
            //得到工作表
            XSSFSheet sheet = workbook.getSheet("s");
            pictures = ExcelUtil.getSheetPictrues07((XSSFSheet) sheet, (XSSFWorkbook) workbook);

            for (Row row : sheet) {
                if (row.getRowNum() == 0 || row.getRowNum() == 1) {// 首行表头不读取
                    continue;
                }
                DwdbBasic dwdbBasic =new DwdbBasic();
                //获取图片
                //dwdbPhotoInputPojo.setPicture(row.getCell(0).getStringCellValue());
                dwdbBasic.setName(row.getCell(1).getStringCellValue());
                dwdbBasic.setIdnumber(row.getCell(2).getStringCellValue());
                try{
                    dwdbBasic.setPicture(ExcelUtil.printImg(pictures,String.valueOf( row.getRowNum())));
                } catch (Exception e){
                    System.out.println("名字:"+dwdbBasic.getName()+";    身份证:"+dwdbBasic.getIdnumber()+"    第"+row.getRowNum()+"行---------------------跳过");
                    continue;
                }

                //修改attachment表对应数据,修改basic表对应数据
                dwdbBasicMapper.updatePicture(dwdbBasic);
                dwdbBasic.setId(dwdbBasicMapper.getdwdbBasicID(dwdbBasic.getName(),dwdbBasic.getIdnumber()));
                dwdbBasicMapper.updateAttachment("dwdb_basic",dwdbBasic.getId(),dwdbBasic.getPicture());
                System.out.println("名字:"+dwdbBasic.getName()+";    身份证:"+dwdbBasic.getIdnumber()+";    图片ID:"+dwdbBasic.getPicture()+"    第"+row.getRowNum()+"行");
                dwdbBasicList.add(dwdbBasic);
            }

            workbook.close();
            fio.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return dwdbBasicList;
    }

    @Test
    public void userImport(){
        IDwdbBasicService bean = SpringContextUtils.getBean(IDwdbBasicService.class);
        DemoPictureInPut demoPictureInPut=new DemoPictureInPut();
        demoPictureInPut.setDwdbBasicMapper(dwdbBasicMapper);
        demoPictureInPut.setBasicService(basicService);
        int i=4;
        //for(int i=27;i<35;i++){
            List<DwdbBasic> dwdbBasicList =  demoPictureInPut.phtotImport("C:\\Users\\yxkj\\Desktop\\导出\\"+i+".xlsx",""+i+"");
       // }
        /*for (DwdbBasic dwdbBasic : dwdbBasicList) {
            System.out.println("名字:"+dwdbBasic.getName()+";    身份证:"+dwdbBasic.getIdnumber()+";    图片ID:"+dwdbBasic.getPicture());
        }*/
    }




}
