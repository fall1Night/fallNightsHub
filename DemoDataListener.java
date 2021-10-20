package com.yx.dwdb.controller;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.conditions.update.UpdateWrapper;
import com.yx.dwdb.YxDwdbApplication;
import com.yxkj.common.core.config.SpringContextUtils;
import com.yxkj.common.core.entity.sys.SysAttachment;
import com.yxkj.common.core.enums.StaticParameters;
import com.yxkj.common.core.util.common.TableUtils;
import com.yxkj.common.core.util.encryptionTools.AesEncryptUtil;
import com.yxkj.common.core.util.httpUtils.HttpResult;
import com.yxkj.dwdb_common.mapper.DwdbBasicMapper;
import com.yxkj.dwdb_common.mapper.SysAttachmentMapper;
import com.yxkj.dwdb_common.pojo.DwdbBasic;
import com.yxkj.dwdb_common.pojo.DwdbPhotoInputPojo;
import com.yxkj.dwdb_common.service.IDwdbBasicService;
import net.coobird.thumbnailator.Thumbnails;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.stereotype.Service;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static com.yxkj.dwdb_common.utils.picUploadUtils.getFloadPath;

/*import com.alibaba.easyexcel.test.util.TestFileUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.converters.DefaultConverterLoader;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.read.listener.PageReadListener;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.fastjson.JSON;*/


/**
 * <p>
 * 前端控制器
 * </p>
 *
 * @author zqy
 * @since 2021-09-18
 */

/*@RestController
@RequestMapping("/dwdb-basic1")*/
@RunWith(SpringRunner.class)
@SpringBootTest(classes = YxDwdbApplication.class)
public class DemoDataListener extends AnalysisEventListener<DwdbPhotoInputPojo> {
    @Autowired
    private IDwdbBasicService basicService;
    @Autowired
    private DwdbBasicMapper dwdbBasicMapper;
    @Autowired
    private SysAttachmentMapper sysAttachmentMapper;

    //开始压缩的大小
    private final static long START_COMPRESS = 1024 * 1024;

    public IDwdbBasicService getBasicService() {
        return basicService;
    }

    public void setBasicService(IDwdbBasicService basicService) {
        this.basicService = basicService;
    }

    public DwdbBasicMapper getDwdbBasicMapper() {
        return dwdbBasicMapper;
    }

    public void setDwdbBasicMapper(DwdbBasicMapper dwdbBasicMapper) {
        this.dwdbBasicMapper = dwdbBasicMapper;
    }

    /*@Autowired
            private ISysAttachmentService iSysAttachmentService;
        */
    //elecl表获取的信息
    List<DwdbPhotoInputPojo> list = new ArrayList<DwdbPhotoInputPojo>();
    //数据表实体list
    //List<DwdbBasic> dwdbBasicList=new ArrayList<>();

    /**
     * 如果使用了spring,请使用这个构造方法。每次创建Listener的时候需要把spring管理的类传进来
     */
    public DemoDataListener() {
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data
     * @param context
     */
    @Override
    public void invoke(DwdbPhotoInputPojo data, AnalysisContext context) {
        list.add(data);
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println(JSON.toJSONString(list));
        for (DwdbPhotoInputPojo dwdbPhotoInputPojo : list) {
            DwdbBasic dwdbBasic = new DwdbBasic();
            dwdbBasic.setName(dwdbPhotoInputPojo.getName());
            dwdbBasic.setIdnumber(dwdbPhotoInputPojo.getIdNum());
            dwdbBasic.setPicture(getAttId(dwdbPhotoInputPojo.getPicture()));
            //dwdbBasicMapper.update(dwdbBasic, new UpdateWrapper<DwdbBasic>().eq("name", dwdbBasic.getName()).eq("idnumber", dwdbBasic.getIdnumber()));
            dwdbBasicMapper.updatePicture(dwdbBasic);
            //basicService.updatePicture(dwdbBasic);


        }
    }


    @Test
    public void test(){
        IDwdbBasicService bean = SpringContextUtils.getBean(IDwdbBasicService.class);
        System.out.println(JSON.toJSONString(list));
        String fileName = "C:\\Users\\yxkj\\Desktop\\测试111.xlsx";
        //String fileName=TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        DemoDataListener demoDataListener = new DemoDataListener();
        demoDataListener.setBasicService(basicService);
        demoDataListener.setDwdbBasicMapper(dwdbBasicMapper);
        EasyExcel.read(fileName, DwdbPhotoInputPojo.class, demoDataListener).sheet().doRead();













        /*for (DwdbPhotoInputPojo dwdbPhotoInputPojo : list) {
            DwdbBasic dwdbBasic = new DwdbBasic();
            dwdbBasic.setName(dwdbPhotoInputPojo.getName());
            dwdbBasic.setIdnumber(dwdbPhotoInputPojo.getIdNum());

            dwdbBasic.setPicture("picture");
            dwdbBasicMapper.update(dwdbBasic, new UpdateWrapper<DwdbBasic>().eq("name", dwdbBasic.getName()).eq("idnumber", dwdbBasic.getIdnumber()));
            //basicService.updatePicture(dwdbBasic);
        }*/
    }


    //附件配套----------------------------------------------------------------------------------------------------------------------------------------
    //附件上传返回attachmentID
    public String getAttId(MultipartFile file){
        String uuid= null;
        SysAttachment sysAttachment = null;
        String fileName = file.getOriginalFilename(),
                file_path,
                type;
        /*if (fileName.contains(",")) {
            return HttpResult.errorMsg("上传附件文件名不允许包含英文逗号,请检查!");
        }*/
        type = FilenameUtils.getExtension(file.getOriginalFilename());
        try {
            //获取附件上传的地址
            file_path = commonUpload(file);
            if (StringUtils.isEmpty(file_path)) {
                return "上传失败";
            }
            try {
                sysAttachment = new SysAttachment(TableUtils.getNewID("Sys_Attachment"), null, null, AesEncryptUtil.encrypt(file_path.substring(file_path.lastIndexOf("/"))),
                        fileName, (int) file.getSize(), type, new Date(), uuid, false, file_path.substring(0, file_path.lastIndexOf("/")));
                sysAttachmentMapper.insert(sysAttachment);
            } catch (Exception e) {
                e.printStackTrace();
            }

        } catch (Exception e) {
            e.printStackTrace();
            return "上传失败";
        }
        return  sysAttachment.getAttachmentID();
    }

    //附件上传
    public String commonUpload(MultipartFile file) throws Exception {

        String real_path;
        String fload_path = getFloadPath();
        File fload = new File(StaticParameters.INPUT_FOLDER + fload_path);
        if (!fload.exists()) {
            if (!fload.mkdirs()) {
                return null;
            }
        }
        FileOutputStream fileOutputStream = null;
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String b = FilenameUtils.getExtension(fileName);
            String n = System.currentTimeMillis() + "" + (int) (Math.random() * 10000); // 避免文件同名
            real_path = StaticParameters.INPUT_FOLDER + fload_path + n + "." + b;
            File files = new File(real_path); // 新建一个文件
            try {
                fileOutputStream = new FileOutputStream(files);
                fileOutputStream.write(file.getBytes());
                fileOutputStream.flush();
                if (isPic(fileName) && files.length() > START_COMPRESS) {
                    Thumbnails.of(files).scale(1f).outputQuality(0.5f).toFile(files);
                }
            } finally {
                if (fileOutputStream != null) {
                    try {
                        fileOutputStream.close();
                    } catch (IOException ie) {
                        ie.printStackTrace();
                    }
                }
            }
            return fload_path + n + "." + b;
        }
        return null;
    }



    private boolean isPic(String file_type) {
        String types = "jpeg|gif|jpg|png|bmp|pic";
        return types.contains(file_type);
    }

















   /* @Test
    //@GetMapping("/readExcelShouldSuccess")
    public void readExcelShouldSuccess() {
        IDwdbBasicService bean = SpringContextUtils.getBean(IDwdbBasicService.class);
        String fileName = "C:\\Users\\yxkj\\Desktop\\12.xlsx";
        //String fileName=TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        EasyExcel.read(fileName, DwdbPhotoInputPojo.class, new DemoDataListener()).sheet().doRead();
        //System.out.println(JSON.toJSONString(list));
    }
*/

}








