package com.yxkj.dwdb_common.utils;

import com.yxkj.common.core.config.SpringContextUtils;
import com.yxkj.common.core.service.sys.ISysAttachmentService1;
import com.yxkj.common.core.util.other.StringUtils;
import org.apache.http.entity.ContentType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.*;

public class ExcelUtil {

    /**
     * 获取Excel2003图片
     *
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     * @throws IOException
     */
    public static Map<String, PictureData> getSheetPictrues(HSSFSheet sheet, HSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.size() != 0) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = "row_" + String.valueOf(anchor.getRow1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    /**
     * 获取Excel2007图片
     *
     * @param sheet    当前sheet对象
     * @param workbook 工作簿对象
     * @return Map key:图片单元格索引（0_1_1）String，value:图片流PictureData
     */
    public static Map<String, PictureData> getSheetPictrues07(XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = "row_" + ctMarker.getRow();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }

        return sheetIndexPicMap;
    }

    public static String printImg(Map<String, PictureData> map, String row) throws IOException {
        ISysAttachmentService1 sysAttachmentService = SpringContextUtils.getBean(ISysAttachmentService1.class);
        Date date = new Date();
        String name = date.getTime() + "";
        // 获取图片流
        PictureData pic = map.get("row_" + row);
        // 获取图片格式
        String ext = "." + pic.suggestFileExtension();
        // 获取图片索引
        String picName = name + ext;

        byte[] data = new byte[0];

        try {
            data = compressImage(pic.getData(), 125);
            InputStream inputStream = new ByteArrayInputStream(data);
            MultipartFile file = new MockMultipartFile(ContentType.APPLICATION_OCTET_STREAM.toString(), picName, ext, inputStream);

            String imgId = sysAttachmentService.uploadForWindows(file, null);
            if ("错误".equals(imgId)) {
                return null;
            }
            return imgId;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 图片压缩
     *
     * @param imageByte
     * @param ppi
     * @return
     */
    public static byte[] compressImage(byte[] imageByte, int ppi) {
        byte[] smallImage = null;
        int width = 0, height = 0;

        if (imageByte == null) {
            return null;
        }

        ByteArrayInputStream byteInput = new ByteArrayInputStream(imageByte);
        try {
            Image image = ImageIO.read(byteInput);
            int w = image.getWidth(null);
            int h = image.getHeight(null);
            // 调整宽度和高度以避免图像失真
            double scale = 0;
            scale = Math.min((float) ppi / w, (float) ppi / h);
            width = (int) (w * scale);
            width -= width % 4;
            height = (int) (h * scale);

            if (scale >= (double) 1)
                return imageByte;

            BufferedImage buffImg = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            buffImg.getGraphics().drawImage(image.getScaledInstance(width, height, Image.SCALE_SMOOTH), 0, 0, null);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ImageIO.write(buffImg, "png", out);
            smallImage = out.toByteArray();
            return smallImage;

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /*记录错误数据*/
    public static Map<String, Object> errorPerson(Map<String, Object> map, String msg) {
        map.put("errorMsg", msg);
        try {
            if (map.containsKey("gender") && StringUtils.isNotEmpty(map.get("gender").toString())) {
                if (Integer.parseInt(map.get("gender").toString()) == 1) {
                    map.put("gender", "女");
                } else {
                    map.put("gender", "男");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            if (map.containsKey("negative_list") && StringUtils.isNotEmpty(map.get("negative_list").toString())) {
                if (Integer.parseInt(map.get("negativeList").toString()) == 1) {
                    map.put("negativeList", "是");
                } else {
                    map.put("negativeList", "否");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    /**
     * 数据处理
     *
     * @param map
     * @param basicMap
     * @return
     */
    public static Map<String, Object> dealWith(Map<String, List<Map<String, Object>>> map, Map<String, Map<String, Object>> basicMap, Integer type) {
        //区分表错误集合
        Map<String, List<Map<String, Object>>> errorMap = new HashMap<>();
        //正确集合    表,标识,数据
        Map<String, Map<String, List<Map<String, Object>>>> correctMap = new HashMap<>();

        for (Map.Entry<String, List<Map<String, Object>>> entry : map.entrySet()) {
            if (entry.getKey().contains("基本情况表")) {
                continue;
            }
            List<Map<String, Object>> dataList = entry.getValue();
            List<Map<String, Object>> errorList = new ArrayList<>();    //错误数据集合
            Map<String, List<Map<String, Object>>> d = new HashMap<>();
            correctMap.put(entry.getKey(), d);
            for (Map<String, Object> dataMap : dataList) {
                boolean flag = false;
                String biaoshi = "";
                if (type != 9) {
                    if (dataMap.containsKey("id_number")) {
                        biaoshi = dataMap.get("id_number").toString();
                        if (basicMap.containsKey(biaoshi)) {
                            flag = true;
                        }
                    }
                    if (!flag) {
                        if (dataMap.containsKey("name") && dataMap.containsKey("year")) {

                            biaoshi = dataMap.get("name").toString().replaceAll(" ", "") + dataMap.get("year").toString();
                            if (basicMap.containsKey(biaoshi)) {
                                flag = true;
                            }
                        }
                    }
                } else {
                    if (dataMap.containsKey("name") && dataMap.get("name") != null) {
                        biaoshi = dataMap.get("name").toString().replaceAll(" ", "");
                        if (basicMap.containsKey(biaoshi)) {
                            flag = true;
                        }
                    }
                }


                if (flag) {
                    if (correctMap.get(entry.getKey()).containsKey(biaoshi)) {
                        correctMap.get(entry.getKey()).get(biaoshi).add(dataMap);
                    } else {
                        List<Map<String, Object>> list = new ArrayList<>();
                        list.add(dataMap);
                        correctMap.get(entry.getKey()).put(biaoshi, list);
                    }
                } else {
                    dataMap.put("errorMsg", "无法在主表中找到匹配信息");
                    errorList.add(dataMap);
                }
            }
            errorMap.put(entry.getKey(), errorList);
        }

        Map<String, Object> retMap = new HashMap<>();
        retMap.put("errorMap", errorMap);
        retMap.put("correctList", correctMap);

        return retMap;
    }

    public static Cell changeVal(Cell cell, String sqlName) {
        if ("gender".equals(sqlName)) {
            if ("男".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            } else if ("女".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            }
        } else if ("negativelist".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("middle_and_upper_class".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("middle_andupper_class".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("key_areas_of_division_of_labor".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("leader".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("downsizing_sign".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("private_entrepreneur".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("corporate_executives".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("phD".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("is_phD".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("changjiang_scholar".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("communist_party".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("overseasrelations".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        } else if ("right_of_permanent_residence_abroad".equals(sqlName)) {
            if ("否".equals(cell.getStringCellValue())) {
                cell.setCellValue("1");
            } else if ("是".equals(cell.getStringCellValue())) {
                cell.setCellValue("0");
            }
        }
        return cell;
    }

    public static Map<String, Object> changeMapVal(Map<String, Object> map) {
        if (map.containsKey("gender") && map.get("gender") != null) {
            if ("0".equals(map.get("gender").toString())) {
                map.put("gender", "男");
            } else if ("1".equals(map.get("gender").toString())) {
                map.put("gender", "女");
            }
        }
        if (map.containsKey("negativelist") && map.get("negativelist") != null) {
            if ("0".equals(map.get("negativelist").toString())) {
                map.put("negative_list", "是");
            } else if ("1".equals(map.get("negativelist").toString())) {
                map.put("negative_list", "否");
            }
        }
        if (map.containsKey("middle_and_upper_class") && map.get("middle_and_upper_class") != null) {
            if ("0".equals(map.get("middle_and_upper_class").toString())) {
                map.put("middle_and_upper_class", "是");
            } else if ("1".equals(map.get("middle_and_upper_class").toString())) {
                map.put("middle_and_upper_class", "否");
            }
        }
        if (map.containsKey("middle_andupper_class") && map.get("middle_andupper_class") != null) {
            if ("0".equals(map.get("middle_andupper_class").toString())) {
                map.put("middle_andupper_class", "是");
            } else if ("1".equals(map.get("middle_andupper_class").toString())) {
                map.put("middle_andupper_class", "否");
            }
        }
        if (map.containsKey("key_areas_of_division_of_labor") && map.get("key_areas_of_division_of_labor") != null) {
            if ("0".equals(map.get("key_areas_of_division_of_labor").toString())) {
                map.put("key_areas_of_division_of_labor", "是");
            } else if ("1".equals(map.get("key_areas_of_division_of_labor").toString())) {
                map.put("key_areas_of_division_of_labor", "否");
            }
        }
        if (map.containsKey("leader") && map.get("leader") != null) {
            if ("0".equals(map.get("leader").toString())) {
                map.put("leader", "是");
            } else if ("1".equals(map.get("leader").toString())) {
                map.put("leader", "否");
            }
        }
        if (map.containsKey("downsizing_sign") && map.get("downsizing_sign") != null) {
            if ("0".equals(map.get("downsizing_sign").toString())) {
                map.put("downsizing_sign", "是");
            } else if ("1".equals(map.get("downsizing_sign").toString())) {
                map.put("downsizing_sign", "否");
            }
        }
        if (map.containsKey("private_entrepreneur") && map.get("private_entrepreneur") != null) {
            if ("0".equals(map.get("private_entrepreneur").toString())) {
                map.put("private_entrepreneur", "是");
            } else if ("1".equals(map.get("private_entrepreneur").toString())) {
                map.put("private_entrepreneur", "否");
            }
        }
        if (map.containsKey("corporate_executives") && map.get("corporate_executives") != null) {
            if ("0".equals(map.get("corporate_executives").toString())) {
                map.put("corporate_executives", "是");
            } else if ("1".equals(map.get("corporate_executives").toString())) {
                map.put("corporate_executives", "否");
            }
        }
        if (map.containsKey("phD") && map.get("phD") != null) {
            if ("0".equals(map.get("phD").toString())) {
                map.put("phD", "是");
            } else if ("1".equals(map.get("phD").toString())) {
                map.put("phD", "否");
            }
        }
        if (map.containsKey("is_phD") && map.get("is_phD") != null) {
            if ("0".equals(map.get("is_phD").toString())) {
                map.put("is_phD", "是");
            } else if ("1".equals(map.get("is_phD").toString())) {
                map.put("is_phD", "否");
            }
        }
        if (map.containsKey("changjiang_scholar") && map.get("changjiang_scholar") != null) {
            if ("0".equals(map.get("changjiang_scholar").toString())) {
                map.put("changjiang_scholar", "是");
            } else if ("1".equals(map.get("changjiang_scholar").toString())) {
                map.put("changjiang_scholar", "否");
            }
        }
        if (map.containsKey("communist_party") && map.get("communist_party") != null) {
            if ("0".equals(map.get("communist_party").toString())) {
                map.put("communist_party", "是");
            } else if ("1".equals(map.get("communist_party").toString())) {
                map.put("communist_party", "否");
            }
        }
        if (map.containsKey("overseasrelations") && map.get("overseasrelations") != null) {
            if ("0".equals(map.get("overseasrelations").toString())) {
                map.put("overseasrelations", "是");
            } else if ("1".equals(map.get("overseasrelations").toString())) {
                map.put("overseasrelations", "否");
            }
        }
        if (map.containsKey("right_of_permanent_residence_abroad") && map.get("right_of_permanent_residence_abroad") != null) {
            if ("0".equals(map.get("right_of_permanent_residence_abroad").toString())) {
                map.put("right_of_permanent_residence_abroad", "是");
            } else if ("1".equals(map.get("right_of_permanent_residence_abroad").toString())) {
                map.put("right_of_permanent_residence_abroad", "否");
            }
        }
        return map;
    }

    /**
     * 判断对象中属性值是否全为空
     *
     * @param object
     * @return
     */
    public static boolean checkObjAllFieldsIsNull(Object object) {
        if (null == object) {
            return true;
        }

        try {
            Class<?> clazz = object.getClass();
            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                // 获取属性的名字
                String name = fields[i].getName();
                if ("rwid".equals(name) || "serialVersionUID".equals(name)) {
                    continue;
                }
                // 将属性的首字符大写，方便构造get，set方法
                name = name.substring(0, 1).toUpperCase() + name.substring(1);
                Method m = object.getClass().getMethod("get" + name);
                // 调用getter方法获取属性值
                Object value = m.invoke(object);

                fields[i].setAccessible(true);

                if (value != null && String.valueOf(value) != null && !"".equals(String.valueOf(value).trim())) {
                    return false;
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return true;
    }

    public static boolean checkObjFieldIsNull(Object obj, Boolean b, String... name) throws IllegalAccessException {
        boolean flag = false;
        List<String> list = java.util.Arrays.asList(name);
        for (Field f : obj.getClass().getDeclaredFields()) {
            f.setAccessible(true);
            if (b) {
                if (!list.contains(f.getName())) {
                    if (f.get(obj) == null) {
                        flag = true;
                        return flag;
                    }
                }
            } else {
                if (list.contains(f.getName())) {
                    if (f.get(obj) == null) {
                        flag = true;
                        return flag;
                    }
                }
            }
        }
        return flag;
    }

    public static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BOOLEAN) {
                return false;
            }
        }
        return true;
    }

    public static Object changeDate(String value) {
        value = value.replace('.', '-');
        value = value.replace('年', '-');
        value = value.replace('月', '-');
        value = value.replace('/', '-');
        value = value.replaceAll(" ", "");
        value = value.replaceAll("\n", "");
        value = value.replaceAll("\t", "");
        String result = "";
        for (int i = 0; i < value.length(); i++) {
            if ('.' == value.charAt(i) || '-' == value.charAt(i)) {
                continue;
            }
            if (result.length() == 4 || result.length() == 7) {
                result += "-";
            }
            result += value.charAt(i);
        }

        if (result.length() == 4) {
            result += "-01";
        }
        return result;
    }

    public static boolean isNumber(String val) {
        if (val == null || "".equals(val)) {
            return false;
        }
        for (int i = 0; i < val.length(); i++) {
            if (val.charAt(i) == '.') {
                continue;
            }
            if (val.charAt(i) <= '0' || val.charAt(i) >= '9') {
                return false;
            }
        }
        return true;
    }
}
