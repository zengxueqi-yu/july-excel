package com.july.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;

/**
 * 图片操作类
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-07 16:34
 **/
public class ImageUtils {

    /**
     * 是否是图片
     * @param str
     * @return java.lang.Boolean
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static Boolean patternIsImg(String str) {
        if (StringUtils.isEmpty(str)) {
            return false;
        }
        String reg = ".+(.JPEG|.jpeg|.JPG|.jpg|.PNG|.png|.gif)$";
        Pattern pattern = Pattern.compile(reg);
        Matcher matcher = pattern.matcher(str);
        Boolean isPicture = matcher.find();
        return isPicture;
    }

    /**
     * 画图片
     * @param sxssfWorkbook
     * @param sxssfDrawing
     * @param pictureUrl
     * @param colIndex
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void drawPicture(SXSSFWorkbook sxssfWorkbook, SXSSFDrawing sxssfDrawing, String pictureUrl, int colIndex, int rowIndex) {
        //rowIndex代表当前行
        try {
            if (pictureUrl != null) {
                URL url = new URL(pictureUrl);
                //打开链接
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");
                conn.setConnectTimeout(5 * 1000);
                InputStream inStream = conn.getInputStream();
                byte[] data = ExcelUtils.readInputStream(inStream);
                //设置图片大小，
                XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 50, 50, colIndex, rowIndex, colIndex + 1, rowIndex + 1);
                anchor.setAnchorType(DONT_MOVE_AND_RESIZE);
                sxssfDrawing.createPicture(anchor, sxssfWorkbook.addPicture(data, XSSFWorkbook.PICTURE_TYPE_JPEG));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取图片和位置 (xls)
     * @param sheet
     * @return java.util.Map<java.lang.String, org.apache.poi.ss.usermodel.PictureData>
     * @author zengxueqi
     * @since 2020/5/9
     */
    public static Map<String, PictureData> getPictures1(HSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<>();
        if (sheet.getDrawingPatriarch() != null) {
            List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
            for (HSSFShape shape : list) {
                if (shape instanceof HSSFPicture) {
                    HSSFPicture picture = (HSSFPicture) shape;
                    HSSFClientAnchor cAnchor = (HSSFClientAnchor) picture.getAnchor();
                    PictureData pdata = picture.getPictureData();
                    /**行号-列号**/
                    String key = cAnchor.getRow1() + "-" + cAnchor.getCol1();
                    map.put(key, pdata);
                }
            }
        }
        return map;
    }

    /**
     * 获取图片和位置 (xlsx)
     * @param sheet
     * @return java.util.Map<java.lang.String, org.apache.poi.ss.usermodel.PictureData>
     * @author zengxueqi
     * @since 2020/5/9
     */
    public static Map<String, PictureData> getPictures2(XSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<POIXMLDocumentPart> list = sheet.getRelations();
        for (POIXMLDocumentPart part : list) {
            if (part instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) part;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker marker = anchor.getFrom();
                    String key = marker.getRow() + "-" + marker.getCol();
                    map.put(key, picture.getPictureData());
                }
            }
        }
        return map;
    }

}
