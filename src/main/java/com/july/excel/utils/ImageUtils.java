package com.july.excel.utils;

import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
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
        String reg = ".+(.JPEG|.jpeg|.JPG|.jpg|.png|.gif)$";
        Pattern pattern = Pattern.compile(reg);
        Matcher matcher = pattern.matcher(str);
        Boolean temp = matcher.find();
        return temp;
    }

    /**
     * 画图片
     * @param wb
     * @param sxssfDrawing
     * @param pictureUrl
     * @param colIndex
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void drawPicture(SXSSFWorkbook wb, SXSSFDrawing sxssfDrawing, String pictureUrl, int colIndex, int rowIndex) {
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
                sxssfDrawing.createPicture(anchor, wb.addPicture(data, XSSFWorkbook.PICTURE_TYPE_JPEG));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
