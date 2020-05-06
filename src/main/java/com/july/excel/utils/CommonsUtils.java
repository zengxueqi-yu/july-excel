package com.july.excel.utils;

import com.july.excel.constant.ExcelGlobalConstants;
import com.july.excel.entity.ExcelData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
import static org.apache.poi.ss.util.CellUtil.createCell;

/**
 * 集合操作工具
 * @author zengxueqi
 * @program july-excel
 * @since 2020-05-06 17:42
 **/
public class CommonsUtils {

    private static Logger log = LoggerFactory.getLogger(CommonsUtils.class);

    /**
     * 设置数据：有样式（行、列、单元格样式）
     * @param wb
     * @param sxssfRow
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataList(SXSSFWorkbook wb, SXSSFRow sxssfRow, ExcelData excelData) throws Exception {
        if (CollectionUtils.isEmpty(excelData.getDataLists())) {
            log.debug("===> Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (excelData.getSheetName() == null) {
            log.debug("===> Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        for (List<String[]> listRow : excelData.getDataLists()) {
            SXSSFSheet sxssfSheet = wb.createSheet();
            sxssfSheet.setDefaultColumnWidth(excelData.getDefaultColumnWidth());
            wb.setSheetName(k, excelData.getSheetName()[k]);
            CellStyle cellStyle = wb.createCellStyle();
            XSSFFont font = (XSSFFont) wb.createFont();
            SXSSFDrawing sxssfDrawing = sxssfSheet.createDrawingPatriarch();

            int jRow = 0;

            //自定义：大标题（看该方法说明）。
            jRow = setLabelName(jRow, k, wb, excelData.getLabelName(), sxssfRow, sxssfSheet, listRow);

            //自定义：每个表格固定表头（看该方法说明）。
            Integer pane = 1;
            if (!CollectionUtils.isEmpty(excelData.getPaneMap()) && excelData.getPaneMap().get(k + 1) != null) {
                pane = (Integer) excelData.getPaneMap().get(k + 1) + (excelData.getLabelName() != null ? 1 : 0);
                createFreezePane(sxssfSheet, pane);
            }
            //自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getRegionMap())) {
                setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) excelData.getRegionMap().get(k + 1));
            }
            //自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getDropDownMap())) {
                setDataValidation(sxssfSheet, (List<String[]>) excelData.getDropDownMap().get(k + 1), listRow.size());
            }
            //自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getMapColumnWidth())) {
                setColumnWidth(sxssfSheet, (HashMap) excelData.getMapColumnWidth().get(k + 1));
            }
            //默认样式。
            setStyle(cellStyle, font, excelData.getFontSize());

            CellStyle cell_style = null;
            CellStyle row_style = null;
            CellStyle column_style = null;
            //写入小标题与数据。
            Integer SIZE = listRow.size() < ExcelGlobalConstants.MAX_ROWSUM ? listRow.size() : ExcelGlobalConstants.MAX_ROWSUM;
            Integer MAXSYTLE = listRow.size() < ExcelGlobalConstants.MAX_ROWSTYLE ? listRow.size() : ExcelGlobalConstants.MAX_ROWSTYLE;
            for (int i = 0; i < SIZE; i++) {
                sxssfRow = sxssfSheet.createRow(jRow);
                for (int j = 0; j < listRow.get(i).length; j++) {
                    //样式过多会导致GC内存溢出！
                    try {
                        Cell cell = null;
                        if (patternIsImg(listRow.get(i)[j])) {
                            cell = createCell(sxssfRow, j, " ");
                            drawPicture(wb, sxssfDrawing, listRow.get(i)[j], j, jRow);
                        } else {
                            cell = createCell(sxssfRow, j, listRow.get(i)[j]);
                        }
                        cell.setCellStyle(cellStyle);

                        //自定义：每个表格每一列的样式（看该方法说明）。
                        if (excelData.getColumnStyles() != null && jRow >= pane && i <= MAXSYTLE) {
                            setExcelRowStyles(cell, wb, sxssfRow, (List) excelData.getColumnStyles().get(k + 1), j);
                        }
                        //自定义：每个表格每一行的样式（看该方法说明）。
                        if (excelData.getRowStyles() != null && i <= MAXSYTLE) {
                            setExcelRowStyles(cell, wb, sxssfRow, (List) excelData.getRowStyles().get(k + 1), jRow);
                        }
                        //自定义：每一个单元格样式（看该方法说明）。
                        if (excelData.getStyles() != null && i <= MAXSYTLE) {
                            setExcelStyles(cell, wb, sxssfRow, (List<List<Object[]>>) excelData.getStyles().get(k + 1), j, i);
                        }
                    } catch (Exception e) {
                        log.debug("===> Exception Message：The maximum number of cell styles was exceeded. You can define up to 4000 styles!");
                    }
                }
                jRow++;
            }
            k++;
        }
    }

    /**
     * 设置excel样式
     * @param cell         Cell对象
     * @param wb           SXSSFWorkbook对象
     * @param sxssfRow
     * @param fontSize     字体大小
     * @param bold         是否加粗
     * @param center       是否左右上下居中
     * @param isBorder     是否忽略边框
     * @param leftBoolean  左对齐
     * @param rightBoolean 右对齐
     * @param fontColor    字体颜色
     * @param height       行高
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, Integer fontSize, Boolean bold, Boolean center, Boolean isBorder, Boolean leftBoolean,
                                      Boolean rightBoolean, Integer fontColor, Integer height) {
        CellStyle cellStyle = cell.getRow().getSheet().getWorkbook().createCellStyle();
        //保证了既可以新建一个CellStyle，又可以不丢失原来的CellStyle 的样式
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        //左右居中、上下居中
        if (center != null && center) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        //右对齐
        if (rightBoolean != null && rightBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        }
        //左对齐
        if (leftBoolean != null && leftBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
        }
        //是否忽略边框
        if (isBorder != null && isBorder) {
            setBorderColor(cellStyle, isBorder);
        }
        //设置单元格字体样式
        XSSFFont font = (XSSFFont) wb.createFont();
        if (bold != null && bold) {
            font.setBold(bold);
        }
        //行高
        if (height != null) {
            sxssfRow.setHeight((short) (height * 2));
        }
        font.setFontName("宋体");
        font.setFontHeight(fontSize == null ? 12 : fontSize);
        cellStyle.setFont(font);
        //点击可查看颜色对应的值： BLACK(8), WHITE(9), RED(10),
        font.setColor(IndexedColors.fromInt(fontColor == null ? 8 : fontColor).index);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置excel行的样式
     * @param cell
     * @param wb
     * @param sxssfRow
     * @param rowstyleList
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelRowStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<Object[]> rowstyleList, int rowIndex) {
        if (rowstyleList != null && rowstyleList.size() > 0) {
            Integer[] rowstyle = (Integer[]) rowstyleList.get(1);
            for (int i = 0; i < rowstyle.length; i++) {
                if (rowIndex == rowstyle[i]) {
                    Boolean[] bool = (Boolean[]) rowstyleList.get(0);
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (rowstyleList.size() >= 3) {
                        int leng = rowstyleList.get(2).length;
                        if (leng >= 1) {
                            fontColor = (Integer) rowstyleList.get(2)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) rowstyleList.get(2)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) rowstyleList.get(2)[2];
                        }
                    }
                    setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                }
            }
        }
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
    private static void drawPicture(SXSSFWorkbook wb, SXSSFDrawing sxssfDrawing, String pictureUrl, int colIndex, int rowIndex) {
        //rowIndex代表当前行
        try {
            if (pictureUrl != null) {
                URL url = new URL(pictureUrl);
                //打开链接
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");
                conn.setConnectTimeout(5 * 1000);
                InputStream inStream = conn.getInputStream();
                byte[] data = readInputStream(inStream);
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
     * 流操作
     * @param inStream
     * @return byte[]
     * @author zengxueqi
     * @since 2020/5/6
     */
    private static byte[] readInputStream(InputStream inStream) throws Exception {
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        //创建一个Buffer字符串
        byte[] buffer = new byte[1024];
        //每次读取的字符串长度，如果为-1，代表全部读取完毕
        int len = 0;
        //使用一个输入流从buffer里把数据读取出来
        while ((len = inStream.read(buffer)) != -1) {
            //用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度
            outStream.write(buffer, 0, len);
        }
        //关闭输入流
        inStream.close();
        //把outStream里的数据写入内存
        return outStream.toByteArray();
    }

    /**
     * 自定义：大标题
     * @param jRow
     * @param k
     * @param wb
     * @param labelName
     * @param sxssfRow
     * @param sxssfSheet
     * @param listRow
     * @return int
     * @author zengxueqi
     * @since 2020/5/6
     */
    private static int setLabelName(Integer jRow, Integer k, SXSSFWorkbook wb, String[] labelName, SXSSFRow sxssfRow, SXSSFSheet sxssfSheet, List<String[]> listRow) {
        if (labelName != null) {
            //自定义：大标题和样式。参数说明：new String[]{"表格数据一", "表格数据二", "表格数据三"}
            sxssfRow = sxssfSheet.createRow(0);
            Cell cell = createCell(sxssfRow, 0, labelName[k]);
            setMergedRegion(sxssfSheet, 0, 0, 0, listRow.get(0).length - 1);
            setLabelStyles(wb, cell, sxssfRow);
            jRow = 1;
        }
        return jRow;
    }

    /**
     * 设置数据：无样式（行、列、单元格样式）
     * @param wb
     * @param sxssfRow
     * @param excelData
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataListNoStyle(SXSSFWorkbook wb, SXSSFRow sxssfRow, ExcelData excelData) throws Exception {
        if (CollectionUtils.isEmpty(excelData.getDataLists())) {
            log.debug("===> Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (excelData.getSheetName() == null) {
            log.debug("===> Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        for (List<String[]> listRow : excelData.getDataLists()) {
            SXSSFSheet sxssfSheet = wb.createSheet();
            sxssfSheet.setDefaultColumnWidth(excelData.getDefaultColumnWidth());
            wb.setSheetName(k, excelData.getSheetName()[k]);
            CellStyle cellStyle = wb.createCellStyle();
            XSSFFont font = (XSSFFont) wb.createFont();

            int jRow = 0;
            //自定义：大标题（看该方法说明）。
            jRow = setLabelName(jRow, k, wb, excelData.getLabelName(), sxssfRow, sxssfSheet, listRow);

            //自定义：每个表格固定表头（看该方法说明）。
            Integer pane = 1;
            if (!CollectionUtils.isEmpty(excelData.getPaneMap()) && excelData.getPaneMap().get(k + 1) != null) {
                pane = (Integer) excelData.getPaneMap().get(k + 1) + (excelData.getLabelName() != null ? 1 : 0);
                createFreezePane(sxssfSheet, pane);
            }
            //自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getRegionMap())) {
                setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) excelData.getRegionMap().get(k + 1));
            }
            //自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getDropDownMap())) {
                setDataValidation(sxssfSheet, (List<String[]>) excelData.getDropDownMap().get(k + 1), listRow.size());
            }
            //自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
            if (!CollectionUtils.isEmpty(excelData.getMapColumnWidth())) {
                setColumnWidth(sxssfSheet, (HashMap) excelData.getMapColumnWidth().get(k + 1));
            }
            //默认样式。
            setStyle(cellStyle, font, excelData.getFontSize());

            //写入小标题与数据。
            Integer SIZE = listRow.size() < ExcelGlobalConstants.MAX_ROWSUM ? listRow.size() : ExcelGlobalConstants.MAX_ROWSUM;
            for (int i = 0; i < SIZE; i++) {
                sxssfRow = sxssfSheet.createRow(jRow);
                for (int j = 0; j < listRow.get(i).length; j++) {
                    Cell cell = createCell(sxssfRow, j, listRow.get(i)[j]);
                    cell.setCellStyle(cellStyle);
                }
                jRow++;
            }
            k++;
        }
    }

    /**
     * 写数据与流关闭
     * @param sxssfWorkbook
     * @param outputStream
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void writeAndColse(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream) throws Exception {
        try {
            if (outputStream != null) {
                sxssfWorkbook.write(outputStream);
                sxssfWorkbook.dispose();
                outputStream.flush();
                outputStream.close();
            }
        } catch (Exception e) {
            System.out.println(" Andyczy ExcelUtils Exception Message：Output stream is not empty !");
            e.getSuppressed();
        }
    }

    /**
     * 所有自定义单元格样式
     * @param cell
     * @param wb
     * @param sxssfRow
     * @param styles
     * @param cellIndex
     * @param rowIndex
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<Object[]>> styles, int cellIndex, int rowIndex) {
        if (styles != null) {
            for (int z = 0; z < styles.size(); z++) {
                List<Object[]> stylesList = styles.get(z);
                if (stylesList != null) {
                    //样式
                    Boolean[] bool = (Boolean[]) stylesList.get(0);
                    //颜色和字体
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (stylesList.size() >= 2) {
                        int leng = stylesList.get(1).length;
                        if (leng >= 1) {
                            fontColor = (Integer) stylesList.get(1)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) stylesList.get(1)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) stylesList.get(1)[2];
                        }
                    }
                    //第几行第几列
                    for (int m = 2; m < stylesList.size(); m++) {
                        Integer[] str = (Integer[]) stylesList.get(m);
                        if (cellIndex + 1 == (str[1]) && rowIndex + 1 == (str[0])) {
                            setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                        }
                    }
                }
            }
        }
    }

    /**
     * 大标题样式
     * @param wb
     * @param cell
     * @param sxssfRow
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setLabelStyles(SXSSFWorkbook wb, Cell cell, SXSSFRow sxssfRow) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sxssfRow.setHeight((short) (399 * 2));
        XSSFFont font = (XSSFFont) wb.createFont();
        font.setFontName("宋体");
        font.setFontHeight(16);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置默认样式
     * @param cellStyle
     * @param font
     * @param fontSize
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setStyle(CellStyle cellStyle, XSSFFont font, Integer fontSize) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName("宋体");
        cellStyle.setFont(font);
        font.setFontHeight(fontSize);
        setBorder(cellStyle, true);
    }

    /**
     * 锁定行（固定表头）
     * @param sxssfSheet
     * @param row
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void createFreezePane(SXSSFSheet sxssfSheet, Integer row) {
        if (row != null && row > 0) {
            sxssfSheet.createFreezePane(0, row, 0, 1);
        }
    }

    /**
     * 自定义列宽
     * @param sxssfSheet
     * @param map
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setColumnWidth(SXSSFSheet sxssfSheet, HashMap map) {
        if (map != null) {
            Iterator iterator = map.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry entry = (Map.Entry) iterator.next();
                Object key = entry.getKey();
                Object val = entry.getValue();
                sxssfSheet.setColumnWidth((int) key, (int) val * 512);
            }
        }
    }

    /**
     * excel 合并单元格
     * @param sheet
     * @param rowColList
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setMergedRegion(SXSSFSheet sheet, ArrayList<Integer[]> rowColList) {
        if (rowColList != null && rowColList.size() > 0) {
            for (int i = 0; i < rowColList.size(); i++) {
                Integer[] str = rowColList.get(i);
                if (str.length > 0 && str.length == 4) {
                    Integer firstRow = str[0];
                    Integer lastRow = str[1];
                    Integer firstCol = str[2];
                    Integer lastCol = str[3];
                    setMergedRegion(sheet, firstRow, lastRow, firstCol, lastCol);
                }
            }
        }
    }

    /**
     * 合并单元格
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setMergedRegion(SXSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * 下拉列表
     * @param sheet
     * @param dropDownListData
     * @param dataListSize
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataValidation(SXSSFSheet sheet, List<String[]> dropDownListData, int dataListSize) {
        if (dropDownListData.size() > 0) {
            for (int col = 0; col < dropDownListData.get(0).length; col++) {
                Integer colv = Integer.parseInt(dropDownListData.get(0)[col]);
                setDataValidation(sheet, dropDownListData.get(col + 1), 1, dataListSize < 100 ? 500 : dataListSize, colv, colv);
            }
        }
    }

    /**
     * 下拉列表
     * @param xssfWsheet
     * @param list
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setDataValidation(SXSSFSheet xssfWsheet, String[] list, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
        DataValidationHelper helper = xssfWsheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(list);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.createErrorBox(ExcelGlobalConstants.DataValidationError1, ExcelGlobalConstants.DataValidationError2);
        //处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        xssfWsheet.addValidationData(dataValidation);
    }

    /**
     * 设置边框
     * @param cellStyle
     * @param isBorder
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setBorder(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
        } else {
            //添加白色背景，统一设置边框后但不能选择性去掉，只能通过背景覆盖达到效果。
            cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
        }
    }

    /**
     * 设置边框颜色
     * @param cellStyle
     * @param isBorder
     * @return void
     * @author zengxueqi
     * @since 2020/5/6
     */
    public static void setBorderColor(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
            //添加白色背景，统一设置边框后但不能选择性去掉，只能通过背景覆盖达到效果。
            cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.WHITE.getIndex());
            cellStyle.setTopBorderColor(IndexedColors.WHITE.getIndex());
        }
    }

}
