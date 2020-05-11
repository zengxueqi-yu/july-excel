###介绍
- July Excel一款开源的Excel操纵工具，支持Excel的导入与导出，后期支持Excel深度定制操作。
###使用方式：
#####第一步：在pom.xml中引入July Doc依赖即可
```
<dependency>
    <groupId>com.github.zengxueqi-yu</groupId>
    <artifactId>july-excel</artifactId>
    <version>0.0.1-RELEASE</version>
</dependency>
```
#####第二步：控制器使用Excel工具类实现Excel操作
```
/**
 * excel测试控制器
 * @author zengxueqi
 * @program july-bpm
 * @since 2020-05-07 16:57
 **/
@RestController
@RequestMapping("/excel")
@Slf4j
public class ExcelController {
    /**
     * 导出excel数据
     * @param response
     * @return void
     * @author zengxueqi
     * @since 2020/5/9
     */
    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletResponse response) {
        ExcelData excelData = ExcelData.builder()
                //需要导出的excel数据
                .excelData(exportExcel())
                //sheet名称
                .sheetName("测试数据")
                //文件名称
                .fileName("测试数据")
                //字体大小
                .fontSize(12)
                //列宽
                .cellWidth(25)
                //大标题
                .labelName("教师请假信息数据")
                .build();
        long startTime = System.currentTimeMillis();
        com.july.excel.excel.ExcelFactory.exportExcelData(excelData, TeacherLeaveRecordExcel.class, response);
        log.info("===> time:" + (System.currentTimeMillis() - startTime) + " ms!");
    }
    /**
     * 模拟需要导出的excel数据
     * @param
     * @return java.util.List<com.july.bpm.vo.node.EduTeacherLeaveRecordExcel1>
     * @author zengxueqi
     * @since 2020/5/9
     */
    public List<TeacherLeaveRecordExcel> exportExcel() {
        List<TeacherLeaveRecordExcel> list = new ArrayList<>();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        for (int i = 0; i < 10; i++) {
            TeacherLeaveRecordExcel excel = TeacherLeaveRecordExcel.builder()
                    .schoolName("鹭岛小学" + i)
                    .teacherName("曾雪琪" + i)
                    .activeTime(simpleDateFormat.format(new Date()))
                    .img("http://tdog.oss-cn-hangzhou.aliyuncs.com/default/dJWRp2TXihjpeaxH4ANnExzmSG77GTcF.jpg")
                    .build();
            list.add(excel);
        }
        return list;
    }
    /**
     * 导入excel数据
     * @param file
     * @return void
     * @author zengxueqi
     * @since 2020/5/9
     */
    @GetMapping("/importExcel")
    public void importExcel(@RequestBody MultipartFile file) throws Exception {
        ExcelReadData excelTitleRowNum = ExcelReadData.builder()
                .sheetNum(0)
                .rowNum(1)
                .build();
        ExcelReadData excelReadData = ExcelReadData.builder()
                .sheetNum(0)
                .rowNum(2)
                .build();
        ExcelData excelData = ExcelData.builder()
                //导入测试excel这个sheet
                .sheetName("测试数据")
                //时间格式化
                .expectDateFormatStr("yyyy-MM-dd")
                //sheet标题所在行数(读取sheet下标为0，行数为1的标题)
                .excelTitleRowNum(Arrays.asList(excelTitleRowNum))
                //sheet解析数据开始行数(从sheet下标为0，行数为2开始解析数据)
                .excelReadDataList(Arrays.asList(excelReadData))
                .build();
        List<TeacherLeaveRecordExcel> excel1s = ExcelFactory.importExcelData(file, TeacherLeaveRecordExcel.class, excelData);
        log.info("excel data===>" + JSON.toJSONString(excel1s));
    }
}
```
#####如下图：
- 1.导出的excel数据
- 2.导入excel解析的数据
```
[{
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": "",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学0",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪0"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学1",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪1"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学2",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪2"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学3",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪3"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学4",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪4"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学5",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪5"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学6",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪6"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学7",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪7"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学8",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪8"
}, {
    "activeTime": "2020-05-09",
    "createdTime": "",
    "endTime": "",
    "img": " ",
    "leaveTypeName": "",
    "pid": 0,
    "schoolName": "鹭岛小学9",
    "startTime": "",
    "status": "",
    "teacherName": "曾雪琪9"
}]
```
