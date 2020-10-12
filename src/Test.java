/**
 * github：https://github.com/lieyanfeimao/PoiExcelExport  
 * 码云：https://gitee.com/edadmin/PoiExcelExport
 */
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xuanyimao.poiexcelexporttool.ExcelExportManager;
import com.xuanyimao.poiexcelexporttool.bean.CellField;
import com.xuanyimao.poiexcelexporttool.bean.TempletExcel;
import com.xuanyimao.poiexcelexporttool.common.Constants;
import com.xuanyimao.poiexcelexporttool.common.ExcelUtil;
import com.xuanyimao.poiexcelexporttool.listener.WorkbookListener;
/**
 * 测试类
 * 在本类中 右键 > Run As > Java Application 查看输出结果
 * @author liuming
 */
public class Test {

    public static String json1="[" + 
            "    [" + 
            "        {field:'name', title: '姓名', width:10, rowspan: 2, comment:'这是批注'}," + 
            "        {field:'age', title: '年龄', width:10, rowspan: 2}," + 
            "        {field:'age', title: '性别', width:10, rowspan: 2}," + 
            "        {title: '成绩', width:10, rowspan: 1,colspan:3}" + 
            "    ]," + 
            "    [" + 
            "        {field:'yw', title: '语文', width:10}," + 
            "        {field:'sx', title: '数学', width:10}," + 
            "        {field:'yy', title: '英语', width:10}" + 
            "    ]" + 
            "]";
    public static String json2="[" + 
            "    [" + 
            "        {title: '狗子名单', width:10, rowspan: 5}," + 
            "        {field:'name', title: '姓名', width:10, rowspan: 5, comment:'这是批注'}," + 
            "        {field:'age', title: '年龄上层', width:10, rowspan: 4}," + 
            "        {field:'sex', title: '性别上层', width:10, rowspan: 3}," + 
            "        {title: '成绩', width:10, rowspan: 1,colspan:3}" + 
            "    ]," + 
            "    [" + 
            "        {field:'yw', title: '语文1', width:10}," + 
            "        {field:'sx', title: '数学1', width:10}," + 
            "        {field:'yy', title: '英语1', width:10}" + 
            "    ]," + 
            "    [" + 
            "        {field:'yw', title: '语文2', width:10}," + 
            "        {field:'sx', title: '数学2', width:10}," + 
            "        {field:'yy', title: '英语2', width:10}" + 
            "    ]," + 
            "    [" + 
            "        {title: '性别', width:10, rowspan: 2}," + 
            "        {field:'yw', title: '语文3', width:10}," + 
            "        {field:'sx', title: '数学3', width:10}," + 
            "        {field:'yy', title: '英语4', width:10}" + 
            "    ]," + 
            "    [" + 
            "        {field:'age', title: '年龄', width:10}," + 
            "        {field:'yw', title: '语文', width:10}," + 
            "        {field:'sx', title: '数学', width:10}," + 
            "        {field:'yy', title: '英语', width:10}" + 
            "    ]," + 
            "]";
    
    public static String json3="[" + 
            "    [" + 
            "        {field:'name', title: '姓名', width:10, rowspan: 2, comment:'这是批注'}," + 
            "        {field:'age', title: '年龄', width:10, rowspan: 2}," + 
            "        {field:'age', title: '性别', width:10, rowspan: 2,cellStyle:'test1'}," + 
            "        {title: '成绩', width:10, rowspan: 1,colspan:3,titleStyle:'test1'}" + 
            "    ]," + 
            "    [" + 
            "        {field:'yw', title: '语文', width:10}," + 
            "        {field:'sx', title: '数学', width:10,titleStyle:'test1',cellStyle:'test1'}," + 
            "        {field:'yy', title: '英语', width:10}" + 
            "    ]" + 
            "]";
    
    private static String cellStyleJson="[{'name':'test1',row:0,col:0},{'name':'test2',row:0,col:1},{'name':'test3',row:0,col:2}]";
    /**excel保存目录，路径末尾斜杠可加可不加*/
    private static String excelSaveFolder="D:/exceltest";
    
    public static void main(String[] args) {
        if(StringUtils.isBlank(excelSaveFolder)){
            System.out.println("请先为变量excelSaveFolder赋值,设置一个excel存储路径");
            return;
        }
        //简单的导出演示
//        normalExport();
        //设置导出程序的参数进行微操，做一名微操圣手
//        paramExport();
        //试试五级表头
//        levelFiveExport();
        //试试20级表头
//        level20Export();
        //只创建表头。如果您只想用本程序创建多级表头，请参考这里
//        createSheetTitleOnly();
        //根据模板导出
//        exportByTemplet();
        //使用模板文件设置样式，不保证所有样式都能设置上去
//        createStyleByTemplet();
        //自定义单元格样式。您也可以修改 com.xuanyimao.poiexcelexporttool.common.ExcelUtil 的getDefaultTitleCellStyle和getDefaultDataCellStyle方法设置模式样式
        createStyleByMe();
    }
    /**
     * 简单的导出演示
     * @author:liuming
     */
    public static void normalExport() {
        List<Map<String,Object>> datas=initData1();
        ExcelExportManager em=ExcelExportManager.Builder();
        try {
            String fileName=em.createExcel(excelSaveFolder, json1, datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    /**
     * 设置导出程序的参数进行微操
     * @author:liuming
     */
    public static void paramExport() {
        List<Map<String,Object>> datas=initData1();
        //一些常用的微操，默认值
        ExcelExportManager em=ExcelExportManager.Builder()
                .excelType(Constants.EXCEL_TYPE_XSSF)  //设置导出格式为xlsx
                .model(Constants.ROW_OVERFLOW_MODEL_NEWFILE) //设置超出sheet数据最大条数时的导出模式为生成新的excel文件，默认为创建新sheet
                .maxRow(5) //设置每个sheet最大的数据条数(表头+数据)为5，即一个sheet最多只会生成5条数据。默认为xls/xlsx支持的最大条数
                .fileName("test") //设置导出的excel文件名(不需要后缀)，不设置使用UUID/“excel”[多文件模式下使用]命名，数据量大且以压缩包生成时建议设置此参数
                .zipFileName("score") //zip文件导出模式下压缩包的文件名
                ;
        try {
            String fileName=em.createExcel(excelSaveFolder, json1, datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    /**
     * 试试五级表头？
     * @author:liuming
     */
    public static void levelFiveExport() {
        List<Map<String,Object>> datas=initData1();
        //一些常用的微操，默认值
        ExcelExportManager em=ExcelExportManager.Builder();
        try {
            String fileName=em.createExcel(excelSaveFolder, json2, datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    /**
     * 二十级表头
     * @author:liuming
     */
    public static void level20Export() {
        System.out.println("一个优秀的程序员不应该这么无聊");
    }
    /**
     * 只创建表头。如果您只想用本程序创建表头，请参考这里
     * @author:liuming
     */
    public static void createSheetTitleOnly() {
        Workbook workbook=new HSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        ExcelExportManager em=ExcelExportManager.Builder()
                .cellPropertys(ExcelUtil.jsonToListData(json1));
        em.initCellStyle(workbook);
//        int rowIndex=em.createSheetTitle(sheet);
      //您可以在指定位置插入表头
        int rowIndex=em.createSheetTitle(sheet,1,2);
        System.out.println("行索引:"+rowIndex);
        em.createExcelFile(workbook, "D:\\sheetTitle");
    }
    /**
     * 根据模板导出
     * @author:liuming
     */
    public static void exportByTemplet() {
//        System.out.println(System.getProperty("user.dir"));
//        File file=new File(System.getProperty("user.dir")+"\\model\\test.xlsx");
//        System.out.println(file.exists());
        //模板文件除了用来导出，也可以只用来做样式。cellStyleJson对应的row和col会以其name做为样式名。
        //例如：{'name':'test1',row:0,col:0}，表示模板文件第一行第一个单元格的样式名字为test1。
//        如果设置{field:'yw', title: '语文', width:10,titleStyle:'test1'}，则模板文件(0,0)的样式会应用到“语文”这列表头
        //在为单元格配置titleStyle和cellStyle时，请确保已初始化相应的样式
        TempletExcel templetExcel=new TempletExcel(System.getProperty("user.dir")+"\\model\\test.xlsx", true,cellStyleJson);
        //因为使用模板文件，需要配置模板每一列对应的数据字段名
        CellField[] fields=new CellField[7];
        fields[0]=new CellField("name",null);
        fields[1]=new CellField("sex",null);
        fields[2]=new CellField("age",null);
        fields[3]=new CellField("yw",null);
        fields[4]=new CellField("sx",null);
        fields[6]=new CellField("yy",null);
        
        List<Map<String,Object>> datas=initData1();
        
        ExcelExportManager em=ExcelExportManager.Builder()
                .fields(fields)
                .templetExcel(templetExcel)
                ;
        try {
            String fileName=em.createExcel(excelSaveFolder, "[[]]", datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    /**
     * 根据模板文件创建样式
     * @author:liuming
     */
    public static void createStyleByTemplet() {
        //在为单元格配置titleStyle和cellStyle时，请确保已初始化相应的样式
        TempletExcel templetExcel=new TempletExcel(System.getProperty("user.dir")+"\\model\\test.xlsx", false,cellStyleJson);
        ExcelExportManager em=ExcelExportManager.Builder()
                .templetExcel(templetExcel)
                ;
        List<Map<String,Object>> datas=initData1();
        try {
            String fileName=em.createExcel(excelSaveFolder, json3, datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    /**
     * 自定义单元格样式
     * @author:liuming
     */
    public static void createStyleByMe() {
        //自定义单元格样式可通过接口 WorkbookListener的addCellStyle添加
        ExcelExportManager em=ExcelExportManager.Builder()
                .workbookListener(new WorkbookListener() {
                    
                    @Override
                    public boolean workbookComplete(Workbook workbook) {
                        return false;
                    }
                    
                    @Override
                    public void updateTitleCell(Sheet sheet, Cell cell, int left, int right, int top, int bottom) {
                        
                    }
                    
                    @Override
                    public String setSheetName(int index) {
                        return null;
                    }
                    
                    @Override
                    public String setFileName(int index) {
                        return null;
                    }
                    
                    @Override
                    public void addCellStyle(Workbook workbook, Map < String, CellStyle > cellStyles) {
                        CellStyle cellStyle=workbook.createCellStyle();
                        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                        cellStyles.put("test1", cellStyle);
                      //如果需要更改默认的单元格样式，请如下设置
//                        cellStyles.put(Constants.DEFAULT_TITLE_CELLSTYLE_NAME, cellStyle); //表头
//                        cellStyles.put(Constants.DEFAULT_DATA_CELLSTYLE_NAME, cellStyle); //单元格
                    }
                });
        
        List<Map<String,Object>> datas=initData1();
        try {
            String fileName=em.createExcel(excelSaveFolder, json3, datas);
            System.out.println("生成的excel文件："+fileName);
        }
        catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    
    /**
     * 初始化数据
     * @author:liuming
     * @return
     */
    public static List<Map<String,Object>> initData1(){
        List<Map<String,Object>> list=new ArrayList<Map<String,Object>>();
        Map<String,Object> map=new HashMap<String, Object>();
        map.put("name", "狗王");
        map.put("age", 100);
        map.put("sex", "雄性");
        map.put("sx", 10);
        map.put("yy", 22.5);
        map.put("yw", 998);
        list.add(map);
        map=new HashMap<String, Object>();
        map.put("name", "大狗子");
        map.put("age", 100);
        map.put("sex", "雄性");
        map.put("sx", 10);
        map.put("yy", 22.5);
        map.put("yw", 94);
        list.add(map);
        map=new HashMap<String, Object>();
        map.put("name", "二狗子");
        map.put("age", 100);
        map.put("sex", "雄性");
        map.put("sx", 10);
        map.put("yy", 22.5);
        map.put("yw", 939);
        list.add(map);
        map=new HashMap<String, Object>();
        map.put("name", "小狗子");
        map.put("age", 100);
        map.put("sex", "雌性");
        map.put("sx", 10);
        map.put("yy", 22.5);
        map.put("yw", 92);
        list.add(map);
        map=new HashMap<String, Object>();
        map.put("name", "狗雄");
        map.put("age", 100);
        map.put("sex", "雄性");
        map.put("sx", 99);
        map.put("yy", 22.5);
        map.put("yw", 91);
        list.add(map);
        
        return list;
    }
}
