import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.xuanyimao.poiexcelexporttool.ExcelExportManager;

/**  
 * http://www.xuanyimao.com
 * @author:liuming
 * @date: 2020年10月12日
 * @version V1.0 
 */

/**
 * @Description:
 * @author liuming
 */
public class SimpleTest {
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
    
    public static void main(String[] args) {
        List<Map<String,Object>> datas=initData1();
        ExcelExportManager em=ExcelExportManager.Builder();
        try {
            String fileName=em.createExcel("D:/exceltest", json1, datas);
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
