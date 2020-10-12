## PoiExcelExport用于处理Excel多级表头数据导出，您只需要给它一个类似于easyui或laytable格式的cols，或者，把它们的cols稍加修改，就可以用于PoiExcelExport。
## PoiExcelExport is used to process Excel multi-level header data export. You only need to give it a cols similar to easyui or laytable format, or you can use PoiExcelExport by modifying their cols slightly.
## 因为我在项目中使用的是poi-3.7，所以本程序使用的也是poi-3.7，我粗略测试了一下，最新的版本里面已经去掉了很多东西，本程序无法在最新版本的poi下使用，没测试最高能兼容到哪个版本。本程序可直接用于导出，有多种模式和配置，但是不建议用于大数据导出（怎样算大数据？看你服务器有多强）。大数据导出个人觉得应该用POI另一种不消耗内存的模式。所以，您可以参考本程序的设计思路，自行设计用于大数据导出的程序。
## 本项目用于导出的数据格式为List<Map<String,Object>>，您的数据可能是List<Object>，这不支持。请自行编写程序使用反射进行导出。或者，只使用本项目的生成表头功能。本项目创建的初衷便是创建表头，生成Excel只是附带功能  
## PoiExcelExport的设计思路请参阅doc目录下的index.html。PoiExcelExport的测试类是src目录下的Test.java
## github：https://github.com/lieyanfeimao/PoiExcelExport  
## 码云：https://gitee.com/edadmin/PoiExcelExport
## 若在使用过程中发现BUG，请尽管提出，反正我不会改。

## 工程环境
JDK1.7+，将lib下的jar包Add to build path即可。项目里用json-lib来处理json数据，它需要关联一些jar包，建议用Gson  
若需换成Gson或其他Json jar，请重写**com.xuanyimao.poiexcelexporttool.common.ExcelUtil**中的**jsonToListData**和**jsonToCellStyles**方法(Gson的处理程序已写好，打开注释即可)  
POIMaven引入参考

```xml
    <dependency>
		<groupId>org.apache.poi</groupId>
		<artifactId>poi</artifactId>
		<version>3.7</version>
	</dependency>
	<dependency>
	    <groupId>org.apache.poi</groupId>
	    <artifactId>poi-ooxml</artifactId>
	    <version>3.7</version>
	</dependency>
	<dependency>
	    <groupId>org.apache.poi</groupId>
	    <artifactId>poi-ooxml-schemas</artifactId>
	    <version>3.7</version>
	</dependency>
	<dependency>
	    <groupId>commons-codec</groupId>
	    <artifactId>commons-codec</artifactId>
	    <version>1.11</version>
	</dependency>
	<dependency>
	    <groupId>commons-beanutils</groupId>
	    <artifactId>commons-beanutils</artifactId>
	    <version>1.9.3</version>
	</dependency>
```

## 示例代码(二级表头)
```java
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
        //简单的导出演示
        normalExport();
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
```
以上代码可生成一个二级表头Excel，建议先在JSON编写工具里写好表头json，再粘贴到Java代码中  

## JSON结构详解：  
表格配置数据为一个二维数组对象，数组的每个维度代表一行，即array\[0\]\[0\]为第一行，array\[0\]\[1\]为第二行    
二维数组内是一个对象，包含表头和单元格的配置属性  

## JSON属性值详解：  
**field**：字段名。对应数据集合(List<Map<String,Object>>)中Map的Key。多级表头中，一列只需设置一次此值。  
**title**：列名。Excel表头的列名  
**width**：单元格宽度。一般一个汉字的宽度为2  
**colspan**：单元格跨多少列，默认为1。横向合并指定个数的单元格  
**rowspan**：单元格跨多少行，默认为1。纵向合并指定个数单元格。  
**titleStyle**：表头样式。可通过接口设置或通过模板文件设置  
**cellStyle**：单元格样式。可通过接口设置或通过模板文件设置  
**comment**：批注。有时候设置了不显示，原因不明。  

## Excel模板文件：  
模板文件可用于按照模板生成Excel，也可以只用于获取单元格样式。  
如果导出数据的单元格样式来自于模板文件，则需要配置模板对象的templetCellStyles属性，以告诉程序如何处理指定位置的单元格样式  

个人网站：http://xuanyimao.com  
CSDN博客：https://blog.csdn.net/xymmwap  