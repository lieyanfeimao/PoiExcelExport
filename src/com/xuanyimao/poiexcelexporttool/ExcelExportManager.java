/**
 * github：https://github.com/lieyanfeimao/PoiExcelExport  
 * 码云：https://gitee.com/edadmin/PoiExcelExport
 */
package com.xuanyimao.poiexcelexporttool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xuanyimao.poiexcelexporttool.bean.CellField;
import com.xuanyimao.poiexcelexporttool.bean.CellProperty;
import com.xuanyimao.poiexcelexporttool.bean.TempletCellStyle;
import com.xuanyimao.poiexcelexporttool.bean.TempletExcel;
import com.xuanyimao.poiexcelexporttool.common.Constants;
import com.xuanyimao.poiexcelexporttool.common.ExcelUtil;
import com.xuanyimao.poiexcelexporttool.common.ZipUtil;
import com.xuanyimao.poiexcelexporttool.listener.SheetListener;
import com.xuanyimao.poiexcelexporttool.listener.WorkbookListener;

/**
 * excel导出程序主类，主要用于项目上的excel导出
 * @author liuming
 *
 */
public class ExcelExportManager {
	
	/**单元格样式集合*/
	private Map<String,CellStyle> cellStyles=new HashMap<String,CellStyle>();
	
	/**单元格属性列表*/
	private List<CellProperty[]> prosList=null;
	
	/**超出sheet数据最大条数时的导出模式，默认为创建新sheet*/
	private int model=Constants.ROW_OVERFLOW_MODEL_NEWSHEET;
	
	/**一个sheet数据最大条数*/
	private int maxRow=Constants.SHEET_MAX_SIZE_HSSF;
	
	/**excel类型：xls/xlsx*/
	private int excelType=-1;
	
	/** excel文件名 */
	private String fileName=null;
	
	/**sheet名*/
	private String sheetName=null;
	
	/** zip文件名 */
	private String zipFileName=null;
	
	/**模板样式配置*/
	private TempletExcel templetExcel;
	
	/**工作簿监听器*/
	private WorkbookListener workbookListener=null;
	
	/**Sheet监听器*/
	private SheetListener sheetListener=null;
	
	/**单元格最大个数，创建表头时计算得出*/
	private int maxCellSize=0;
	
	/**从单元格属性列表获取的字段名数组*/
	private CellField[] fields=null;
	
	/**excel文件的保存路径*/
	private String excelFolder="";
	
	/**创建的所有excel文件,包括zip*/
	private List<String> filePathList=new ArrayList<String>();
	
	/**数据插入行起始索引*/
	private int startRowIndex;
	
	private ExcelExportManager(){}
	
	/**
	 * 创建excel文件
	 * @param saveFolder 文件保存目录[必填],建议为程序单独设置一个目录
	 * @param cellJson   单元格属性值集合的json，用于生成任意级表头，属性值参考CellProperty对象
	 * @param datas      excel数据
	 * @return excel文件的完整路径
	 * @throws IOException
	 */
	public String createExcel(String saveFolder,String cellJson,List<Map<String,Object>> datas) throws IOException{
		return createExcel(saveFolder, ExcelUtil.jsonToListData(cellJson), datas);
	}
	
	/**
	 * 创建excel文件
	 * @param saveFolder 文件保存目录[必填],建议为程序单独设置一个目录
	 * @param prosList   单元格属性值集合对象，用于生成任意级表头
	 * @param datas      excel数据
	 * @return excel文件的完整路径
	 * @throws IOException 
	 */
	public String createExcel(String saveFolder,List<CellProperty[]> prosList,List<Map<String,Object>> datas) throws IOException{
		if(StringUtils.isBlank(saveFolder)){
			System.out.println("保存的目录路径不能为空!");
			return null;
		}
		//如果存在模板文件，根据模板文件名获取excel类型
		if(templetExcel!=null && templetExcel.getTempletFilePath()!=null && new File(templetExcel.getTempletFilePath()).exists()
				&& this.excelType==-1){
			this.excelType=ExcelUtil.getExcelType(templetExcel.getTempletFilePath(), this.excelType);
		}
		
		//设置生成的excel文件类型
		if(this.excelType==-1) this.excelType=ExcelUtil.getExcelType(this.fileName, this.excelType);
		
		if(excelType==Constants.EXCEL_TYPE_HSSF){
			if(this.maxRow>Constants.SHEET_MAX_SIZE_HSSF){
				this.maxRow=Constants.SHEET_MAX_SIZE_HSSF;
			}
		}else{
			if(maxRow>Constants.SHEET_MAX_SIZE_XSSF){
				this.maxRow=Constants.SHEET_MAX_SIZE_XSSF;
			}
		}
		
		//获取单元格属性对象集合
		this.prosList=prosList;
		
		//获取文件创建目录
		if(!saveFolder.endsWith("\\") && !saveFolder.endsWith("/")){
			saveFolder+=File.separator;
		}
		//设置excel保存目录
		this.excelFolder=saveFolder+ExcelUtil.getFolderName()+File.separator;
		//创建目录
		File file=new File(this.excelFolder);
		if(!file.isDirectory()){
			file.mkdirs();
		}
		
		//填充数据并生成excel文件
		String filePath="";
		if(this.model == Constants.ROW_OVERFLOW_MODEL_NEWFILE){//多excel文件模式
			filePath=createExcelByZipModel(fileName,zipFileName, datas, workbookListener);
		}else{//多sheet模式
			filePath=createExcelBySheetModel(fileName, datas, workbookListener);
		}
		return filePath;
	}
	
	/**
	 * 以多sheet模式创建excel表格
	 * @param fileName  文件名
	 * @param datas     数据集合
	 * @param workbookListener  工作簿监听器
	 * @return  返回文件路径
	 * @throws IOException 
	 */
	public String createExcelBySheetModel(String fileName,List<Map<String,Object>> datas,WorkbookListener workbookListener) throws IOException{
		Workbook workbook=initCellStyle();
		//以第一个Sheet为模板
		workbook.setSheetName(0,"model");
		int sheetIndex=0;
		int dataIndex=0;
		while(true){
			String sheetName=this.sheetName;
			Sheet sheet=workbook.cloneSheet(0);
			if(workbookListener!=null){
				sheetName=workbookListener.setSheetName(sheetIndex);
			}
			if(StringUtils.isBlank(sheetName)){
				sheetName="Sheet";
			}
			
			workbook.setSheetName(sheetIndex+1, sheetName+(sheetIndex==0?"":sheetIndex));
			//填充表格内容，数据写入完成返回true
//			boolean flag=addDataToSheet(sheet, datas,startRowIndex);
			boolean flag=false;
            if(sheetListener!=null){
                int result=sheetListener.addDataToSheet(sheet, startRowIndex, maxRow);
                if(result==-1){
//                    flag=addDataToSheet(sheet, datas,startRowIndex);
                    dataIndex=addDataToSheet(sheet, datas,startRowIndex,dataIndex);
                    flag=dataIndex==datas.size();
//                    System.out.println(dataIndex);
                }else if(result==0){
                    flag=true;
                }
            }else{
//                flag=addDataToSheet(sheet, datas,startRowIndex);
                dataIndex=addDataToSheet(sheet, datas,startRowIndex,dataIndex);
                flag=dataIndex==datas.size();
//                System.out.println(dataIndex);
            }
			
			sheetIndex++;
			if(flag){
				break;
			}
		}
		//移除Sheet模板
		workbook.removeSheetAt(0);
		
		if(workbookListener!=null) workbookListener.workbookComplete(workbook);
		
		String filePath=createExcelFile(workbook,fileName);
		filePathList.add(filePath);//将excel文件添加到路径
		return filePath;
	}
	
	/**
	 * 以多个excel文件格式创建excel
	 * @param fileName  文件名
	 * @param zipFileName  压缩文件名
	 * @param datas     数据集合
	 * @param workbookListener   工作簿监听器
	 * @return
	 * @throws IOException 
	 */
	public String createExcelByZipModel(String fileName,String zipFileName,List<Map<String,Object>> datas,WorkbookListener workbookListener) throws IOException{
		int fileIndex=0;
		String filePath="";
		int dataIndex=0;
		while(true){
			Workbook workbook=initCellStyle();
			Sheet sheet=workbook.getSheetAt(0);
			//填充表格内容，数据写入完成返回true
			boolean flag=false;
			if(sheetListener!=null){
                int result=sheetListener.addDataToSheet(sheet, startRowIndex, maxRow);
                if(result==-1){
//                    flag=addDataToSheet(sheet, datas,startRowIndex);
                    dataIndex=addDataToSheet(sheet, datas,startRowIndex,dataIndex);
                    flag=dataIndex==datas.size();
                }else if(result==0){
                    flag=true;
                }
            }else{
//                flag=addDataToSheet(sheet, datas,startRowIndex);
                dataIndex=addDataToSheet(sheet, datas,startRowIndex,dataIndex);
                flag=dataIndex==datas.size();
            }
			
			//将文件保存到excel中
			if(workbookListener!=null){
				fileName=workbookListener.setFileName(fileIndex);
				workbookListener.workbookComplete(workbook);
			}
			
			if(StringUtils.isBlank(fileName)){
				fileName=Constants.DEFAULT_MODEL_NEWFILE_FILENAME;
			}
			filePath=createExcelFile(workbook,fileName+(fileIndex==0?"":fileIndex));
			filePathList.add(filePath);//将excel文件添加到路径
			fileIndex++;
			workbook=null;
			if(flag){
				break;
			}
		}
		if(filePathList.size()>1){
			//压缩excel文件到zip目录
			filePath=this.excelFolder+(StringUtils.isBlank(zipFileName)?UUID.randomUUID().toString().replace("-",""):zipFileName)+".zip";
			try {
				ZipUtil.folderToZip(filePathList, filePath);
			} catch (Exception e) {
				e.printStackTrace();
				System.out.println("Excel文件压缩失败");
			}
			filePathList.add(filePath);//将zip文件添加到路径
		}
		return filePath;
	}
	
	/**
	 * 创建sheet的表头
	 * @param sheet sheet对象
	 * @return 行索引
	 */
	public int createSheetTitle(Sheet sheet){
		return createSheetTitle(sheet, 0, 0);
	}
	
	/**
	 * 创建sheet的表头
	 * @param sheet sheet对象
	 * @param rindex 起始行索引
	 * @param cindex 起始列索引
	 * @return
	 */
	public int createSheetTitle(Sheet sheet,int rindex,int cindex){
		if(prosList==null || prosList.isEmpty()){
			return rindex;
		}
		//记录每行已经被使用的单元格
		List<List<Integer>> rowList=new ArrayList<List<Integer>>();
		int colspan=0;//单元格跨列
		int rowspan=0;//单元格跨行
		int rowindex=rindex;//行号
		int colindex=cindex;//列号
		Map<Integer,CellField> fieldMap=new HashMap<Integer,CellField>();//记录字段集合
		for(CellProperty[] pros:prosList){
			colindex=cindex;
			Row row = sheet.createRow(rowindex);
			if(rowList.size()<=0){
				rowList.add(new ArrayList<Integer>());
			}
			List<Integer> colList=rowList.get(0);
			
			for(CellProperty cellProperty:pros){//遍历单元格对象
				for(int i=colindex;i<=maxCellSize;i++){
					if(!colList.contains(i)){
						colindex=i;
						break;
					}
				}
				
				if(cellProperty.getColspan()==null || cellProperty.getColspan()<=0){
					colspan=0;
				}else{
					colspan=cellProperty.getColspan()-1;
				}
				if(cellProperty.getRowspan()==null || cellProperty.getRowspan()<=0){
					rowspan=0;
				}else{
					rowspan=cellProperty.getRowspan()-1;
				}
				Cell cell=row.createCell(colindex);
				sheet.addMergedRegion(new CellRangeAddress(rowindex,rowindex+rowspan,colindex,colindex+colspan));
				cell.setCellValue(cellProperty.getTitle());
				if(cellProperty.getTitleStyle()==null || cellStyles.get(cellProperty.getTitleStyle())==null ){
					cell.setCellStyle(cellStyles.get(Constants.DEFAULT_TITLE_CELLSTYLE_NAME));
				}else{
					cell.setCellStyle(cellStyles.get(cellProperty.getTitleStyle()));
				}
				if(rowspan>0){//如果纵向跨度超过一个单元格
					for(int i=1;i<=rowspan;i++){//将后续行被使用的单元格添加到集合
						if(rowList.size()<=i){
							rowList.add(new ArrayList<Integer>());
						}
						List<Integer> list=rowList.get(i);
						for(int j=0;j<=colspan;j++){
							list.add(colindex+j);
						}
					}
				}
				//获取当前列的字段名
				if(cellProperty.getField()!=null){
					fieldMap.put(colindex,new CellField(cellProperty.getField(),cellProperty.getCellStyle()));
				}
				//设置单元格宽度
				if(cellProperty.getWidth()!=null){
					sheet.setColumnWidth(colindex, cellProperty.getWidth()*256);
				}
				//添加批注
				if(cellProperty.getComment()!=null){
					if(sheet instanceof XSSFSheet){
						XSSFDrawing draw=(XSSFDrawing)sheet.createDrawingPatriarch();
						XSSFComment comment=draw.createCellComment(new XSSFClientAnchor(0,0,0,0,(short)3,3,(short)5,6));
						comment.setString(new XSSFRichTextString(cellProperty.getComment()));
						cell.setCellComment(comment);
					}else if(sheet instanceof HSSFSheet){
						HSSFPatriarch patriarch=(HSSFPatriarch)sheet.createDrawingPatriarch();
						HSSFComment comment=patriarch.createComment(new HSSFClientAnchor(0,0,0,0,(short)3,3,(short)5,6));
						comment.setString(new HSSFRichTextString(cellProperty.getComment()));
						cell.setCellComment(comment);
					}
				}
				//自定义处理
				if(workbookListener!=null) workbookListener.updateTitleCell(sheet, cell, rowindex,rowindex+rowspan,colindex,colindex+colspan);
				
				colindex+=colspan+1;
			}
			if(maxCellSize<colindex) maxCellSize=colindex;
			rowList.remove(0);
			rowindex++;
		}
		
		if(fieldMap.size()>0){
			fields=new CellField[maxCellSize];
			for(int i=0;i<maxCellSize;i++){
				if(fieldMap.get(i)!=null){
					fields[i]=fieldMap.get(i);
				}
			}
		}
		
		return rowindex;
	}
	/**
	 * 向excel填充数据，单独调用此方法请先调用fields()方法，否则会没数据
	 * @param sheet sheet对象
	 * @param datas 数据
	 * @param rowindex 起始索引
	 * @param dataIndex 数据(datas参数)起始索引
	 * @return
	 */
	public int addDataToSheet(Sheet sheet,List<Map<String,Object>> datas,int rowindex,int dataIndex){
		if(datas!=null && fields!=null){
			for(int i=dataIndex;i<datas.size();i++){
//				Row row = sheet.createRow(rowindex);
				Row row = sheet.getRow(rowindex);
				if(row==null){
					row = sheet.createRow(rowindex);
				}
				Map<String,Object> map=datas.get(i);
				for(int j=0;j<fields.length;j++){
					if(fields[j]!=null){
//						Cell cell=row.createCell(j);
						Cell cell=row.getCell(j);
						if(cell==null){
							cell=row.createCell(j);
						}
						cell.setCellValue(map.get(fields[j].getField())==null?"":map.get(fields[j].getField()).toString());
						//设置单元格样式
						if(fields[j].getCellStyle()==null || cellStyles.get(fields[j].getCellStyle())==null ){
							cell.setCellStyle(cellStyles.get(Constants.DEFAULT_DATA_CELLSTYLE_NAME));
						}else{
							cell.setCellStyle(cellStyles.get(fields[j].getCellStyle()));
						}
					}
				}
				//去掉,浪费时间。这是一个愚蠢的写法
//				datas.remove(0);
				
				rowindex++;
                if(rowindex>=this.maxRow){
                    return i+1;
                }
			}
		}
		return datas.size();
	}
	
	/***
	 * 创建excel文件
	 * @param workbook 工作簿对象
	 * @param fileName 文件名
	 * @return
	 */
	public String createExcelFile(Workbook workbook,String fileName){
		if(StringUtils.isBlank(fileName)){
			fileName=UUID.randomUUID().toString().replace("-","");
		}
		String filePath=this.excelFolder+fileName;
		if(!fileName.toLowerCase().endsWith(".xls") && !fileName.toLowerCase().endsWith(".xlsx")){//不是excel文件后缀结尾的根据excelType追加后缀
			filePath+=(excelType==Constants.EXCEL_TYPE_XSSF)?".xlsx":".xls";
		}
		FileOutputStream stream=null;
		try {
			stream = new FileOutputStream(filePath);
			workbook.write(stream);
			stream.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return filePath;
	}
	/**
	 * 删除excel文件和目录，用于excel文件无用时删除文件清理磁盘空间
	 * 应该在excel文件不需要使用的时候再调用此方法，文件被占用时调用会导致文件删除失败，留下残留文件
	 */
	public void deleteExcelFolder(){
		for(String filePath:filePathList){
			File file=new File(filePath);
			file.delete();
		}
		File file=new File(this.excelFolder);
		file.delete();//删除非空目录
	}
	
	/**
	 * 初始化单元格样式
	 * @param workbook
	 */
	public void initCellStyle(Workbook workbook){
		putCellStyle(Constants.DEFAULT_TITLE_CELLSTYLE_NAME,ExcelUtil.getDefaultTitleCellStyle(workbook));
		putCellStyle(Constants.DEFAULT_DATA_CELLSTYLE_NAME,ExcelUtil.getDefaultDataCellStyle(workbook));
		//通过接口添加单元格样式
		if(workbookListener!=null) workbookListener.addCellStyle(workbook,cellStyles);
	}
	/**
	 * 初始化样式
	 * @return 工作簿对象
	 * @throws IOException 
	 */
	public Workbook initCellStyle() throws IOException{
		//工作簿对象
		Workbook workbook=null;
		//存在模板文件
		if(templetExcel!=null && templetExcel.getTempletFilePath()!=null && new File(templetExcel.getTempletFilePath()).exists()){
			FileInputStream fis=new FileInputStream(templetExcel.getTempletFilePath());
			if(excelType==Constants.EXCEL_TYPE_HSSF){
				workbook=new HSSFWorkbook(fis);
			}else{
				workbook=new XSSFWorkbook(fis);
			}
			fis.close();
			//通过接口添加样式
			initCellStyle(workbook);
			
			//获取第一个sheet
			Sheet sheet=workbook.getSheetAt(0);
			if(templetExcel.getTempletCellStyles()!=null && !templetExcel.getTempletCellStyles().isEmpty()){
				for(TempletCellStyle templetCellStyle:templetExcel.getTempletCellStyles()){
					cellStyles.put(templetCellStyle.getName(),sheet.getRow(templetCellStyle.getRow()).getCell(templetCellStyle.getCol()).getCellStyle());
				}
			}
//			System.out.println("最后一行的索引："+sheet.getLastRowNum()+"  第一行索引："+sheet.getFirstRowNum()+"  有效行索引："+sheet.getPhysicalNumberOfRows());
//			System.out.println("合并的单元格的个数："+sheet.getNumMergedRegions());
			//如果不是依照模板导出，清空所有行，创建表头
			if(!templetExcel.isExportModel()){
				//移除合并的单元格
				for(int i=sheet.getNumMergedRegions()-1;i>=0;i--){
//					CellRangeAddress ca=sheet.getMergedRegion(i);
//					System.out.println(ca.getFirstRow()+"="+ca.getFirstColumn()+"="+ca.getLastRow()+"="+ca.getLastColumn());
					sheet.removeMergedRegion(i);
				}
				//移除行
				for(int i=sheet.getLastRowNum();i>=0;i--){
//					System.out.println(i);
					Row row=sheet.getRow(i);
					if(row!=null) sheet.removeRow(row);
				}
				
				startRowIndex=createSheetTitle(sheet);
			}else{
				if(templetExcel.getStartRowIndex()!=null && templetExcel.getStartRowIndex()>=0){
					startRowIndex=templetExcel.getStartRowIndex();
				}else{
					startRowIndex=sheet.getLastRowNum()+1;
				}
			}
		}else{
			if(excelType==Constants.EXCEL_TYPE_HSSF){
				workbook=new HSSFWorkbook();
			}else{
				workbook=new XSSFWorkbook();
			}
			//通过接口添加样式
			initCellStyle(workbook);
			Sheet sheet=workbook.createSheet();
			startRowIndex=createSheetTitle(sheet);
		}
		return workbook;
	}
	
	/*********************************** 属性初始化部分 **********************************/
	/***
	 * 获取一个对象实例
	 * @return
	 */
	public static ExcelExportManager Builder(){
		return new ExcelExportManager();
	}
	/**
	 * 设置fields的值，使用模板导出时需设置此字段的值
	 * @param fields
	 * @return
	 */
	public ExcelExportManager fields(CellField[] fields){
		this.fields=fields;
		return this;
	}
	
	/**
	 * 超出sheet数据最大条数时的导出模式，默认为创建新sheet。xls文件最多为65536条,xlsx文件最多为1048576行
	 * @param model 导出模式
	 * @return
	 */
	public ExcelExportManager model(int model){
		if(model!=Constants.ROW_OVERFLOW_MODEL_NEWFILE && model!=Constants.ROW_OVERFLOW_MODEL_NEWSHEET){
			System.out.println("导出模式无法识别，使用默认导出模式：创建新sheet!");
			return this;
		}
		this.model=model;
		return this;
	}
	
	/***
	 * 添加一个单元格样式
	 * @param name
	 * @param cellStyle
	 */
	private ExcelExportManager putCellStyle(String name,CellStyle cellStyle){
		cellStyles.put(name, cellStyle);
		return this;
	}
	
	/**
	 * 添加模板样式对象
	 * @param templetStyle 模板样式对象
	 * @return
	 */
	public ExcelExportManager templetExcel(TempletExcel templetExcel){
		this.templetExcel=templetExcel;
		return this;
	}
	
	/**
	 * 设置每个sheet最大的数据条数(表头+数据)，超出最大值会以最大值计算。xls文件最大65535，xlsx文件最大1048575
	 * @param maxRow
	 * @return
	 */
	public ExcelExportManager maxRow(int maxRow){
		this.maxRow=maxRow;
		return this;
	}
	
	/**
	 * 设置Excel类型，[瞎几把传会采用默认的xls文件类型]
	 * @param excelType
	 * 			xls：com.xuanyimao.poiexcelexporttool.common.Constants.EXCEL_TYPE_HSSF
	 * 			xlsx：com.xuanyimao.poiexcelexporttool.common.Constants.EXCEL_TYPE_XSSF
	 * @return
	 */
	public ExcelExportManager excelType(int excelType){
		if(excelType!= Constants.EXCEL_TYPE_HSSF && excelType!=Constants.EXCEL_TYPE_XSSF){
			System.out.println("excel类型无法识别，使用默认模式：生成xls文件!");
			excelType = Constants.EXCEL_TYPE_HSSF;
		}
		this.excelType=excelType;
		return this;
	}
	
	/**
	 * 设置导出的excel文件名(不需要后缀)，不设置使用UUID/“excel”[多文件模式下使用]命名，数据量大且以压缩包生成时建议设置此参数
	 * @param fileName
	 * @return
	 */
	public ExcelExportManager fileName(String fileName){
		this.fileName=fileName;
		return this;
	}
	
	/**
	 * zip文件导出模式下压缩包的文件名
	 * @param zipFileName
	 * @return
	 */
	public ExcelExportManager zipFileName(String zipFileName){
		this.zipFileName=zipFileName;
		return this;
	}
	
	/**
	 * 设置sheet名
	 * @param sheetName
	 * @return
	 */
	public ExcelExportManager sheetName(String sheetName){
		this.sheetName=sheetName;
		return this;
	}
	
	/**
	 * 设置工作薄监听器
	 * @param workbookListener
	 * @return
	 */
	public ExcelExportManager workbookListener(WorkbookListener workbookListener){
		this.workbookListener=workbookListener;
		return this;
	}
	/**
	 * 设置sheet监听器
	 * @param sheetListener
	 * @return
	 */
	public ExcelExportManager sheetListener(SheetListener sheetListener){
		this.sheetListener=sheetListener;
		return this;
	}
	
	/**
	 * 设置单元格属性集合，此设置仅在单独创建表头时有效
	 * @author:liuming
	 * @param cellPropertys
	 * @return
	 */
	public ExcelExportManager cellPropertys(List<CellProperty[]> cellPropertys) {
	    this.prosList=cellPropertys;
	    return this;
	}
	
	/************** 已弃用的方法 *******************/
	/**
	 * 创建excel文件
	 * @param saveFolder 文件保存目录[必填],建议为程序单独设置一个目录
	 * @param fileName   文件名(不需要后缀)[非必填]，不设置使用UUID命名，数据量大且以压缩包生成时建议设置此参数
	 * @param excelType  生成的excel类型[瞎几把传会采用默认的xls文件类型],EXCEL_TYPE_HSSF(xls)或者EXCEL_TYPE_XSSF(xlsx)
	 * @param maxRow     每个sheet最大的数据条数(表头+数据)，超出最大值会以最大值计算。xls文件最大65535，xlsx文件最大1048575。
	 * @param cellJson   单元格属性值集合的json，用于生成任意级表头，属性值参考CellProperty对象
	 * @param datas      excel数据
	 * @param workbookListener  工作簿监听器
	 * @return excel文件的完整路径
	 * @throws IOException 
	 */
	@Deprecated
	public String createExcel(String saveFolder,String fileName,int excelType,int maxRow,String cellJson,List<Map<String,Object>> datas,WorkbookListener workbookListener) throws IOException{
		return createExcel(saveFolder, fileName, excelType, maxRow, ExcelUtil.jsonToListData(cellJson), datas, workbookListener);
	}
	/**
	 * 创建excel文件
	 * @param saveFolder 文件保存目录[必填],建议为程序单独设置一个目录
	 * @param fileName   文件名(不需要后缀)[非必填]，不设置使用UUID命名，数据量大且以压缩包生成时建议设置此参数
	 * @param excelType  生成的excel类型[瞎几把传会采用默认的xls文件类型],EXCEL_TYPE_HSSF(xls)或者EXCEL_TYPE_XSSF(xlsx)
	 * @param maxRow     每个sheet最大的数据条数(表头+数据)，超出最大值会以最大值计算。xls文件最大65535，xlsx文件最大1048575。
	 * @param prosList   单元格属性值集合对象，用于生成任意级表头
	 * @param datas      excel数据
	 * @param workbookListener  工作簿监听器
	 * @return excel文件的完整路径
	 * @throws IOException 
	 */
	@Deprecated
	public String createExcel(String saveFolder,String fileName,int excelType,int maxRow,List<CellProperty[]> prosList,List<Map<String,Object>> datas,WorkbookListener workbookListener) throws IOException{
		return createExcel(saveFolder, fileName, "", excelType, maxRow, prosList, datas, workbookListener);
	}
	/**
	 * 创建excel文件
	 * @param saveFolder 文件保存目录[必填],建议为程序单独设置一个目录
	 * @param fileName   文件名(不需要后缀)[非必填]，不设置使用UUID命名，数据量大且以压缩包生成时建议设置此参数
	 * @param excelType  生成的excel类型[瞎几把传会采用默认的xls文件类型],EXCEL_TYPE_HSSF(xls)或者EXCEL_TYPE_XSSF(xlsx)
	 * @param maxRow     每个sheet最大的数据条数(表头+数据)，超出最大值会以最大值计算。xls文件最大65535，xlsx文件最大1048575。
	 * @param prosList   单元格属性值集合对象，用于生成任意级表头
	 * @param datas      excel数据
	 * @param workbookListener  工作簿监听器
	 * @return excel文件的完整路径
	 * @throws IOException 
	 */
	@Deprecated
	public String createExcel(String saveFolder,String fileName,String zipFileName,int excelType,int maxRow,List<CellProperty[]> prosList,List<Map<String,Object>> datas,WorkbookListener workbookListener) throws IOException{
		this.fileName=fileName;
		this.zipFileName=zipFileName;
		this.maxRow=maxRow;
		if(excelType!= Constants.EXCEL_TYPE_HSSF && excelType!=Constants.EXCEL_TYPE_XSSF){
			System.out.println("excel类型无法识别，使用默认模式：生成xls文件!");
			excelType = Constants.EXCEL_TYPE_HSSF;
		}
		this.excelType=excelType;
		
		return createExcel(saveFolder, prosList, datas);
	}

    /**
     * 获取本次导出产生的所有文件路径，包含最终导出的文件
     * @return filePathList
     */
    public List < String > getFilePathList() {
        return filePathList;
    }
}
