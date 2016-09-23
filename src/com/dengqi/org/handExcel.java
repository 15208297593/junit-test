package com.dengqi.org;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

public class handExcel {
	//public List<List<String>> readXls(String path) throws Exception{
	
	/**
	 * @param path
	 * @return
	 * @throws Exception
	 */
	List<String> formated  = new ArrayList<String>();
	List<String> oneRes = new ArrayList<>();
	List<String> twoRes = new ArrayList<>();
	@Test
	public void writeXls() throws Exception{
		String path = "D:\\regInfo.xls";
		List<List<String>> readRes1 = readXls1("D:\\info1.xls");
		List<List<String>> readRes2 = readXls2("D:\\info2.xls");
		List<String> readRes3 = readXls3("D:\\info3.xls");
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		//循环每一页,并处理当前循环页
		for(int numSheet = 0 ;numSheet < hssfWorkbook.getNumberOfSheets();numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			System.out.println("当前是第"+numSheet+"页");
			if(hssfSheet == null){
				System.out.println("hssfSheet is null");
				continue;
			}
			
			for(int i = 0 ;i<readRes1.size();i++){
				HSSFRow hssfRow = hssfSheet.createRow(i+1);
				//遍历该行	,获取处理每个cell元素
				for(int colIx = 0; colIx < readRes1.get(i).size(); colIx++){
					HSSFCell hssfCell = hssfRow.createCell(colIx);
					hssfCell.setCellValue(readRes1.get(i).get(colIx));
				}
				for(int colIx = 10; colIx<15; colIx++){
					HSSFCell hssfCell = hssfRow.createCell(colIx);
					hssfCell.setCellValue(readRes2.get(i).get(colIx-10));
				}
					HSSFCell hssfCell = hssfRow.createCell(15);
					hssfCell.setCellValue(readRes3.get(i));
			}
			
		}
		//创建文件流   
        OutputStream stream = new FileOutputStream(path);  
        //写入数据   
        hssfWorkbook.write(stream);  
        //关闭文件流   
        stream.close();
        is.close();
        hssfWorkbook.close();
	}
	
	
	
	
	public List<List<String>> readXls1(String path) throws Exception{
		List<List<String>> result = new ArrayList<List<String>>();
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		//循环每一页,并处理当前循环页
		for(int numSheet = 0 ;numSheet < hssfWorkbook.getNumberOfSheets();numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null){
				continue;
			}
			//处理当前页,循环读取每一行
			for(int rowNum = 1;rowNum <= hssfSheet.getLastRowNum(); rowNum++){
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				int minColIx = hssfRow.getFirstCellNum();
				int maxColIx = hssfRow.getLastCellNum();
				//遍历该行	,获取处理每个cell元素
				for(int colIx = minColIx;colIx<maxColIx;colIx++){
					HSSFCell hssfCell = hssfRow.getCell(colIx);
					if(hssfCell == null){
						continue;
					}					
					List<String> oneRes = format0(hssfCell.toString());
					result.add(oneRes);
					
				}
			}
			
		}
		is.close();
		hssfWorkbook.close();
		return result;
	}
	
	public List<List<String>> readXls2(String path) throws Exception{
		List<List<String>> result2 = new ArrayList<List<String>>();
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		//循环每一页,并处理当前循环页
		for(int numSheet = 0 ;numSheet < hssfWorkbook.getNumberOfSheets();numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null){
				continue;
			}
			//处理当前页,循环读取每一行
			for(int rowNum = 1;rowNum <= hssfSheet.getLastRowNum(); rowNum++){
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				int minColIx = hssfRow.getFirstCellNum();
				int maxColIx = hssfRow.getLastCellNum();
				//遍历该行	,获取处理每个cell元素
				for(int colIx = minColIx;colIx<maxColIx;colIx++){
					HSSFCell hssfCell = hssfRow.getCell(colIx);
					if(hssfCell == null){
						continue;
					}					
					List<String> twoRes = format1(hssfCell.toString());
					formated = formatUserOptions(twoRes.get(0), twoRes.get(1),twoRes.get(2),twoRes.get(3),twoRes.get(4));
					result2.add(formated);
					
				}
			}
			
		}
		is.close();
		hssfWorkbook.close();
		return result2;
	}
	
	
	public List<String> readXls3(String path) throws Exception{
		List<String> result3 = new ArrayList<String>();
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		//循环每一页,并处理当前循环页
		for(int numSheet = 0 ;numSheet < hssfWorkbook.getNumberOfSheets();numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null){
				continue;
			}
			//处理当前页,循环读取每一行
			for(int rowNum = 1;rowNum <= hssfSheet.getLastRowNum(); rowNum++){
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				int minColIx = hssfRow.getFirstCellNum();
				int maxColIx = hssfRow.getLastCellNum();
				//遍历该行	,获取处理每个cell元素
				for(int colIx = minColIx;colIx<maxColIx;colIx++){
					HSSFCell hssfCell = hssfRow.getCell(colIx);
					if(hssfCell == null){
						continue;
					}					
					result3.add(hssfCell.getStringCellValue());
				}
			}
		}
		is.close();
		hssfWorkbook.close();
		return result3;
	}
	@Test
	public void readXlsTest() throws Exception{
		String path = "D:\\info2.xls";
		List<List<String>> result = new ArrayList<List<String>>();
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		
		//循环每一页,并处理当前循环页
		for(int numSheet = 0 ;numSheet < hssfWorkbook.getNumberOfSheets();numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null){
				continue;
			}
			//处理当前页,循环读取每一行
			for(int rowNum = 1;rowNum <= hssfSheet.getLastRowNum(); rowNum++){
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				int minColIx = hssfRow.getFirstCellNum();
				int maxColIx = hssfRow.getLastCellNum();
				//遍历该行	,获取处理每个cell元素
				for(int colIx = minColIx;colIx<maxColIx;colIx++){
					HSSFCell hssfCell = hssfRow.getCell(colIx);
					if(hssfCell == null){
						continue;
					}
					 oneRes = format1(hssfCell.toString());
					 formated = formatUserOptions(oneRes.get(0),oneRes.get(1),oneRes.get(2),oneRes.get(3),oneRes.get(4));
					 result.add(formated);
					
				}
			}
		}
			System.out.println(result.size());
			for(List<String> one : result){
				System.out.println(one);
			}
		is.close();
		hssfWorkbook.close();
		
//		return result;
	}
public List<String> format0(String str){
		
		String[] strArray = str.split(",");
		List<String> lstr1 = new ArrayList<String>();
		List<String> lstr2 = new ArrayList<String>();
		for(int i = 0 ;i<strArray.length;i++){
			if(i%2 == 0){
				continue;
			}else{
			lstr1.add(strArray[i]);
			}
		}
		for(String str3:lstr1){
			String new1 = str3.replace("\"value\":\"", "").replace("\"}", "").replace("]", "");
			lstr2.add(new1);
		}
		return lstr2;
	}


@Test
public void formatTest(){
	String str ="[{\"optionId\":\"f332cba5-61bd-4b07-8685-1bb3ed07df45\"},"
			+ "{\"optionId\":\"d1321076-c6ad-4355-930a-38cc3a780480\"},"
			+ "{\"optionId\":\"00999fcc-33a9-446a-887f-35a9e5a087e9\"}]";
	String[] strArray = str.split(",");
	
	List<String> lstr1 = new ArrayList<String>();
	List<String> lstr2 = new ArrayList<String>();
	for(int i = 0 ;i<strArray.length;i++){
		String new1 = strArray[i].replace("{\"optionId\":\"","").replace("\"}","").replace("[","").replace("]", "");
		lstr1.add(new1);
	}
	
	System.out.println(lstr1);
	System.out.println(lstr1.get(0));
	}


public List<String> format1(String str){
	/*String str ="[{\"optionId\":\"f332cba5-61bd-4b07-8685-1bb3ed07df45\"},"
			+ "{\"optionId\":\"d1321076-c6ad-4355-930a-38cc3a780480\"},"
			+ "{\"optionId\":\"00999fcc-33a9-446a-887f-35a9e5a087e9\"}]";*/
	String[] strArray = str.split(",");
	
	List<String> lstr1 = new ArrayList<String>();
	for(int i = 0 ;i<strArray.length;i++){
		String new1 = strArray[i].replace("{\"optionId\":\"","").replace("\"}","").replace("[","").replace("]", "");
		lstr1.add(new1);
	}
	return lstr1;
	}
@SuppressWarnings({ "rawtypes", "unused", "unchecked" })

public List<String> formatUserOptions(String sex,String one,String two,String three,String change){
		Map map1 = new HashMap();
		Map map2 = new HashMap();
		Map map3 = new HashMap();
		Map mapSex = new HashMap();
		Map mapChange = new HashMap();
		map1.put("0d341dc3-0f26-469d-a0f7-7716ee69f8b6","团委办公室");
		map1.put("f7f5012d-8aab-4ac8-b9ac-af45e700fe68","团委组织部");
		map1.put("3be3f4df-dbb6-412d-97b8-015b3f60037c","团委宣传部");
		map1.put("a88ecf6c-1862-4617-9cf4-b17133fa66a0","团委信息技术中心");
		map1.put("62e66168-e2a8-443a-ae27-8e0c4dd761bf","团委志愿者工作部");
		map1.put("155af65d-0a9f-4bae-8df1-bed6d780dd85","团委社会实践部");
		map1.put("0651044d-250a-4c1b-a5ab-846374216aff","学生会办公室");
		map1.put("6fb86e5d-9bd4-46c0-ae59-757cd46d6ebc","学生会纪检部");
		map1.put("94fb6145-c646-4ec9-aa84-6b807b1677c5","学生会学术部");
		map1.put("0a8a6d08-03c0-4311-8455-166273a1ad95","学生会职业发展部");
		map1.put("7bd736fb-1f6e-4ffb-afa4-06012be2add7","学生会文体部");
		map1.put("7217b28c-92a6-4c01-b8f3-d9182303cad1","学生会生活权益部");
		map1.put("ee9ca4c2-5945-4b55-a577-dfc65230d960","学生会女生部");
		map1.put("d1321076-c6ad-4355-930a-38cc3a780480","学生会外联部");
		map1.put("a30ce1ea-121c-48c7-a04b-50426c0b395d","学生会国际部");
		//map2
		map2.put("7e3aaca9-a7cc-4ca8-a746-13ae6fda215a","团委办公室");
		map2.put("e400dda8-1fbe-4270-8463-7fe35ccfb26a","团委组织部");
		map2.put("ba49951f-1750-4bd2-b84f-34273ad04995","团委宣传部");
		map2.put("9ad4b90f-eabd-4d0a-8928-d12824b1248a","团委信息技术中心");
		map2.put("00999fcc-33a9-446a-887f-35a9e5a087e9","团委志愿者工作部");
		map2.put("321cf6bc-a9ae-44b6-8113-e2ac08d59aee","团委社会实践部");
		map2.put("90a343c7-3eea-4cc8-9b43-64826be1dd31","学生会办公室");
		map2.put("a6cf3973-1b64-4390-8d8c-5bdcc5aa5959","学生会纪检部");
		map2.put("4b6bd920-891d-4328-b7a2-221b6a3578cc","学生会学术部");
		map2.put("4b143a32-48ca-453e-b10c-edd580c0a76f","学生会职业发展部");
		map2.put("3c3d4fe7-ae29-47b0-b433-d74b704b8dcd","学生会文体部");
		map2.put("1d28c999-97b7-468d-bd69-d67629714629","学生会生活权益部");
		map2.put("e0806ed5-7c5a-46c3-b6a1-507570f68103","学生会女生部");
		map2.put("852a77ab-3b0c-4243-adcb-7f6a20b3d6f4","学生会外联部");
		map2.put("a203367d-fc5d-4955-8ed9-579e261abd53","学生会国际部");
		//map3
		map3.put("5b2cf251-9d97-4ea7-b34c-45f0e9622712","团委办公室");
		map3.put("d6558aa7-c95e-4dc4-a5c3-57b83de4a77a","团委组织部");
		map3.put("8814e076-9c7b-4c75-9d58-376b958f3ebe","团委宣传部");
		map3.put("5f25f43f-b031-4165-8bda-e2e2b5b7f9b7","团委信息技术中心");
		map3.put("94bbbd42-7ceb-43fb-ad0f-13c66d388dec","团委志愿者工作部");
		map3.put("ed51c092-9388-4eb3-9c17-d3bc48e6ae0c","团委社会实践部");
		map3.put("a76cc7a0-b3ea-4265-a6aa-76102d8a18f7","学生会办公室");
		map3.put("222cd6bb-1269-464d-9d77-6a2e2deb16b5","学生会纪检部");
		map3.put("41be0f97-e771-43db-9a8c-a58216c5aa2b","学生会学术部");
		map3.put("e42722d5-1fd7-4d29-9a79-def8a66062ba","学生会职业发展部");
		map3.put("17f74d9e-06fd-4c50-86a3-cdef1e2d0802","学生会文体部");
		map3.put("66a9adac-14ba-4058-826d-25e782267e0e","学生会生活权益部");
		map3.put("52c46a48-df75-4d48-9931-0c64f4d903d3","学生会女生部");
		map3.put("82d0b964-f719-488f-b6d1-748509e2c3d1","学生会外联部");
		map3.put("4bbeb1c9-11b8-40fa-ae72-74a0dfb06983","学生会国际部");
		//mapSex
		mapSex.put("29cf5eec-893a-4414-b889-14d2e8a10b37", "男");
		mapSex.put("f332cba5-61bd-4b07-8685-1bb3ed07df45", "女");
		//mapChange
		mapChange.put("ecbfd7d9-4edc-4551-a234-c8f5d0dc11eb","是");
		mapChange.put("909769cd-6a9d-4f79-a7bd-e7421ff6e732","否");
		List<String> optionsRes = new ArrayList<String>();
		String optionOne = (String) map1.get(one);
		String optionTwo = (String) map2.get(two);
		String optionThree = (String) map3.get(three);
		String optionSex = (String) mapSex.get(sex);
		String optionChange = (String) mapChange.get(change);
		
		optionsRes.add(optionOne);
		optionsRes.add(optionTwo);
		optionsRes.add(optionThree);
		optionsRes.add(optionSex);
		optionsRes.add(optionChange);
		
		return optionsRes;
	}
}
