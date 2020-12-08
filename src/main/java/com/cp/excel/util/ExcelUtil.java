package com.cp.excel.util;

import com.cp.excel.handle.ExcelHandle;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/**
 * @author cuipeng
 * @Create 2020-12-03-3:34 下午
 **/
public class ExcelUtil {

    private static String version = "2007";


    private static final String ZERO = "0.000";

    /**
    * 复制模板 重新命名
    *
    * @author cuipeng
    * @Create 2020/12/3 3:40 下午
    **/
    public static String getCopyTemplate(MultipartFile file, String name, List<Integer> addColList) throws Exception {
        //文件夹地址
        String fileDir = System.getProperty("user.dir");
        //模板名称
        String templateName= name+"."+FileUtil.getFileType(file);
        //模板地址加名称
        String templatePathName = fileDir+File.separator+templateName;
        //创建模板
        File Template = new File(templatePathName);
        FileUtil.CopyFile(file,Template);
        //初始化数据增加总数统计
        initTemplateData(Template,addColList,FileUtil.trimType(file.getOriginalFilename()));
        return templatePathName ;
    }

    public static List<Integer> getAllColsAddOne(List<Integer> allcols){
        List<Integer> newList = new ArrayList<>();
        for (Integer allcol : allcols) {
            newList.add(allcol+1);
        }
        return newList;
    }

    /**
    * 处理添加新增模板数据
    *
    * @author cuipeng
    * @Create 2020/12/4 2:08 下午
    **/
    public static void addTemplateData(File Template,MultipartFile addFile,List<Integer> compareCols,List<Integer> addCols,List<Integer> allcols) throws Exception {
        //添加统计总数列
        ExcelHandle templateExcelHandle = new ExcelHandle(new FileInputStream(Template),version);
        ExcelHandle addExcelHandle = new ExcelHandle(addFile.getInputStream(),version);
        int sheetSize = templateExcelHandle.getSheetCount();
        for(int i=0;i<sheetSize;i++) {
            //两个excel数据
            List<List<String>> templateList = templateExcelHandle.read(i);
            List<List<String>> addFileList = addExcelHandle.read(i);
            addTemaplteNewData(templateList, addFileList,compareCols,addCols,allcols,i,FileUtil.trimType(addFile.getOriginalFilename()));
            templateExcelHandle.clearSheet(i);
            templateExcelHandle.write(i,templateList,0);
            templateExcelHandle.saveExcel(Template);
        }
    }

    /**
     * 处理添加新增模板数据
     *
     * @author cuipeng
     * @Create 2020/12/4 2:08 下午
     **/
    public static void clearUpTemplate(File Template) throws Exception {
        //添加统计总数列
        ExcelHandle templateExcelHandle = new ExcelHandle(new FileInputStream(Template),version);
        int sheetSize = templateExcelHandle.getSheetCount();
        for(int i=0;i<sheetSize;i++) {
            //两个excel数据
            List<List<String>> templateList = templateExcelHandle.read(i);
            clearRow(templateList);
            templateExcelHandle.clearSheet(i);
            templateExcelHandle.write(i,templateList,0);
            templateExcelHandle.saveExcel(Template);
        }
    }

    private static void  initTemplateData(File Template, List<Integer> addColList,String sheetName) throws Exception {
        //添加统计总数列
        ExcelHandle templateExcelHandle = new ExcelHandle(new FileInputStream(Template),version);
        int sheetSize = templateExcelHandle.getSheetCount();
        for(int i=0;i<sheetSize;i++) {
            List<List<String>> templateList = templateExcelHandle.read(i);
            int sheetAddCol = addColList.get(i)-1;
            for (int rownum = 0; rownum < templateList.size(); rownum++) {
                List<String> colString = templateList.get(rownum);
                String s = "总";
                if(rownum>0) {
                    s = colString.get(sheetAddCol);
                }else{
                    colString.set(sheetAddCol,sheetName+colString.get(sheetAddCol));
                }
                if(sheetAddCol>=colString.size()){
                    colString.add(s);
                }else {
                    colString.add(sheetAddCol+1,getAccumulation(s,"0"));
                }
                colString.set(sheetAddCol,getAccumulation(colString.get(sheetAddCol),"0"));
                templateList.set(rownum,colString);
            }
            templateExcelHandle.clearSheet(i);
            templateExcelHandle.write(i,templateList,0);
            templateExcelHandle.saveExcel(Template);
        }
    }

    private static void clearRow(List<List<String>> templateList){
        for (int rownum = 0; rownum < templateList.size(); rownum++) {
            List<String> templateColString = templateList.get(rownum);
            if(templateColString.get(0).equals(templateColString.get(1))){
                for (int i=1;i<templateColString.size();i++) {
                    templateColString.set(i," ");
                }
            }
        }
    }


    private static void addTemaplteNewData(List<List<String>> templateList, List<List<String>> addFileList,
                                           List<Integer> compareCols, List<Integer> addCols, List<Integer> allcols,
                                           int sheetNum, String sheetName){
        int sheetAddCol = addCols.get(sheetNum)-1;
        int compareCol = compareCols.get(sheetNum)-1;
        int allCol = allcols.get(sheetNum);
        for (int rownum = 0; rownum < templateList.size(); rownum++) {
            List<String> templateColString = templateList.get(rownum);
            //如果不是第一行循环找到对应该行的数据
            List<String> addColString = null;
            if(rownum>0) {
                addColString =  getAddRowData(addFileList,templateColString.get(compareCol),compareCol);
            }else{
                addColString = addFileList.get(rownum);
            }
            String s = "";
            //如果没有找到数据  下次循环
            if(ObjectUtils.isEmpty(addColString)){
                //在原来的总数前插入
               s = ZERO;
            }else{
                s = addColString.get(sheetAddCol);
            }
            //找到了  插入并更新 总数
            if(rownum==0) {
                templateColString.add(allCol,templateColString.get(allCol-1));
                templateColString.set(allCol-1,sheetName+s);
            }else{
                //在原来的总数前插入
                templateColString.add(allCol-1,getAccumulation(s,"0"));
                //更新总数
                templateColString.set(allCol,getAccumulation(templateColString.get(allCol-1),templateColString.get(allCol)));
            }
            templateList.set(rownum,templateColString);
        }

        //判断要插入的 数中 是否有 模板中没有的 对比列
        handleTemplateNull(templateList,addFileList,compareCol,sheetAddCol,allCol);
    }

    private static void handleTemplateNull(List<List<String>> templateList, List<List<String>> addFileList,int compareCol,int sheetAddCol, int allCol){
        for (int addN=0;addN<addFileList.size();addN++){
            boolean flag =true;
            for(int temN=0;temN<templateList.size();temN++){
                if(addFileList.get(addN).get(compareCol).equals(templateList.get(temN).get(compareCol))){
                    flag =false;
                    break;
                }
            }
            //新增不包含的行
            if(flag){
                //查找要插入的模板的行数
                for(int temN=0;temN<templateList.size();temN++){
                    if(addFileList.get(addN-1).get(compareCol).equals(templateList.get(temN).get(compareCol))){
                        //处理当前要添加行数据
                        List<String> newList = handleNewList(addFileList.get(addN),sheetAddCol,allCol);
                        //找到跟上一行相同的行  插入
                        if(temN==templateList.size()-1){
                            templateList.add(newList);
                        }else{
                            templateList.add(temN+1,newList);
                        }
                        break;
                    }
                }
            }
        }
    }


    private static List<String> handleNewList(List<String> list,int valNum,int allnum){
        List<String> newList = new ArrayList<>();
        newList.addAll(list);
        if(newList.size()<=allnum){
            for(int num=newList.size();num<=allnum;num++){
                newList.add(ZERO);
            }
        }
        newList.set(allnum-1,getAccumulation(list.get(valNum),"0"));
        newList.set(allnum,getAccumulation(list.get(valNum),"0"));
        newList.set(valNum,ZERO);
        return newList;
    }

    /**
    * 获取template 对应 value 对应的row数据
    *
    * @param addFileList  新增的template 该sheet 全部数据
     * @param comparecol  对比行
     * @param value         对比值
     * @author cuipeng
    * @Create 2020/12/4 2:20 下午
    **/
    private static List<String> getAddRowData(List<List<String>> addFileList, String value, int comparecol){
        for(int i =0 ;i<addFileList.size();i++){
            if(addFileList.get(i).get(comparecol).equals(value)){
                return addFileList.get(i);
            }
        }
        return null;
    }

    /**
    * 计算数字 保留3位小数
    *
    * @author cuipeng
    * @Create 2020/12/4 3:03 下午
    **/
    private static String getAccumulation(String a, String b){
        try{
            if(StringUtils.isEmpty(a)){
                a = "0";
            }
            if(StringUtils.isEmpty(b)){
                b = "0";
            }
            BigDecimal bigDecimal1 = new BigDecimal(a);
            bigDecimal1.setScale(3, RoundingMode.HALF_UP);
            BigDecimal bigDecimal2 = new BigDecimal(b);
            bigDecimal2.setScale(3, RoundingMode.HALF_UP);
            return String.format("%.3f", bigDecimal1.add(bigDecimal2));
        }catch (NumberFormatException e){
            //正常是空返回 0 不是就是汉字 返回汉字
            if(StringUtils.isEmpty(a)){
                return ZERO;
            }
        }
       return a;
    }
}
