package com.glaway.pageofficedemo.page.controller;

import com.glaway.pageofficedemo.util.StringUtils;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.*;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@RestController
@RequestMapping(value = "/pageController")
public class PageController {

    @Value("${pageOffice.posyspath}")
    private String poSysPath;
    @Value("${pageOffice.popassword}")
    private String poPassWord;
    /**
     * 添加PageOffice的服务器端授权程序Servlet（必须）
     * @return
     */
    @Bean
    public ServletRegistrationBean servletRegistrationBean() {
        com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
        //设置PageOffice注册成功后,license.lic文件存放的目录
        poserver.setSysPath(poSysPath);
        ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
        srb.addUrlMappings("/poserver.zz");
        srb.addUrlMappings("/posetup.exe");
        srb.addUrlMappings("/pageoffice.js");
        srb.addUrlMappings("/jquery.min.js");
        srb.addUrlMappings("/pobstyle.css");
        srb.addUrlMappings("/sealsetup.exe");
        return srb;//
    }
    /**
     * 添加印章管理程序Servlet（可选）
     * @return
     */
    @Bean
    public ServletRegistrationBean servletRegistrationBean2() {
        com.zhuozhengsoft.pageoffice.poserver.AdminSeal adminSeal = new com.zhuozhengsoft.pageoffice.poserver.AdminSeal();
        adminSeal.setAdminPassword(poPassWord);//设置印章管理员admin的登录密码
        adminSeal.setSysPath(poSysPath);//设置印章数据库文件poseal.db存放目录
        ServletRegistrationBean srb = new ServletRegistrationBean(adminSeal);
        srb.addUrlMappings("/adminseal.zz");
        srb.addUrlMappings("/sealimage.zz");
        srb.addUrlMappings("/loginseal.zz");
        return srb;//
    }



    @RequestMapping(value = "/index",method = RequestMethod.GET)
    public ModelAndView showIndex(){
        return new ModelAndView("index");
    }

    /**
     * 1.打开Word文件
     * 2.添加保存、打印、全屏切换按钮
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/word",method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
        //--PageOffice的调用代码 开始
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1);//添加自定义按钮
        poCtrl.addCustomToolButton("打印","ShowPrintDialog",6);
        poCtrl.addCustomToolButton("-","",0);
        poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen", 4);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.setSaveFilePage("/pageController/save");//设置保存的action
        poCtrl.setCaption("国睿集成PageOffice测试系统");//设置控件的标题栏内容

//        poCtrl.setTitlebar(false); //隐藏标题栏
//        poCtrl.setMenubar(false); //隐藏菜单栏
//        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
//        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
//        poCtrl.webOpen("d:\\test.docx", OpenModeType.docReadOnly,"张三");//打开文档只读
//        poCtrl.setAllowCopy(false);//禁止拷贝
//        poCtrl.setJsFunction_AfterDocumentOpened("js函数名");//打开文件时执行的函数
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");//打开文件时执行的函数
        //其中WebOpen方法的第一个参数是office文件在服务器端的磁盘路径，在此demo中暂时使用常量：d:\\test.doc
        poCtrl.webOpen("d:\\test.docx", OpenModeType.docAdmin,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        poCtrl.setTagId("PageOfficeCtrl1");
        //--PageOffice的调用代码 结束
        return new ModelAndView("Word");
    }

    /**
     * 保存文件
     * @param request
     * @param response
     */
    @RequestMapping(value = "/save")
    public void saveFile(HttpServletRequest request, HttpServletResponse response){
        System.out.println("2222");
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile("d:\\" + fs.getFileName());
        fs.setCustomSaveResult("ok");//设置保存结果
        fs.close();//必须close
    }

    /**
     * 待开Excel
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/excel",method = RequestMethod.GET)
    public ModelAndView showExcel(HttpServletRequest request, Map<String,Object> map){
        //--PageOffice的调用代码 开始
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1);//添加自定义按钮
        poCtrl.addCustomToolButton("打印","ShowPrintDialog",6);
        poCtrl.addCustomToolButton("-","",0);
        poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen", 4);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.setSaveFilePage("/pageController/save");//设置保存的action
        //其中WebOpen方法的第一个参数是office文件在服务器端的磁盘路径，在此demo中暂时使用常量：d:\\test.doc
        poCtrl.webOpen("d:\\test.xlsx", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--PageOffice的调用代码 结束
        return new ModelAndView("Excel");
    }

    /**
     * 待开PPT
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/ppt",method = RequestMethod.GET)
    public ModelAndView showPPT(HttpServletRequest request, Map<String,Object> map){
        //--PageOffice的调用代码 开始
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1);//添加自定义按钮
        poCtrl.addCustomToolButton("打印","ShowPrintDialog",6);
        poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen", 4);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.setSaveFilePage("/pageController/save");//设置保存的action
        //其中WebOpen方法的第一个参数是office文件在服务器端的磁盘路径，在此demo中暂时使用常量：d:\\test.doc
        poCtrl.webOpen("d:\\test.pptx", OpenModeType.pptNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--PageOffice的调用代码 结束
        return new ModelAndView("ppt");
    }

    @RequestMapping(value = "/createWordTable",method = RequestMethod.GET)
    public ModelAndView createWordTable(HttpServletRequest request, Map<String,Object> map){
        System.out.println("进入》》》》》》》》》createWordTable>>>>" + request.getContextPath() );
        //--PageOffice的调用代码 开始
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);

//        poCtrl.addCustomToolButton("保存","Save",1);//添加自定义按钮
//        poCtrl.addCustomToolButton("打印","ShowPrintDialog",6);
//        poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen", 4);
//        poCtrl.addCustomToolButton("关闭","Close",21);
//        poCtrl.setSaveFilePage("/pageController/save");//设置保存的action
        poCtrl.setCaption("国睿集成PageOffice测试系统");//设置控件的标题栏内容

        WordDocument document = new WordDocument();
        //在Word中指定的“PO_table1”的数据区域内动态创建一个3行5列的表格
        Table table1 = document.openDataRegion("PO_table1").createTable(3,5, WdAutoFitBehavior.wdAutoFitWindow);
        //合并(1,1)到(3,1)的单元格并赋值
        table1.openCellRC(1, 1).mergeTo(3, 1);
        table1.openCellRC(1, 1).setValue("合并后的单元格");
        //给表格table1中剩余的单元格赋值
        for (int i = 1; i < 4; i++) {
            table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
            table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
            table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
            table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
        }

        //在"PO_table1"后面动态创建一个新的数据区域"PO_table2",用于创建新的一个5行5列的表格table2
        DataRegion drTable2 = document.createDataRegion("PO_table2", DataRegionInsertType.After, "PO_table1");
        Table table2 = drTable2.createTable(5, 5, WdAutoFitBehavior.wdAutoFitWindow);
        //给新表格table2赋值
        for (int i = 1; i < 6; i++) {
            table2.openCellRC(i, 1).setValue("AA" + String.valueOf(i));
            table2.openCellRC(i, 2).setValue("BB" + String.valueOf(i));
            table2.openCellRC(i, 3).setValue("CC" + String.valueOf(i));
            table2.openCellRC(i, 4).setValue("DD" + String.valueOf(i));
            table2.openCellRC(i, 5).setValue("EE" + String.valueOf(i));
        }
        poCtrl.setWriter(document);//此行必须
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet(必须)
        //其中WebOpen方法的第一个参数是office文件在服务器端的磁盘路径，在此demo中暂时使用常量：d:\\test.doc
        poCtrl.webOpen("E:\\pageOfficeDocument\\createTable.docx", OpenModeType.docNormalEdit,"张三");
        System.out.println(poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("pageoffice1",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--PageOffice的调用代码 结束
        return new ModelAndView("CreateWordTable");
    }

    @RequestMapping(value = "/setExcelValue",method = RequestMethod.GET)
    public ModelAndView setExcelValue(HttpServletRequest request,Map<String,Object> map){

        //设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
        poCtrl1.setServerPage("/poserver.zz");
        poCtrl1.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workbook = new Workbook();
        //定义Sheet对象，‘Sheet1’是打开的Excel表单的名称
        Sheet sheet = workbook.openSheet("Sheet1");
        //定义Cell对象
        for (int i = 4;i < 14;i++){
            Cell B = sheet.openCell("B"+i);
            //给单元格赋值
            B.setValue(i-3+"月");
            Cell C = sheet.openCell("C" + i);
            C.setValue("300" + i*10);
            Cell D = sheet.openCell("D" + i);
            D.setValue("270" + i*10);
            Cell E = sheet.openCell("E" + i);
            E.setValue("270" + i*10);
            Cell F = sheet.openCell("F" + i);
            DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
            F.setValue(df.format((270.00 + i*10) / (300+i*10) * 100) + "%");
        }
        poCtrl1.setWriter(workbook);
        //隐藏菜单栏
        poCtrl1.setMenubar(false);
        //隐藏工具栏
        poCtrl1.setCustomToolbar(false);
        //打开Word文件
        poCtrl1.webOpen("E:\\pageOfficeDocument\\setExcelValue.xlsx", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice",poCtrl1.getHtmlCode("PageOfficeCtrl1"));
        return new ModelAndView("setExcelValue");
    }

    /**
     * 1.打开Word文件
     * 2.添加保存、打印、全屏切换按钮
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/submitWord",method = RequestMethod.GET)
    public ModelAndView submitWord(HttpServletRequest request, Map<String,Object> map){
        //--PageOffice的调用代码 开始
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1);//添加自定义按钮
        poCtrl.addCustomToolButton("打印","ShowPrintDialog",6);
        poCtrl.addCustomToolButton("-","",0);
        poCtrl.addCustomToolButton("全屏切换", "SwitchFullScreen", 4);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.setSaveDataPage("/pageController/saveSubmitWord");//设置保存的action
        poCtrl.setCaption("国睿集成PageOffice测试系统");//设置控件的标题栏内容

        WordDocument document = new WordDocument();
        DataRegion dataRegion1 = document.openDataRegion("PO_userName");
        dataRegion1.setEditing(true);//数据区域可编辑
        dataRegion1.setValue("张三");//设置初始值

        DataRegion dataRegion2 = document.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);//数据区域可编辑
        dataRegion2.setValue("业务三部");//设置初始值

        DataRegion dataRegion3 = document.openDataRegion("PO_userAge");
        dataRegion3.setEditing(true);//数据区域可编辑
        dataRegion3.setValue("15");//设置初始值

        poCtrl.setWriter(document);

        //其中WebOpen方法的第一个参数是office文件在服务器端的磁盘路径，在此demo中暂时使用常量：d:\\test.doc
        poCtrl.webOpen("E:\\pageOfficeDocument\\test.docx", OpenModeType.docSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--PageOffice的调用代码 结束
        return new ModelAndView("submitWord");
    }

    /**
     * 保存文件
     * @param request
     * @param response
     */
    @RequestMapping(value = "/saveSubmitWord")
    public static ModelAndView saveSubmitWord(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) throws IOException {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request,response);
        //获取提交的数值
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion userName = doc.openDataRegion("PO_userName");
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion deptName = doc.openDataRegion("PO_deptName");
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion userAge = doc.openDataRegion("PO_userAge");
        System.out.println("userName:" + userName.getValue());
        System.out.println("deptName:" + deptName.getValue());
        System.out.println("userAge:" + userAge.getValue());
        String content = "";
        content +="用户名：" + userName.getValue();
        content +="<br/>年龄：" + userAge.getValue();
        content +="<br/>部门：" + deptName.getValue();
        doc.showPage(500,400);
        doc.close();
        response.setCharacterEncoding("UTF-8");
        request.setAttribute("content",new String(content.getBytes(), "utf-8"));
        System.out.println(new String(content.getBytes(), "utf-8"));
    //    map.put("content",new String(content.getBytes("GBK"), "utf-8"));
        return new ModelAndView("submitWordDialog");
    }

    /**
     * 提交Excel数据
     * @param request
     * @return
     */
    @RequestMapping(value = "/submitExcel",method = RequestMethod.GET)
    public ModelAndView submitExcel(HttpServletRequest request){

        PageOfficeCtrl ptrCtrl = new PageOfficeCtrl(request);
        ptrCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        ptrCtrl.setCaption("集成pageOffice测试系统");
        ptrCtrl.setTagId("PageOfficeCtrl1");
        ptrCtrl.addCustomToolButton("保存","save",1);
        ptrCtrl.setSaveDataPage("/pageController/saveSubmitExcel");//设置保存的action
        //定义Workbook对象
        Workbook workbook = new Workbook();
        //定义Sheet对象，“sheet1”是打开Excel表单的名称
        Sheet sheet = workbook.openSheet("Sheet1");
        //定义table对象，设置table的设置范围
        com.zhuozhengsoft.pageoffice.excelwriter.Table table = sheet.openTable("B4:F13");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");
        ptrCtrl.setWriter(workbook);
        ptrCtrl.webOpen("E:\\pageOfficeDocument\\setExcelValue.xlsx",OpenModeType.xlsSubmitForm,"张三");
        request.setAttribute("pageOffice",ptrCtrl.getHtmlCode("PageOfficeCtrl1"));
        return new ModelAndView("submitExcel");
    }

    /**
     * 保存提交excel中的数据
     * @param request
     * @param response
     * @return
     */
    @RequestMapping(value = "/saveSubmitExcel")
    public  ModelAndView saveSubmitExcel(HttpServletRequest request,HttpServletResponse response){

        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request,response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        com.zhuozhengsoft.pageoffice.excelreader.Table table = sheet.openTable("Info");
        String content = "";
        int result = 0;
        while (!table.getEOF()){//获取 Table 的当前数据行是否指向结束
            //获取提交的数值
            if(!table.getDataFields().getIsEmpty()){
                content += "<br/>月份名称："
                    + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                    + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                    + table.getDataFields().get(2).getText();
                content += "<br/>累积完成量："
                    + table.getDataFields().get(3).getText();

                if (table.getDataFields().get(2).getText().equals(null)
                    || table.getDataFields().get(2).getText().trim().length() == 0){
                    content += "<br/>完成率：0%";
                }else {
                    float f = Float.parseFloat(table.getDataFields().get(2).getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "<br/>******************************";

            }
            //循环进入下一行
            table.nextRow();
        }
        System.out.println("***************<br/>" + content);
        System.out.println("******************************");
        table.close();
        workBook.setCustomSaveResult("OK");
        workBook.showPage(500,400);
        workBook.close();
        request.setAttribute("content",new String(content.getBytes(),StandardCharsets.UTF_8));
        return new ModelAndView("submitExcelDialog");
    }

    /**
     *
     * @param request
     * @return
     */
    @RequestMapping(value = "/jsControlBars",method = RequestMethod.GET)
    public  ModelAndView jsControlBars(HttpServletRequest request){

        PageOfficeCtrl ctrl = new PageOfficeCtrl(request);
        ctrl.setServerPage("/poserver.zz");//设置授权程序servlet
        ctrl.webOpen("E:\\pageOfficeDocument\\test.docx",OpenModeType.docNormalEdit,"张三");
        request.setAttribute("pageoffice",ctrl.getHtmlCode("PageOfficeCtrl1"));
        return new ModelAndView("Word");
    }

    @RequestMapping(value = "/concurrencyCtrl",method = RequestMethod.GET)
    public ModelAndView concurrencyCtrl(HttpServletRequest request){
        return new ModelAndView("testConcurrency");
    }
    @RequestMapping(value = "/concurrencyWord",method = RequestMethod.POST)
    public void concurrencyWord(HttpServletRequest request){

        String id = request.getParameter("id");
        System.out.println("id:" + id);
        String userName = "无";
        if(StringUtils.isNotEmpty(id)){
            if (id.equals("1")){
                userName = "张三";
            }else if(id.equals("2")){
                userName = "李四";
            }
        }

        PageOfficeCtrl ctrl = new PageOfficeCtrl(request);
        ctrl.setServerPage("/poserver.zz");
        ctrl.addCustomToolButton("保存","save",1);
        ctrl.setSaveFilePage("/pageController/save");

        //设置并发控制时间(单位：分钟)
        ctrl.setTimeSlice(20);
        ctrl.webOpen("",OpenModeType.docRevisionOnly,userName);
        request.setAttribute("username",userName);
  //      return new ModelAndView("concurrencyWord");
    }
}
