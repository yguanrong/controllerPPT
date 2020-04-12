package com.znv.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.File;

public class JacobPptUtil {

    public static final int WORD_HTML = 8;

    public static final int WORD_TXT = 7;

    public static final int EXCEL_HTML = 44;

    public static final int ppSaveAsJPG = 17;

    private static final String ADD_CHART = "AddChart";

    private Boolean pptIsShow = false;

    private ActiveXComponent ppt;
    private ActiveXComponent presentation;

    public Boolean getPptIsShow() {
        return pptIsShow;
    }

    public void setPptIsShow(Boolean pptIsShow) {
        this.pptIsShow = pptIsShow;
    }

    /**
     * 构造一个新的PPT
     * @param isVisble
     */
    public JacobPptUtil(boolean isVisble){
        ppt = new ActiveXComponent("PowerPoint.Application");
        ppt.setProperty("Visible", new Variant(isVisble));
        ActiveXComponent presentations = ppt.getPropertyAsComponent("Presentations");
        presentation =presentations.invokeGetComponent("Add", new Variant(1));
    }
    public JacobPptUtil(String filePath, boolean isVisble)throws Exception{
        if (null == filePath || "".equals(filePath)) {
            throw new Exception("文件路径为空!");
        }
        File file = new File(filePath);
        if (!file.exists()) {
            throw new Exception("文件不存在!");
        }
        ppt = new ActiveXComponent("PowerPoint.Application");
        setIsVisble(ppt, isVisble);
        // 打开一个现有的 Presentation 对象
        ActiveXComponent presentations = ppt.getPropertyAsComponent("Presentations");
        presentation = presentations.invokeGetComponent("Open", new Variant(filePath), new Variant(true));

    }
    /**
     * 将ppt转化为图片
     *
     * @param pptfile
     * @param saveToFolder
     * @author liwx
     */
    public void PPTToJPG(String pptfile, String saveToFolder) {
        try {
            saveAs(presentation, saveToFolder, ppSaveAsJPG);
            if (presentation != null) {
                Dispatch.call(presentation, "Close");
            }
        } catch (Exception e) {
            ComThread.Release();
        } finally {
            if (presentation != null) {
                Dispatch.call(presentation, "Close");
            }
            ppt.invoke("Quit", new Variant[] {});
            ComThread.Release();
        }
    }
    /**
     * 播放ppt
     * @date 2009-7-4
     * @author YHY
     */
    public void PPTShow() {

        if (!pptIsShow && presentation != null){
            // powerpoint幻灯展示设置对象
            ActiveXComponent setting = presentation
                    .getPropertyAsComponent("SlideShowSettings");

            // 调用该对象的run函数实现全屏播放
            setting.invoke("Run");

            // 释放控制线程
            ComThread.Release();
            pptIsShow = true;
        }

    }

    /**
     * ppt另存为
     *
     * @param presentation
     * @param saveTo
     * @param ppSaveAsFileType
     * @date 2009-7-4
     * @author YHY
     */
    public void saveAs(Dispatch presentation, String saveTo,
                       int ppSaveAsFileType)throws Exception {
        Dispatch.call(presentation, "SaveAs", saveTo, new Variant(
                ppSaveAsFileType));
    }
    /**
     * 关闭PPT并释放线程
     * @throws Exception
     */
    public void closePpt()throws Exception{

        if ( pptIsShow && presentation != null) {
            Dispatch.call(presentation, "Close");
            ppt.invoke("Quit", new Variant[]{});
            ComThread.Release();
            pptIsShow = false;
        }
    }

    /**
     * 演示PPT
     * @throws Exception
     */

    public void runPpt()throws Exception{

        if (!pptIsShow && presentation != null) {
            ActiveXComponent setting = presentation.getPropertyAsComponent("SlideShowSettings");
            setting.invoke("Run");
            // 释放控制线程
//            ComThread.Release();
            pptIsShow = true;
        }
    }
    /**
     * 停止演示PPT
     */
    public void stopPpt(){
        if (pptIsShow && presentation != null){
            Robot robot = getRobot();
            System.out.println("退出播放");
            robot.keyPress(KeyEvent.VK_ESCAPE);
            robot.keyRelease(KeyEvent.VK_ESCAPE);
            pptIsShow = false;
        }
    }
    /**
     * 设置是否可见
     * @param visble
     * @param obj
     */
    private void setIsVisble(Dispatch obj,boolean visble)throws Exception{
        Dispatch.put(obj, "Visible", new Variant(visble));
    }
    /**
     * 在幻灯片对象上添加新的幻灯片
     * @param slides
     * @param pptPage 幻灯片编号
     * @param type 4:标题+表格  2:标题+文本 3:标题+左右对比文本 5:标题+左文本右图表 6:标题+左图表右文本 7:标题+SmartArt图形 8:标题+图表
     * @return
     * @throws Exception
     */
    private Variant addPptPage(ActiveXComponent slides,int pptPage,int type)throws Exception{
        return slides.invoke("Add", new Variant(pptPage), new Variant(type));
    }


    /**
     * 获取第几个幻灯片
     * @param pageIndex 序号，从1开始
     * @return
     * @throws Exception
     */

    public Dispatch getPptPage(int pageIndex)throws Exception{
        //获取幻灯片对象
        ActiveXComponent slides = presentation.getPropertyAsComponent("Slides");
        //获得第几个PPT
        Dispatch pptPage = Dispatch.call(slides, "Item", new Object[]{new Variant(pageIndex)}).toDispatch();
        Dispatch.call(pptPage, "Select");
        return pptPage;
    }

    /**
     * 获取控制键盘的Robot
     */
    private Robot getRobot() {
        Robot robot = null;
        try {
            robot = new Robot();
        } catch (AWTException e) {
            e.printStackTrace();
        }
        return robot;
    }

    /**
     * 幻灯片下一页
     */
    public void nextPage(){

        if (pptIsShow && presentation != null){
            Robot robot = getRobot();
            System.out.println("下一页 !!!");
            robot.keyPress(KeyEvent.VK_RIGHT);
            robot.keyRelease(KeyEvent.VK_RIGHT);
        }
    }
    /**
     * 幻灯片上一页
     */
    public void upPage(){

        if (pptIsShow && presentation != null){
            Robot robot = getRobot();;
            System.out.println("上一页 !!!");
            robot.keyPress(KeyEvent.VK_LEFT);
            robot.keyRelease(KeyEvent.VK_LEFT);
        }

    }

    /**
     * 退出
     */
    public void exitPPT() throws Exception {
        if (ppt != null){
            String command = "taskkill /f /im POWERPNT.EXE";
            Runtime.getRuntime().exec(command);
            ppt = null;
            presentation = null;
        }
        pptIsShow = false;
    }


}
