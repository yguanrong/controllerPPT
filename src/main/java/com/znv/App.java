package com.znv;

import com.znv.utils.JacobPptUtil;
import com.znv.utils.PPTFileFilter;
import com.znv.utils.WebSocketClientTest;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.TimerTask;


/**
 * Hello world!
 *
 */
public class App implements ActionListener{

    private static final long serialVersionUID = 1L;

    JButton btn = null;
    JButton open = null;
    JTextField textField = null;
    private static JFrame jFrame ;

    private App()
    {
        jFrame = new JFrame();
        jFrame.setTitle("选择ppt窗口");
        FlowLayout layout = new FlowLayout();// 布局
        JLabel label = new JLabel("请选择要控制的ppt：");// 标签
        textField = new JTextField(20);// 文本域
        btn = new JButton("浏览");// 钮1
        open = new JButton("打开");

        // 设置布局
        layout.setAlignment(FlowLayout.LEFT);// 左对齐
        jFrame.setLayout(layout);
        jFrame.setBounds(450, 240, 500, 80);
        jFrame.setVisible(true);
        jFrame.setResizable(false);
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        btn.addActionListener(this);
        jFrame.add(label);
        jFrame.add(textField);
        jFrame.add(btn);
        jFrame.add(open);
        open.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!"".equals(textField.getText())){
                    try {
                        jFrame.setExtendedState(JFrame.ICONIFIED);//最小化窗体
                        // 创建Jacob对象
                        JacobPptUtil jacobPptUtil = new JacobPptUtil(textField.getText(),true);
                        // 与服务器建立连接
                        WebSocketClientTest.subscribeFss(jacobPptUtil);

                    } catch (Exception e1) {
                        JOptionPane.showMessageDialog(jFrame, e1.getMessage(),
                                "警告",JOptionPane.WARNING_MESSAGE);
                    }
                }
                else {

                    JOptionPane.showMessageDialog(jFrame, "文件不能为空",
                            "警告",JOptionPane.WARNING_MESSAGE);
                }
            }
        });
    }

    public TimerTask timerTask(final JacobPptUtil jac){

        TimerTask timerTask = new TimerTask() {
            @Override
            public void run() {
                jac.nextPage();
//                try {
//                    jac.exitPPT();
//                } catch (Exception e) {
//                    e.printStackTrace();
//                }
            }
        };
        return timerTask;
    }

    @Override
    public void actionPerformed(ActionEvent e)
    {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        //拦截ppt
        PPTFileFilter pptFileFilter = new PPTFileFilter();
        chooser.setFileFilter(pptFileFilter);
        chooser.addChoosableFileFilter(pptFileFilter);
        chooser.showDialog(new JLabel(), "选择");
        File file = chooser.getSelectedFile();
        if (file.getName().endsWith(".pptx")){
            textField.setText(file.getAbsoluteFile().toString());
        }
        else {
            textField.setText("");
            JOptionPane.showMessageDialog(jFrame, "文件只能是PPT文件",
                    "警告",JOptionPane.WARNING_MESSAGE);
        }

    }

    public static void Waraning(String string){
        JOptionPane.showMessageDialog( jFrame, string,
                "警告",JOptionPane.WARNING_MESSAGE);
    }

    public static void main(String[] args)
    {
        new App();
    }
}
