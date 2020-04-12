package com.znv.utils;

import javax.swing.filechooser.FileFilter;
import java.io.File;


public class PPTFileFilter extends FileFilter {

    @Override
    public String getDescription() {
        return "*.pptx";
    }

    public boolean accept(File file) {
        String name = file.getName();
        return file.isDirectory() || name.toLowerCase().endsWith(".ppt") || name.toLowerCase().endsWith(".pptx");  // 仅显示目录和xls、xlsx文件
    }

}
