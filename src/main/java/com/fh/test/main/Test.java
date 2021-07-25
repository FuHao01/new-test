package com.fh.test.main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import javax.swing.*;

public class Test {
    public static void main(String[] args) {
        //设置panel的layout以及sieze
        JPanel jpanel = new JPanel();
        jpanel.setLayout(null);
        jpanel.setPreferredSize(new Dimension(680, 240));

        //添加输入框
        final JTextField text = new JTextField(20);
        text.setBounds(20, 20, 500, 30);
        jpanel.add(text);

        final JTextField text2 = new JTextField(20);
        text2.setBounds(20, 60, 500, 30);
        jpanel.add(text2);

        //添加按钮
        JButton button1 = new JButton("选择目标路径");
        button1.setBounds(530, 20, 120, 30);
        jpanel.add(button1);

        JButton button2 = new JButton("选择输出位置");
        jpanel.add(button2);
        button2.setBounds(530, 60, 120, 30);

        JButton button3 = new JButton("输出目标下所有文件夹");
        jpanel.add(button3);
        button3.setBounds(20, 100, 180, 30);

        JButton button4 = new JButton("将目标下所有空文件夹移动到输出位置");
        jpanel.add(button4);
        button4.setBounds(210, 100, 280, 30);


        // 设置窗体属性
        JFrame frame = new JFrame("付氏小程序");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.add(jpanel);
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

        //添加事件
        button1.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int option = fileChooser.showOpenDialog(frame);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                text.setText(file.getAbsolutePath());
            }
        });

        button2.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int option = fileChooser.showOpenDialog(frame);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                text2.setText(file.getAbsolutePath());
            }
        });

        //输出目标下所有文件夹 xlsx
        button3.addActionListener(e -> {
            if (check(jpanel, text, text2)) {
                out(jpanel, text, text2);
            }
        });

        //将目标下所有空文件夹移动到输出位置
        button4.addActionListener(e -> {
            if (check(jpanel, text, text2)) {
                out2(jpanel, text, text2);
            }
        });

    }

    static void out2(JPanel jpanel, JTextField text, JTextField text2) {
        String path = text.getText();
        File file1 = new File(path);
        File[] files = file1.listFiles();
        String path2 = text2.getText();
        for (int i = 0; i < files.length; i++) {
            File file = files[i];
            if (file.isDirectory() && (file.listFiles() == null || file.listFiles().length < 1)) {
                file.renameTo(new File(path2 + File.separator + file.getName()));
            }
        }
        JOptionPane.showMessageDialog(jpanel, "操作成功", "提示", JOptionPane.WARNING_MESSAGE);
    }

    static void out(JPanel jpanel, JTextField text, JTextField text2) {
        String path = text.getText();
        File file1 = new File(path);
        File[] files = file1.listFiles();
        String path2 = text2.getText();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("sheetName");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN); //下边框
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);//左边框
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);//上边框
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);//右边框
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER); // 居中

        int index = 0;
        for (int i = 0; i < files.length; i++) {
            File file = files[i];
            if (file == null || file.isFile()) continue;
            Row row = sheet1.createRow(index);
            Cell cell = row.createCell(0);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(file.getName());
            cell = row.createCell(1);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(file.listFiles() != null && file.listFiles().length > 0 ? "" : "空文件夹");
            index++;
        }

        try {
            wb.write(new FileOutputStream(new File(path2 + File.separator + "文件夹输出"
                    + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + ".xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(jpanel, "操作失败", "提示", JOptionPane.WARNING_MESSAGE);
            return;
        }
        JOptionPane.showMessageDialog(jpanel, "操作成功", "提示", JOptionPane.WARNING_MESSAGE);

    }

    /**
     * 校验路径
     */
    static boolean check(JPanel jpanel, JTextField text, JTextField text2) {
        String path = text.getText();
        if (isBlank(path)) {
            JOptionPane.showMessageDialog(jpanel, "请选择目标路径", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        String path2 = text2.getText();
        if (isBlank(path2)) {
            JOptionPane.showMessageDialog(jpanel, "请选择输出路径", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        File file1 = new File(path);
        if (file1.isFile()) {
            JOptionPane.showMessageDialog(jpanel, "目标路径不是文件夹", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        if (!file1.exists()) {
            JOptionPane.showMessageDialog(jpanel, "目标路径不存在", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        File file2 = new File(path);
        if (file2.isFile()) {
            JOptionPane.showMessageDialog(jpanel, "输出路径不是文件夹", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        if (!file2.exists()) {
            JOptionPane.showMessageDialog(jpanel, "输出路径不存在", "提示", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        return true;
    }

    static boolean isNotBlank(String string) {
        return string != null && !"".equals(string) && !"null".equals(string);
    }

    static boolean isBlank(String string) {
        return !isNotBlank(string);
    }
}
