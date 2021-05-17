package com.spring.security.config;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class WordUtil {


   // @Autowired //千万别忘（奥吐wire） 用来查询数据
  //  private UserInfoService userInfoService;
    @Test
    public void test01() throws IOException, XmlException {

        Long start=System.currentTimeMillis();//我是用来记录下总时长 可以忽略
        Map<String, String> map = new LinkedHashMap<>();
       // List<UserInfo> list1 = userInfoService.selectUserInfoByNa("孙尚香");//根据员工姓名查询员工信息

        String srcPath = "E:\\study\\wordExport\\info.docx";//模板路径
        String destPath = "E:\\study\\wordExport\\" +"info"+ System.currentTimeMillis() + ".docx";
        //导出路径，我是为了让文件名不一样加的 System.currentTimeMillis() ，可去掉

        InputStream inputStream = new FileInputStream(srcPath);
        FileOutputStream outputStream = new FileOutputStream(destPath);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

       // for(UserInfo lis:list1){//遍历集合往map中添加数据 由于我定义的map都是string 所以部分做了相应的转换

            map.put("name","tset") ;
            map.put("num","110") ;
            map.put("path","物化大厦") ;
            map.put("date", String.valueOf(sdf)) ;


       // }
        replaceText(inputStream, outputStream, map);//通过此方法来将map中的数据添加到模板中
        inputStream.close();
        outputStream.close();

        Long end=System.currentTimeMillis();
        System.out.println(end-start);

    }
    //此方法中的代码无需改动，本人已调试过。如果不满足您的需求（那你随便玩）
    public static void replaceText(InputStream inputStream, OutputStream outputStream, Map<String, String> map) {
        try {
            XWPFDocument document = new XWPFDocument(inputStream);
            //1. 替换段落中的指定文字（本人模板中 对应的编号）
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            String text;
            Set<String> set;
            XWPFParagraph paragraph;
            List<XWPFRun> run;
            String key;
            while (itPara.hasNext()) {
                paragraph = itPara.next();
                set = map.keySet();
                Iterator<String> iterator = set.iterator();
                while (iterator.hasNext()) {
                    key = iterator.next();
                    run = paragraph.getRuns();
                    for (int i = 0, runSie = run.size(); i < runSie; i++) {
                        text = run.get(i).getText(run.get(i).getTextPosition());
                        if (text != null && text.equals(key)) {
                            run.get(i).setText(map.get(key), 0);
                        }
                    }
                }
            }
            //2. 替换表格中的指定文字（本人模板中 对应的姓名、性别等）
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            XWPFTable table;
            int rowsCount;
            while (itTable.hasNext()) {
                XWPFParagraph p = document.createParagraph();
                XWPFRun headRun0 = p.createRun();
                table = itTable.next();
                rowsCount = table.getNumberOfRows();
                for (int i = 0; i < rowsCount; i++) {
                    XWPFTableRow row = table.getRow(i);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        for (Map.Entry<String, String> e : map.entrySet()) {
                            if (cell.getText().equals(e.getKey())) {
                                cell.removeParagraph(0);
                                // cell.setText(e.getValue());
                                //设置单元格文本样式
                                XWPFParagraph xwpfParagraph = cell.addParagraph();
                                XWPFRun run1 = xwpfParagraph.createRun();
                                run1.setFontSize(11);
                                run1.setText(e.getValue());
                                //设置内容水平居中
                                CTTc cttc = cell.getCTTc();
                                CTTcPr ctPr = cttc.addNewTcPr();
                                ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                                /*  cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);*/
                            }
                        }
                    }
                }
            }
            //3.输出流
            document.write(outputStream);
            System.out.println("shuchule");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

//      <!-- 添加 woerd poi 导出-->
//        <dependency>
//            <groupId>org.apache.poi</groupId>
//            <artifactId>poi-ooxml</artifactId>
//            <version>3.17</version>
//        </dependency>
//        <dependency>
//            <groupId>org.apache.poi</groupId>
//            <artifactId>poi</artifactId>
//            <version>3.17</version>
//        </dependency>
//        <dependency>
//            <groupId>org.junit.jupiter</groupId>
//            <artifactId>junit-jupiter</artifactId>
//            <version>RELEASE</version>
//            <scope>compile</scope>
//        </dependency>


}
