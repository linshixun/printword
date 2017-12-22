package com.tonylinzhen.wordtemplate;

import java.io.FileInputStream;
import java.io.InputStream;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import org.apache.poi.xwpf.usermodel.*;

public class WordTemplateUtils {

    public static void main(String[] args) {
        List<Map<String, String>> maps = csv2Map("fff.csv", ",");
        System.out.println(maps);
        for (Map<String, String> map : maps) {
            createDocWithTemplate("法律文书_孙律师最终修正版.doc", map , "F:\\hehe\\"+map.get("FULL_NAME")+".doc");
        }

    }

    /**
     * 提供csv内容隐射为map
     *
     * @param filePath
     * @param splitBy
     * @return
     */
    public static List<Map<String, String>> csv2Map(String filePath, String splitBy) {


        List<Map<String, String>> list = new ArrayList<Map<String, String>>();
        try {
            Scanner scanner = new Scanner(new File(filePath));
            scanner.useDelimiter("[\r\n]");
            String[] header = new String[0];
            if (scanner.hasNext()) {
                String next = scanner.next();
                header = next.split(splitBy);

            }
            while (scanner.hasNext()) {
                String[] split = scanner.next().split(splitBy);
                HashMap<String, String> map = new HashMap<String, String>();
                for (int i = 0; i < header.length; i++) {

                    map.put(header[i], split[i]);
                }

                list.add(map);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        return list;
    }


    /**
     * 根据word模板和隐射数据,输出doc文档
     *
     * @param templateFile
     * @param row
     * @param outputFile
     * @return
     */
    public static void createDocWithTemplate(String templateFile, Map<String, String> row, String outputFile) {
        InputStream istream = null;

        OutputStream ostream = null;
        try {
            istream = new FileInputStream(templateFile);

            XWPFDocument document = new XWPFDocument(istream);
            Iterator<IBodyElement> bodyElementsIterator = document.getBodyElementsIterator();
            for (; bodyElementsIterator.hasNext(); ) {
                XWPFParagraph next = (XWPFParagraph) bodyElementsIterator.next();

                List<XWPFRun> runs = next.getRuns();
                Iterator<XWPFRun> iterator = runs.iterator();

                for (; iterator.hasNext(); ) {
                    XWPFRun next1 = iterator.next();

                    String text = next1.text();
                    for (Map.Entry<String, String> o : row.entrySet()) {
                        if (text.contains(o.getKey())) {
                            text = text.replace("${" + o.getKey() + "}", o.getValue());
                        }
                    }
                    next1.setText(text, 0);

                    System.out.println(text);

                }


            }


            ostream = new FileOutputStream(outputFile, false);
            document.write(ostream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (istream != null) {

                try {
                    ostream.close();
                    istream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }
}