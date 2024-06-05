package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    

    public static void replaceParametersInDocx(String inputFilePath, String outputFilePath, Map<String, String> parameters) throws XmlException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {


            // Iterate over paragraphs in the document
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell tableCell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                            XmlCursor cursor = paragraph.getCTP().newCursor();
                            cursor.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:drawing/*/w:txbxContent/w:p/w:r");
                            List<XmlObject> ctrsintxtbx = new ArrayList<XmlObject>();

                            while(cursor.hasNextSelection()) {
                                cursor.toNextSelection();
                                XmlObject obj = cursor.getObject();
                                ctrsintxtbx.add(obj);
                            }
                            for (XmlObject obj : ctrsintxtbx) {
//                                System.out.println("-----------");
                                CTR ctr = CTR.Factory.parse(obj.xmlText());
                                //CTR ctr = CTR.Factory.parse(obj.newInputStream());
                                XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
                                String text = bufferrun.getText(0);
//                                System.out.println(text);
                                if (text != null) {
                                    for (Map.Entry<String, String> entry : parameters.entrySet()) {
                                        if (text.contains(entry.getKey())) {
                                            text = text.replace(entry.getKey(), entry.getValue());
                                            bufferrun.setText(text, 0);
                                        }
                                    }
//                                    text = text.replace("{name}", "replaced");
//                                    bufferrun.setText(text, 0);
                                }
                                obj.set(bufferrun.getCTR());
                            }
                        }
                    }
                }
            }
            // Save the updated document to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                document.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    private static void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> parameters) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs != null) {
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : parameters.entrySet()) {
                        if (text.contains(entry.getKey())) {
                            text = text.replace(entry.getKey(), entry.getValue());
                            run.setText(text, 0);
                        }
                    }
                }
            }
        }
    }

    public static void main(String[] args) throws IOException, InvalidFormatException, XmlException {
//         Define the input and output file paths
        String inputFilePath = "src/main/resources/input3.docx";
        String outputFilePath = "src/main/resources/output.docx";

        // Define the parameters to replace in the document
        Map<String, String> parameters = new HashMap<>();
        parameters.put("{name}", "Nguyễn Cảnh Ba Đình Cầu Giấy");
        parameters.put("{nationality}", "Vietnam");
        parameters.put("{idNo}", "999999999999");
        parameters.put("{idDate}", "30/06/2024");
        parameters.put("{idPlace}", "Quảng Ngãi");
        parameters.put("{birth}", "30/06/2020");
        parameters.put("{ma}", "X");
        parameters.put("{fm}", "X");
        parameters.put("{mobile}", "0999444666");
        parameters.put("{email}", "nguyenconghoan@outlook.com");
        parameters.put("{address}", "SN 01, ngõ 4, xóm dinh, thôn 2, xã quảng bị, huyện Chương Mỹ TP Hà Nội");
        parameters.put("{address2}", "SN 01, ngõ 4, xóm dinh, thôn 2, xã quảng bị, huyện Chương Mỹ TP Hà Nội");

        parameters.put("{userBank}", "Nguyen Canh Ba Dinh Cau Giay");
        parameters.put("{userAcc}", "1903555666777999222");
        parameters.put("{bankBranch}", "Ho Chi Minh");
        parameters.put("{userBankName}", "Ngan Hang Co Phan Thuong Mai Ky Thuong Viet Nam Techcombank");

        // Replace parameters in the DOCX file
        replaceParametersInDocx(inputFilePath, outputFilePath, parameters);
        
//        testt();

        System.out.println("Parameters replaced successfully.");
    }
}