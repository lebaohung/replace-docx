package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Main {
    

    public static void replaceParametersInDocx(String inputFilePath, String outputFilePath, Map<String, String> parameters) {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Iterate over paragraphs in the document
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                replaceInParagraph(paragraph, parameters);
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
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text != null) {
                for (Map.Entry<String, String> entry : parameters.entrySet()) {
                    text = text.replace(entry.getKey(), entry.getValue());
                }
                run.setText(text, 0);
            }
        }
    }

    public static void main(String[] args) {
        // Define the input and output file paths
        String inputFilePath = "src/main/resources/intput.docx";
        String outputFilePath = "src/main/resources/output.docx";

        // Define the parameters to replace in the document
        Map<String, String> parameters = new HashMap<>();
        parameters.put("{name}", "John Doe");
        parameters.put("{nationality}", "Vietnam");

        // Replace parameters in the DOCX file
        replaceParametersInDocx(inputFilePath, outputFilePath, parameters);

        System.out.println("Parameters replaced successfully.");
    }
}