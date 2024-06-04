package org.example;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Text;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import java.util.List;

public class Main2 {
    public static void main(String[] args) {
        try {
            // Load the DOCX file
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("src/main/resources/intput2.docx"));

            // Define the placeholders and their replacements
            String placeholder = "PLACEHOLDER";
            String replacement = "ActualValue";

            // Get the main document part
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

            // Get all texts in the document
            List<Object> texts = documentPart.getJAXBNodesViaXPath("//w:t", true);

            // Replace the placeholders
            for (Object obj : texts) {
                Text text = (Text) obj;
                if (text.getValue().contains(placeholder)) {
                    text.setValue(text.getValue().replace(placeholder, replacement));
                }
            }

            // Save the modified document
            wordMLPackage.save(new java.io.File("src/main/resources/output2.docx"));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
