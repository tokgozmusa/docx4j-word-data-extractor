import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.PropertyConfigurator;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;


/**
 * @author tokgozmusa
 */
public class Main {

    /**
     * Main method
     */
    public static void main(String[] args) throws Exception {
        // configure log4j
        PropertyConfigurator.configure("log4j.properties");
        BasicConfigurator.configure();

        // set input file name and file paths for the input and output
        String inputFileName = "example-file-2.docx";
        String inputFilePath = System.getProperty("user.dir") + "/docx-files/" + inputFileName;
        String outputFilePath = System.getProperty("user.dir") + "/output-files/";

        // check whether output directory exits, if it does not create it
        File file = new File(outputFilePath);
        if (!file.exists()) {
            if (file.mkdir()) {
                System.out.println("Output directory is created successfully.");
            } else {
                System.out.println("Failed to create output directory!");
            }
        }

        // create a new instance of DocExtractor to try its methods on example docx file
        DocxExtractor docxExtractor = new DocxExtractor(new File(inputFilePath));
        saveStringToFile(docxExtractor.getRawXML(),
            outputFilePath + "raw-xml-" + inputFileName + ".xml");
        saveStringToFile(docxExtractor.getObjectSchema(),
            outputFilePath + "object-schema-" + inputFileName + ".txt");
        docxExtractor.parse();

        System.out.println("-----------");

        System.out.println(docxExtractor.getText());
    }


    /**
     * @param inputString
     * @param fileName
     *
     * @throws FileNotFoundException
     */
    public static void saveStringToFile(String inputString, String fileName)
        throws FileNotFoundException {
        PrintWriter printWriter = new PrintWriter(new File(fileName));
        printWriter.println(inputString);
        printWriter.close();
    }
}
