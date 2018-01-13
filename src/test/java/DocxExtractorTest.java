import static org.junit.Assert.assertEquals;

import java.io.File;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.PropertyConfigurator;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class DocxExtractorTest {

    DocxExtractor docxExtractor = null;

    @Before
    public void setUp() throws Exception {
        // configure log4j
        PropertyConfigurator.configure("log4j.properties");
        BasicConfigurator.configure();

        // setup DocxExtractor
        String inputFileName = "example-file-1.docx";
        String inputFilePath = System.getProperty("user.dir") + "/docx-files/" + inputFileName;
        docxExtractor = new DocxExtractor(new File(inputFilePath));
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void getText() throws Exception {
        assertEquals("Hello world! This is a test file.", docxExtractor.getText());
    }

    @Test
    public void getRawXML() throws Exception {
    }

    @Test
    public void getObjectSchema() throws Exception {
    }

    @Test
    public void parse() throws Exception {
    }
}
