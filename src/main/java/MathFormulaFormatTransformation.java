import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import java.io.StringReader;
import java.io.StringWriter;


/**
 * @author tokgozmusa
 */
public class MathFormulaFormatTransformation
{
    private final static String OMML2MML_XSL_FILE_NAME = System.getProperty("user.dir") + "/xsl-files/omml2mml.xsl";
    private final static String MML2TEX_XSL_FILE_NAME = System.getProperty("user.dir") + "/xsl-files/mmltex.xsl";


    /**
     * @param xmlString
     * @param xslFileName
     *
     * @return
     */
    private static String makeConversionWithXSLFile(String xmlString, String xslFileName)
    {
        String result = "";
        try
        {
            StringReader reader = new StringReader(xmlString);
            StringWriter writer = new StringWriter();
            TransformerFactory tFactory = TransformerFactory.newInstance();
            Transformer transformer = tFactory.newTransformer(new javax.xml.transform.stream.StreamSource(xslFileName));
            transformer.transform(
                new javax.xml.transform.stream.StreamSource(reader),
                new javax.xml.transform.stream.StreamResult(writer));
            result = writer.toString();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return result;
    }


    /**
     * @param ommlString
     *
     * @return
     */
    public static String OMML2MML(String ommlString)
    {
        return makeConversionWithXSLFile(ommlString, OMML2MML_XSL_FILE_NAME);
    }


    /**
     * @param mmlString
     *
     * @return
     */
    public static String MML2TeX(String mmlString)
    {
        return makeConversionWithXSLFile(mmlString, MML2TEX_XSL_FILE_NAME);
    }


    /**
     * @param ommlString
     *
     * @return
     */
    public static String OMML2TeX(String ommlString)
    {
        return MML2TeX(OMML2MML(ommlString));
    }
}
