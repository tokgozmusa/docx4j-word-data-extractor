import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.Callback;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.math.CTOMath;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.samples.AbstractSample;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Text;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.io.InputStream;
import java.util.*;


/*
    SAMPLES:
    https://github.com/plutext/docx4j/tree/master/src/main/java/org/docx4j/samples
    https://github.com/plutext/docx4j/tree/master/src/samples/docx4j/org/docx4j/samples
*/

/**
 * This class consists methods that you can extract data from a docx file
 *
 * @author tokgozmusa
 */
public class DocxExtractor extends AbstractSample {

    private WordprocessingMLPackage wordMLPackage = null;


    /**
     * Class constructor.
     */
    public DocxExtractor(File inputFile) {
        try {
            this.loadDocument(inputFile);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }


    /**
     * Class constructor.
     */
    public DocxExtractor(InputStream inputStream) {
        try {
            this.loadDocument(inputStream);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }


    /**
     * @param inputFile
     *
     * @throws Docx4JException
     */
    private void loadDocument(File inputFile) throws Docx4JException {
        wordMLPackage = WordprocessingMLPackage.load(inputFile);
    }


    /**
     * @param inputStream
     *
     * @throws Docx4JException
     */
    private void loadDocument(InputStream inputStream) throws Docx4JException {
        wordMLPackage = WordprocessingMLPackage.load(inputStream);
    }


    /**
     * @param obj
     * @param toSearch
     *
     * @return
     */
    private List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }
        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }


    /**
     * @param object
     *
     * @return
     */
    private String getTextOfObject(Object object) {
        // object can be org.docx4j.wml.P, org.docx4j.wml.R etc
        List<Object> texts = getAllElementFromObject(object, Text.class);
        String result = "";
        for (Object o : texts) {
            Text t = (Text) o;
            result += t.getValue();
        }
        return result;
    }


    /**
     * @return
     */
    public String getText() {
        if (wordMLPackage == null) {
            throw new NullPointerException();
        }

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
            .getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        return getTextOfObject(body);
    }


    /**
     * @return raw XML of the docx file as string
     */
    public String getRawXML() {
        if (wordMLPackage == null) {
            throw new NullPointerException();
        }
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        return XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true);
    }


    /**
     * @return
     *
     * @throws Docx4JException
     */
    public String getObjectSchema() throws Docx4JException {
        if (wordMLPackage == null) {
            throw new NullPointerException();
        }

        final StringBuilder objectSchemaStrBuilder = new StringBuilder();
        objectSchemaStrBuilder.append("org.docx4j.wml.Body\n");

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
            .getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        // start recursive travel in document
        new TraversalUtil(body,
            new Callback() {
                String indent = "";

                public List<Object> apply(Object o) {
                    objectSchemaStrBuilder.append(indent).append(o.getClass().getName())
                        .append("\n");
                    return null;
                }

                public boolean shouldTraverse(Object o) {
                    return true;
                }

                public void walkJAXBElements(Object parent) {
                    indent += "    ";
                    List children = getChildren(parent);
                    if (children != null) {
                        for (Object o : children) {
                            o = XmlUtils.unwrap(o);
                            this.apply(o);
                            if (this.shouldTraverse(o)) {
                                walkJAXBElements(o);
                            }
                        }
                    }
                    indent = indent.substring(0, indent.length() - 4);
                }

                public List<Object> getChildren(Object o) {
                    return TraversalUtil.getChildrenImpl(o);
                }
            }
        );
        return objectSchemaStrBuilder.toString();
    }


    /**
     * @throws Exception
     */
    public void parse() throws Exception {
        if (wordMLPackage == null) {
            throw new NullPointerException();
        }

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
            .getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        // start recursive travel in document
        new TraversalUtil(body,
            new Callback() {
                public List<Object> apply(Object o) {
                    if (o instanceof org.docx4j.wml.P) {
                        org.docx4j.wml.P p = (org.docx4j.wml.P) o;

                        // bullet - enumeration - numbering
                        if (p.getPPr() != null && p.getPPr().getPStyle() != null) {
                            if (p.getPPr().getPStyle().getVal().equals("ListeParagraf")) {
                                // TODO
                                int paragraphID = p.getPPr().getNumPr().getNumId().getVal()
                                    .intValue();
                                System.out.println("paragraphID: " + paragraphID);
                            }
                        }
                        System.out.println("-----NEW PARAGRAPH-----");
                    } else if (o instanceof org.docx4j.wml.R) {
                        org.docx4j.wml.R r = (org.docx4j.wml.R) o;

                        // check if there is bold text
                        if (r.getRPr() != null && r.getRPr().getB() != null) {
                            // TODO
                            boolean isBold = r.getRPr().getB().isVal();
                            String getTextOfBold = getTextOfObject(r);
                            System.out.println("there is a bold text");
                            System.out.println(getTextOfBold);
                        }
                    } else if (o instanceof Text) {
                        // TODO
                        String text = ((Text) o).getValue();
                        System.out.println(text);
                    } else if (o instanceof CTOMath) {
                        // make prettyPrint false to keep it one line
                        boolean prettyPrint = false;
                        String formula = XmlUtils.marshaltoString(o, true, prettyPrint,
                            Context.jc,
                            "http://schemas.openxmlformats.org/officeDocument/2006/math",
                            "oMath",
                            CTOMath.class);
                        String teXFormat = MathFormulaFormatTransformation.OMML2TeX(formula);
                        if (teXFormat != null && !teXFormat.equals("")) {
                            // adapt teXFormat formula to our rich-text editor
                            teXFormat = teXFormat.substring(2, teXFormat.length() - 1);
                            teXFormat = "\\(" + teXFormat + "\\)";
                            // TODO: get teXFormat string
                            System.out.println(teXFormat);
                        }
                    } else if (o instanceof Drawing) {
                        System.out.println("<<There is a picture here!!!>>");
                    }
                    return null;
                }

                public boolean shouldTraverse(Object o) {
                    return true;
                }

                public void walkJAXBElements(Object parent) {
                    List children = getChildren(parent);
                    if (children != null) {
                        for (Object o : children) {
                            o = XmlUtils.unwrap(o);
                            this.apply(o);
                            if (this.shouldTraverse(o)) {
                                walkJAXBElements(o);
                            }
                        }
                    }
                }

                public List<Object> getChildren(Object o) {
                    return TraversalUtil.getChildrenImpl(o);
                }
            }
        );
    }
}
