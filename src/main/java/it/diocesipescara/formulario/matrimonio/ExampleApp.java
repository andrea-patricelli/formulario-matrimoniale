package it.diocesipescara.formulario.matrimonio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import java.io.StringWriter;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import org.apache.fop.apps.FOPException;
import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.FopFactory;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToFoConverter;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.XMLHelper;
import org.apache.xmlgraphics.util.MimeConstants;

/**
 *
 * @author andrea
 */
public class ExampleApp {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
        String filePath = "/home/andrea/Work/tmp/certificato di matrimonio.dot";
        POIFSFileSystem fs = null;
        try {
            fs = new POIFSFileSystem(new FileInputStream(filePath));
            HWPFDocument doc = new HWPFDocument(fs);

            System.out.println("############## " + doc.getDocumentText());
            doc = replaceText(doc, "____", "REPLACED!!!!!!!!");
            saveWord(filePath, doc);
            System.out.println("############## " + doc.getDocumentText());
            final String dotToFO = dotToFO(doc);
            System.out.println("############## " + dotToFO);
            convertToPDF(dotToFO);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText) {
        Range r1 = doc.getRange();

        for (int i = 0; i < r1.numSections(); ++i) {
            Section s = r1.getSection(i);
            for (int x = 0; x < s.numParagraphs(); x++) {
                Paragraph p = s.getParagraph(x);
                for (int z = 0; z < p.numCharacterRuns(); z++) {
                    CharacterRun run = p.getCharacterRun(z);
                    String text = run.text();
                    System.out.println("REPLACE >>>>>>>>>>>>>>>>> " + text + " CONTAINS? " + text.contains(findText));
                    if (text.contains(findText)) {
                        run.replaceText(findText, replaceText);
                    }
                }
            }
        }
        return doc;
    }

    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
        try (FileOutputStream out = new FileOutputStream(filePath)) {
            doc.write(out);
        }
    }

    private static String dotToFO(HWPFDocument doc) throws ParserConfigurationException,
            TransformerConfigurationException, TransformerException {
        WordToFoConverter wordToFoConverter = new WordToFoConverter(
                XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument());
        wordToFoConverter.processDocument(doc);

        StringWriter stringWriter = new StringWriter();

        Transformer transformer = TransformerFactory.newInstance()
                .newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.transform(
                new DOMSource(wordToFoConverter.getDocument()),
                new StreamResult(stringWriter));

        return stringWriter.toString();
    }

    private static void convertToPDF(String source) throws IOException, FOPException, TransformerException {
        // the XSL FO file
//		File xsltFile = new File(RESOURCES_DIR + "//template.xsl");
        // the XML file which provides the input
//		StreamSource xmlSource = new StreamSource(new File(RESOURCES_DIR + "//Employees.xml"));
        StreamSource xmlSource = new StreamSource(new StringReader(source));
        // create an instance of fop factory
        FopFactory fopFactory = FopFactory.newInstance(new File(".").toURI());
        // a user agent is needed for transformation
        FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
        // Setup output
        OutputStream out;
        out = new java.io.FileOutputStream("/tmp/test_formulario.pdf");

        try {
            // Construct fop with desired output format
            Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);

            // Setup XSLT
            TransformerFactory factory = TransformerFactory.newInstance();
//			Transformer transformer = factory.newTransformer(new StreamSource(xsltFile));
            Transformer transformer = factory.newTransformer(new StreamSource(new StringReader(source)));

            // Resulting SAX events (the generated FO) must be piped through to
            // FOP
            Result res = new SAXResult(fop.getDefaultHandler());

            // Start XSLT transformation and FOP processing
            // That's where the XML is first transformed to XSL-FO and then
            // PDF is created
            transformer.transform(xmlSource, res);
        } finally {
            out.close();
        }
    }

}
