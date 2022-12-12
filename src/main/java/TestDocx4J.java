import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.python.core.PyObject;
import org.python.util.PythonInterpreter;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.docx4j.jaxb.Context.jc;

public class TestDocx4J {

    public static String replacer(PythonInterpreter pythonInterpreter, String text) throws UnsupportedEncodingException, UnsupportedEncodingException {
//        PyObject eval = pythonInterpreter.eval("to_cyrillic('" + text.substring(5, text.length()-6) + "')");
        PyObject eval = pythonInterpreter.eval("to_cyrillic('" + text + "')");
        String response = eval.toString();
        byte bytes[] = response.getBytes("ISO-8859-1");
        String result = new String(bytes, StandardCharsets.UTF_8);
//        result = text.replace(text, result);
//        result = text.replace(text.substring(5, text.length()-6), result);
        return result;
    }

    public static void main(String[] args) throws Docx4JException, JAXBException, UnsupportedEncodingException {

        PythonInterpreter pythonInterpreter = new PythonInterpreter();
        pythonInterpreter.execfile("translate.py");

        File doc = new File("bir_nimala.docx");
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(doc);
        MainDocumentPart mainDocumentPart = wordMLPackage
                .getMainDocumentPart();
        String textNodesXPath = "//w:t";
        List<Object> textNodes= mainDocumentPart
                .getJAXBNodesViaXPath(textNodesXPath, true);
        String xml = mainDocumentPart.getXML();
        for (Object obj : textNodes) {
            Text text = (Text) ((JAXBElement) obj).getValue();
            if(text.getValue() == " " || text.getValue() == "\n") {
                continue;
            }
            try {
                xml = xml.replace(text.getValue(), replacer(pythonInterpreter, text.getValue()));
            } catch (Exception e) {
                System.out.println(text.getValue());
            }

        }
        System.out.println(xml);
//        mainDocumentPart.setJaxbElement((org.docx4j.wml.Document) XmlUtils.unmarshalString(xml));
//        wordMLPackage.save(new File("Result.docx"));

    }
}
