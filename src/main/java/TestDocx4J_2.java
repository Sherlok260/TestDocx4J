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
import java.io.FileWriter;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestDocx4J_2 {

    public static void main(String[] args) throws Docx4JException, JAXBException, IOException {

        Map<String, String> map = new HashMap<>();
        map.put("Reja", "almashdi");

        File doc = new File("Reja.docx");
//        File doc = new File("bir_nimala.docx");
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(doc);
        MainDocumentPart mainDocumentPart = wordMLPackage
                .getMainDocumentPart();
        String textNodesXPath = "//w:t";
//        List<Object> textNodes= mainDocumentPart
//                .getJAXBNodesViaXPath(textNodesXPath, false);
        String xml = mainDocumentPart.getXML();

//        for (Object obj : textNodes) {
//            Text text = (Text) ((JAXBElement) obj).getValue();
//
//        }

//        XmlUtils.unmarshallFromTemplate(xml, map);

//        mainDocumentPart.setJaxbElement((org.docx4j.wml.Document) XmlUtils.unmarshalString(xml));
//        mainDocumentPart.variableReplace(map);
        wordMLPackage.save(new File("Result.docx"));

//        FileWriter fileWriter = new FileWriter("xmlDoc.xml");
//        fileWriter.write(xml);
    }
}
