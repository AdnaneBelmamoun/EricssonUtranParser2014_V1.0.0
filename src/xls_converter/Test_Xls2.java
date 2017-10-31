package xls_converter;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.stream.FactoryConfigurationError;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

public class Test_Xls2 {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		/*SAXParserFactory factory = SAXParserFactory.newInstance();
        SAXParser saxParser = factory.newSAXParser();
        saxParser.parse(new File("bookstore.xml"), new MyHandler());
	    */
		OutputStream outputStream;
		try {
			outputStream = new FileOutputStream(new File("C:/Documents and Settings/adnane.acer/Bureau/doc_test_xml_2.xml"));
		
		XMLStreamWriter out = XMLOutputFactory.newInstance().createXMLStreamWriter(new OutputStreamWriter(outputStream, "utf-8"));

		out.writeStartDocument("UTF-8", "1.0");
		out.writeStartElement("Root_element");

		out.writeStartElement("title");
		out.writeCharacters("Document Title");
		out.writeEndElement();

		out.writeEndElement();
		out.writeEndDocument();

		out.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (XMLStreamException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FactoryConfigurationError e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	
	}

}
