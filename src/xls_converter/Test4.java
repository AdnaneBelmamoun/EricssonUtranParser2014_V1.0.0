package xls_converter;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import javanet.staxutils.IndentingXMLEventWriter;

import javax.xml.stream.*;

public class Test4 {

	/**
	 * @param args
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		XMLOutputFactory xmlOutputFactory = XMLOutputFactory.newInstance();
		FileOutputStream file = new FileOutputStream("C:/Documents and Settings/adnane.acer/Bureau/result_test_stax2.xml");
		XMLEventWriter writer = xmlOutputFactory.createXMLEventWriter(file);
		writer = new IndentingXMLEventWriter(writer);
		XMLEventFactory eventFactory = XMLEventFactory.newInstance();
		writer.add(eventFactory.createStartDocument());
		writer.add(eventFactory.createStartElement("a", "aa", "a"));
		writer.add(eventFactory.createStartElement("b", "bb", "b"));
		writer.add(eventFactory.createEndElement("a", "aa", "b"));
		writer.add(eventFactory.createEndElement("b", "bb", "a"));
		writer.add(eventFactory.createEndDocument());
	}

}
