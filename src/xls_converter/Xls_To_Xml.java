package xls_converter;

import jxl.*;
import jxl.read.biff.BiffException;
 
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Properties;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;
 

public class Xls_To_Xml {

	public static String get_info_FDN(String[] FDN_tab){
		String res="";
		
		return res;
	}
	public static void main(String[] args) throws SAXException, IOException {

        try
        {
          DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
          DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

          //root elements
          Document doc = docBuilder.newDocument();
                   doc.setXmlStandalone(false);
                   doc.setXmlVersion("1.0");
          Element root_Element = doc.createElement("bulkCmConfigDataFile");
          
        //set attributes  of the root element bulkCmConfigDataFile 
          // attribue xmlns:un
          Attr attr_root_xmlns_un = doc.createAttribute("xmlns:un");
          attr_root_xmlns_un.setValue("utranNrm.xsd");
          root_Element.setAttributeNodeNS(attr_root_xmlns_un);
          // attribue xmlns:es
          Attr attr_root_xmlns_es = doc.createAttribute("xmlns:es");
          attr_root_xmlns_es.setValue("EricssonSpecificAttributes.12.26.xsd");
          root_Element.setAttributeNodeNS(attr_root_xmlns_es);
          // attribue xmlns:xn
          Attr attr_root_xmlns_xn = doc.createAttribute("xmlns:xn");
          attr_root_xmlns_xn.setValue("genericNrm.xsd");
          root_Element.setAttributeNodeNS(attr_root_xmlns_xn);              
          // attribue xmlns:gn
          Attr attr_root_xmlns_gn = doc.createAttribute("xmlns:gn");
          attr_root_xmlns_gn.setValue("geranNrm.xsd");
          root_Element.setAttributeNodeNS(attr_root_xmlns_gn); 
          // attribue xmlns
          Attr attr_root_xmlns = doc.createAttribute("xmlns");
          attr_root_xmlns.setValue("configData.xsd");
          root_Element.setAttributeNode(attr_root_xmlns);               
          doc.appendChild(root_Element);
          
          //fileHeader elements
          Element fileHeader_Element = doc.createElement("fileHeader");
          root_Element.appendChild(fileHeader_Element);

          //set attributes to fileHeader element
          Attr attr_fileFormatVersion = doc.createAttribute("fileFormatVersion");
          attr_fileFormatVersion.setValue("32.615 V4.5");
          fileHeader_Element.setAttributeNode(attr_fileFormatVersion);
          
          Attr attr_vendorName = doc.createAttribute("vendorName");
          attr_vendorName.setValue("Ericsson");
          fileHeader_Element.setAttributeNode(attr_vendorName);


        //configData elements
          Element configData_Element = doc.createElement("configData");
          root_Element.appendChild(configData_Element);

          //set attributes to fileHeader element
          Attr attr_dnPrefix = doc.createAttribute("dnPrefix");
          attr_dnPrefix.setValue("DC=www.ericsson.com");
          configData_Element.setAttributeNode(attr_dnPrefix);
        
          //xn_SubNetwork elements
          Element xn_SubNetwork_element = doc.createElement("xn:SubNetwork");
          //set attributes to xn_SubNetwork element
          Attr attr_xn_SubNetwork = doc.createAttribute("id");
          attr_xn_SubNetwork.setValue("ONRM_ROOT_MO_R");
          xn_SubNetwork_element.setAttributeNode(attr_xn_SubNetwork);
          configData_Element.appendChild(xn_SubNetwork_element);	
		
		
           Workbook wrk1 = null;
		try {
			wrk1 = Workbook.getWorkbook(new File("XML_Neighbs_3g_aitbieda/3G_AitBieda_Relations_ADD.xls"));
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
             
            //Obtain the reference to the first sheet in the workbook
            Sheet UtranRelation_sheet = wrk1.getSheet(0);
      
        for (int row=1; row < UtranRelation_sheet.getRows();row++){
        	
		            //Read the contents of the Cell using getContents() method, which will return
		            //it as a String
		            String FDN = ((Cell)UtranRelation_sheet.getCell(0,row )).getContents();
		            String WRANrelation_ID = ((Cell)UtranRelation_sheet.getCell(1, row)).getContents();
		            String Source_cell = ((Cell)UtranRelation_sheet.getCell(2, row)).getContents();
		            String Source_RNC = ((Cell)UtranRelation_sheet.getCell(3, row)).getContents();
		            String Neighbour_cell = ((Cell)UtranRelation_sheet.getCell(4, row)).getContents();
		            String Neighbour_RNC = ((Cell)UtranRelation_sheet.getCell(5, row)).getContents();
		            String qOffset1sn = ((Cell)UtranRelation_sheet.getCell(6, row)).getContents();
		            String qOffset2sn = ((Cell)UtranRelation_sheet.getCell(7, row)).getContents();
		            String Load_Sharing = ((Cell)UtranRelation_sheet.getCell(8, row)).getContents();
		            String selectionPriority = ((Cell)UtranRelation_sheet.getCell(9, row)).getContents();
		            String qHcs = ((Cell)UtranRelation_sheet.getCell(10, row)).getContents();
		            String hcsPrio = ((Cell)UtranRelation_sheet.getCell(11, row)).getContents();
		            String penaltyTime = ((Cell)UtranRelation_sheet.getCell(12, row)).getContents();
		            String temporaryOffset1 = ((Cell)UtranRelation_sheet.getCell(13, row)).getContents();
		            String temporaryOffset2 = ((Cell)UtranRelation_sheet.getCell(14, row)).getContents();
		            String OPERATION = ((Cell)UtranRelation_sheet.getCell(15, row)).getContents();
		            String Validations = ((Cell)UtranRelation_sheet.getCell(16, row)).getContents();
		              
		            //Display the cell contents
		            System.out.println("FDN: "+FDN);
		            System.out.println("WCDMA RANrelation ID: "+WRANrelation_ID);
		            System.out.println("Source cell: "+Source_cell);
		            System.out.println("Source RNC: "+Source_RNC);
		            System.out.println("Neighbour cell: "+Neighbour_cell);
		            System.out.println("Neighbour RNC: "+Neighbour_RNC);
		            System.out.println("qOffset1sn: "+qOffset1sn);
		            System.out.println("qOffset2sn: "+qOffset2sn);
		            System.out.println("Load Sharing: "+Load_Sharing);
		            System.out.println("selection Priority: "+selectionPriority);
		            System.out.println("qHcs: "+qHcs);
		            System.out.println("hcsPrio: "+hcsPrio);
		            System.out.println("penalty Time: "+penaltyTime);
		            System.out.println("temporary Offset1: "+temporaryOffset1);
		            System.out.println("temporary Offset2: "+temporaryOffset2);
		            System.out.println("OPERATION: "+OPERATION);
		            System.out.println("Validations: "+Validations);
		        
         // here start the iteration of child nodes subnetworks
            //xn_SubNetwork_child elements
            Element xn_SubNetwork_child_element = doc.createElement("xn:SubNetwork");
            //set attributes to xn_SubNetwork_child element
            Attr attr_xn_SubNetwork_child = doc.createAttribute("id");
            attr_xn_SubNetwork_child.setValue(Source_RNC);//"RNCMAR1");
            xn_SubNetwork_child_element.setAttributeNode(attr_xn_SubNetwork_child);
            xn_SubNetwork_element.appendChild(xn_SubNetwork_child_element);
            
            
            //xn_MeContext elements
            Element xn_MeContext_element = doc.createElement("xn:MeContext");
            //set attributes to xn_SubNetwork_child element
            Attr attr_xn_MeContext = doc.createAttribute("id");
            attr_xn_MeContext.setValue(Source_RNC);//"RNCMAR1");
            xn_MeContext_element.setAttributeNode(attr_xn_MeContext);
            xn_SubNetwork_child_element.appendChild(xn_MeContext_element);
            
            
            //xn_ManagedElement elements
            Element xn_ManagedElement_element = doc.createElement("xn:ManagedElement");
            //set attributes to xn_SubNetwork_child element
            Attr attr_xn_ManagedElement = doc.createAttribute("id");
            attr_xn_ManagedElement.setValue("1");
            xn_ManagedElement_element.setAttributeNode(attr_xn_ManagedElement);
            xn_MeContext_element.appendChild(xn_ManagedElement_element);
            
            
            //un_RncFunction elements
            Element un_RncFunction_element = doc.createElement("un:RncFunction");
            //set attributes to xn_SubNetwork_child element
            Attr attr_un_RncFunction = doc.createAttribute("id");
            attr_un_RncFunction.setValue("1");
            un_RncFunction_element.setAttributeNode(attr_un_RncFunction);
            xn_ManagedElement_element.appendChild(un_RncFunction_element);
            
            //un_UtranCell elements
            Element un_UtranCell_element = doc.createElement("un:UtranCell");
            //set attributes to xn_SubNetwork_child element
            Attr attr_un_UtranCell = doc.createAttribute("id");
            attr_un_UtranCell.setValue(Source_cell);//"3G_AitBieda1");
            un_UtranCell_element.setAttributeNode(attr_un_UtranCell);
            un_RncFunction_element.appendChild(un_UtranCell_element);
            
            //un_UtranRelation elements
            Element un_UtranRelation_element = doc.createElement("un:UtranRelation");
            //set attributes to un_UtranRelation_child element
            Attr attr_un_UtranRelation_id = doc.createAttribute("id");
            attr_un_UtranRelation_id.setValue(WRANrelation_ID);//"61500_61501");
            un_UtranRelation_element.setAttributeNode(attr_un_UtranRelation_id);
            
            Attr attr_un_UtranRelation_modifier = doc.createAttribute("modifier");
            attr_un_UtranRelation_modifier.setValue("create");
            un_UtranRelation_element.setAttributeNode(attr_un_UtranRelation_modifier);
            
            un_UtranCell_element.appendChild(un_UtranRelation_element);
            
            //un_attributes elements
            Element un_attributes_element = doc.createElement("un:attributes");
            un_UtranRelation_element.appendChild(un_attributes_element);
            
            //un_adjacentCell elements
            Element un_adjacentCell_element = doc.createElement("un:adjacentCell");
            String [] FDN_Tab=FDN.split(",");
            String FDN_info = "SubNetwork=ONRM_ROOT_MO_R,SubNetwork="+Neighbour_RNC+",MeContext="+Neighbour_RNC+",ManagedElement=1,RncFunction=1,UtranCell="+Neighbour_cell;
            un_adjacentCell_element.setTextContent(FDN_info);
            un_attributes_element.appendChild(un_adjacentCell_element);
            
            //xn_VsDataContainer elements
            Element xn_VsDataContainer_element = doc.createElement("xn:VsDataContainer");
            //set attributes to un_UtranRelation_child element
            Attr attr_xn_VsDataContainer_id = doc.createAttribute("id");
            attr_xn_VsDataContainer_id.setValue(WRANrelation_ID);//"61500_61501");
            xn_VsDataContainer_element.setAttributeNode(attr_xn_VsDataContainer_id);
            
            Attr attr_xn_VsDataContainer_modifier = doc.createAttribute("modifier");
            attr_xn_VsDataContainer_modifier.setValue("create");
            xn_VsDataContainer_element.setAttributeNode(attr_xn_VsDataContainer_modifier);
            
            un_UtranRelation_element.appendChild(xn_VsDataContainer_element);
            
            
            //un_attributes elements
            Element xn_attributes_element = doc.createElement("xn:attributes");
            xn_VsDataContainer_element.appendChild(xn_attributes_element);
            
            //xn_vsDataType elements
            Element xn_vsDataType_element = doc.createElement("xn:vsDataType");
            xn_vsDataType_element.setTextContent("vsDataUtranRelation");
            xn_attributes_element.appendChild(xn_vsDataType_element);
            
            //xn_vsDataFormatVersion elements
            Element xn_vsDataFormatVersion_element = doc.createElement("xn:vsDataFormatVersion");
            xn_vsDataFormatVersion_element.setTextContent("EricssonSpecificAttributes.12.26");
            xn_attributes_element.appendChild(xn_vsDataFormatVersion_element);
            
          //es_vsDataUtranRelation elements
            Element es_vsDataUtranRelation_element = doc.createElement("es:vsDataUtranRelation");
            xn_attributes_element.appendChild(es_vsDataUtranRelation_element);
            
            //es:qOffset1sn elements
            Element es_qOffset1sn_element = doc.createElement("es:qOffset1sn");
            es_qOffset1sn_element.setTextContent(qOffset1sn);
            es_vsDataUtranRelation_element.appendChild(es_qOffset1sn_element);
            
            //es:qOffset2sn elements
            Element es_qOffset2sn_element = doc.createElement("es:qOffset2sn");
            es_qOffset2sn_element.setTextContent(qOffset2sn);
            es_vsDataUtranRelation_element.appendChild(es_qOffset2sn_element);
            
            //es:selectionPriority elements
            Element es_selectionPriority_element = doc.createElement("es:selectionPriority");
            es_selectionPriority_element.setTextContent(selectionPriority);
            es_vsDataUtranRelation_element.appendChild(es_selectionPriority_element);
            
            //es:hcsSib11Config elements
            Element es_hcsSib11Config_element = doc.createElement("es:hcsSib11Config");
            es_vsDataUtranRelation_element.appendChild(es_hcsSib11Config_element);
            
            //es:qHcs elements
            Element es_qHcs_element = doc.createElement("es:qHcs");
            es_qHcs_element.setTextContent(qHcs);
            es_hcsSib11Config_element.appendChild(es_qHcs_element);
            
            //es:hcsPrio elements
            Element es_hcsPrio_element = doc.createElement("es:hcsPrio");
            es_hcsPrio_element.setTextContent(hcsPrio);
            es_hcsSib11Config_element.appendChild(es_hcsPrio_element);
            
            //es:penaltyTime elements
            Element es_penaltyTime_element = doc.createElement("es:penaltyTime");
            es_penaltyTime_element.setTextContent(penaltyTime);
            es_hcsSib11Config_element.appendChild(es_penaltyTime_element);            
            
            //es:temporaryOffset1 elements
            Element es_temporaryOffset1_element = doc.createElement("es:temporaryOffset1");
            es_temporaryOffset1_element.setTextContent(temporaryOffset1);
            es_hcsSib11Config_element.appendChild(es_temporaryOffset1_element);            
                    
            //es:temporaryOffset2 elements
            Element es_temporaryOffset2_element = doc.createElement("es:temporaryOffset2");
            es_temporaryOffset2_element.setTextContent(temporaryOffset2);
            es_hcsSib11Config_element.appendChild(es_temporaryOffset2_element);            
                    
        }
            System.out.println("Nbr of Sets: "+UtranRelation_sheet.getRows());
            
            // Write the XML Test: 
           

              
                      
              
              
            //write the content into xml file
              TransformerFactory transformerFactory = TransformerFactory.newInstance();
              Transformer transformer = transformerFactory.newTransformer();
              transformer.setOutputProperty(OutputKeys.INDENT,"yes");
              Properties pr = transformer.getOutputProperties();
              PrintStream pstr = new PrintStream(System.out);
              pr.list(pstr);
              //transformer.setOutputProperty(OutputKeys.ENCODING,"UTF-8");
              DOMSource source = new DOMSource(doc);
               
              StreamResult result =  new StreamResult(new File("testing.xml"));
              transformer.transform(source, result);

              System.out.println("Writing XML Operation is Successfully Done");

            }catch(ParserConfigurationException pce){
              pce.printStackTrace();
            }catch(TransformerException tfe){
              tfe.printStackTrace();
            }

        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        //dbf.setCoalescing(true);
       // dbf.setIgnoringElementContentWhitespace(true);
        dbf.setIgnoringComments(true);
        DocumentBuilder db;
		try {
			db = dbf.newDocumentBuilder();
		
        Document doc1 = db.parse(new File("testing.xml"));
        doc1.normalizeDocument();

        Document doc2 = db.parse(new File("ref_utran.xml"));
        doc2.normalizeDocument();

        System.out.print(doc1.isEqualNode(doc2));
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        try
        {
          DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
          DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

          //root elements
          Document doc = docBuilder.parse(new File("ref_utran.xml"));
          TransformerFactory transformerFactory = TransformerFactory.newInstance();
          Transformer transformer2 = transformerFactory.newTransformer((Source) new File("ref_utran.xml"));
		/*try {
			 
			File fXmlFile = new File("C:/Documents and Settings/adnane.acer/Bureau/ref_utran.xml");
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
		 
			//optional, but recommended
			//read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
			doc.getDocumentElement().normalize();
			*/
			Properties pr = transformer2.getOutputProperties();
            PrintStream pstr = new PrintStream(System.out);
            pr.list(pstr);
		} catch (Exception e) {
			e.printStackTrace();
		    }
		
	}

}
