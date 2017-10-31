package xls_converter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javanet.staxutils.IndentingXMLStreamWriter;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Test_Stax_API {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception{
        String fileName = "C:/Documents and Settings/adnane.acer/Bureau/result_test_stax.xml";
        XMLOutputFactory xof = XMLOutputFactory.newInstance();
        //xof.setProperty(XMLOutputFactory.IS_REPAIRING_NAMESPACES, "/n");
        FileOutputStream output = new FileOutputStream(fileName);
        
        XMLStreamWriter xtw  = xof.createXMLStreamWriter(output, "UTF-8");//new FileWriter(fileName));
        
        xtw = new IndentingXMLStreamWriter(xtw);
        
        xtw.writeStartDocument("UTF-8", "1.0");
                
          //xtw.
			//xtw.writeComment("all elements here are explicitly in the HTML namespace");
        xtw.setDefaultNamespace("Utran");//,"bulkCmConfigDataFile");
        xtw.writeStartElement("Utran","bulkCmConfigDataFile");
        xtw.writeAttribute("xmlns:un","utranNrm.xsd");
        xtw.writeAttribute("xmlns:es","EricssonSpecificAttributes.12.26.xsd");
        xtw.writeAttribute("xmlns:xn","genericNrm.xsd");
        xtw.writeAttribute("xmlns:gn","geranNrm.xsd");
        xtw.writeAttribute("xmlns","configData.xsd");
        
        xtw.writeEmptyElement("Utran","fileHeader");
        xtw.writeAttribute("fileFormatVersion","32.615 V4.5");
        xtw.writeAttribute("vendorName","Ericsson");
        //xtw.writeEndElement();
        
        xtw.writeStartElement("Utran","configData");
        xtw.writeAttribute("dnPrefix","DC=www.ericsson.com");
       
        xtw.writeStartElement("Utran","xn:SubNetwork");
        xtw.writeAttribute("id","ONRM_ROOT_MO_R");
       
        
      // lecture fichier utran.xls
        
        Workbook wrk1 = null;
      		try {
      			wrk1 = Workbook.getWorkbook(new File("C:/Documents and Settings/adnane.acer/Bureau/XML_Neighbs_3g_aitbieda/3G_AitBieda_Relations_ADD.xls"));
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
      		        
            
              
      		        xtw.writeStartElement("Utran","xn:SubNetwork");
      		        xtw.writeAttribute("id",Source_RNC);
      		        
      		        xtw.writeStartElement("Utran","xn:MeContext");
      		        xtw.writeAttribute("id",Source_RNC);
      		        
      		        xtw.writeStartElement("Utran","xn:ManagedElement");
      		        xtw.writeAttribute("id","1");
      		        
      		        xtw.writeStartElement("Utran","un:RncFunction");
      		        xtw.writeAttribute("id","1");
      		        
      		        xtw.writeStartElement("Utran","un:UtranCell");
      		        xtw.writeAttribute("id",Source_cell);
      		        
       		        xtw.writeStartElement("Utran","un:UtranRelation");
      		        xtw.writeAttribute("id",WRANrelation_ID);
      		        xtw.writeAttribute("modifier","create");
      		        
      		        xtw.writeStartElement("Utran","un:attributes");
                    xtw.writeStartElement("Utran","un:adjacentCell");
 String FDN_info = "SubNetwork=ONRM_ROOT_MO_R,SubNetwork="+Neighbour_RNC+",MeContext="+Neighbour_RNC+",ManagedElement=1,RncFunction=1,UtranCell="+Neighbour_cell;
      		        xtw.writeCharacters(FDN_info);
      		        xtw.writeEndElement();
      		        xtw.writeEndElement();
      		        
       		        xtw.writeStartElement("Utran","xn:VsDataContainer");
      		        xtw.writeAttribute("id",WRANrelation_ID);
      		        xtw.writeAttribute("modifier","create");
      		        
      		        xtw.writeStartElement("Utran","xn:attributes");
      		        
      		        xtw.writeStartElement("Utran","xn:vsDataType");
      		        xtw.writeCharacters("vsDataUtranRelation");
     		        xtw.writeEndElement();
      		        
     		        xtw.writeStartElement("Utran","xn:vsDataFormatVersion");
      		        xtw.writeCharacters("EricssonSpecificAttributes.12.26");
     		        xtw.writeEndElement();
     		        
     		        xtw.writeStartElement("Utran", "es:vsDataUtranRelation");
     		        
     		        xtw.writeStartElement("Utran", "es:qOffset1sn");
     		        xtw.writeCharacters(qOffset1sn);
   		            xtw.writeEndElement();
     		        
   		            xtw.writeStartElement("Utran", "es:qOffset2sn");
     		        xtw.writeCharacters(qOffset2sn);
   		            xtw.writeEndElement();
   		            
     		        xtw.writeStartElement("Utran", "es:selectionPriority");
     		        xtw.writeCharacters(selectionPriority);
   		            xtw.writeEndElement();
     		        
   		            xtw.writeStartElement("Utran", "es:hcsSib11Config");
     		        
   		            xtw.writeStartElement("Utran", "es:qHcs");
  		            xtw.writeCharacters(qHcs);
		            xtw.writeEndElement();
  		        
		            xtw.writeStartElement("Utran", "es:hcsPrio");
  		            xtw.writeCharacters(hcsPrio);
		            xtw.writeEndElement();
		            
  		            xtw.writeStartElement("Utran", "es:penaltyTime");
  		            xtw.writeCharacters(penaltyTime);
		            xtw.writeEndElement();
   		            
		            xtw.writeStartElement("Utran", "es:temporaryOffset1");
  		            xtw.writeCharacters(temporaryOffset1);
		            xtw.writeEndElement();
		            
  		            xtw.writeStartElement("Utran", "es:temporaryOffset2");
  		            xtw.writeCharacters(temporaryOffset2);
		            xtw.writeEndElement();
   		            
		            
     		        xtw.writeEndElement();
     		       xtw.writeEndElement();
     		      xtw.writeEndElement();
     		     xtw.writeEndElement();
     		    xtw.writeEndElement();
     		   xtw.writeEndElement();
     		  xtw.writeEndElement();
     		 xtw.writeEndElement();
     		xtw.writeEndElement();
     		xtw.writeEndElement();
  
              
              }
              
              
              
              xtw.writeEndElement();
              xtw.writeEndElement();
              xtw.writeEmptyElement("Utran","fileFooter");
              xtw.writeAttribute("dateTime","2004-01-10T12:00:00+02:00");
              xtw.writeEndElement();
              xtw.flush();
              xtw.close();
              System.out.println("Writing XML Operation is Successfully Done");
                  
      
	}

}
