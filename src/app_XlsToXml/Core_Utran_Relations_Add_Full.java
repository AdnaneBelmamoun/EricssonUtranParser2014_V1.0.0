package app_XlsToXml;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;

import javanet.staxutils.IndentingXMLStreamWriter;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamWriter;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class Core_Utran_Relations_Add_Full{
	
	public static void Add_UtranRelations_conversion(String target_xls_dir_path,String XML_res_dir_suffix,String XML_res_File_name_suffix)throws Exception{
		
		// *************************************************************************************************************
        //		                   NodeB parameters preparation
		// *************************************************************************************************************

		String xls_target_directory_path =  target_xls_dir_path.replace("\\", "/");//+"/";
		
		// call of Method to find the XLS Relations_ADD file :
		String xls_path = get_xls_relations_add_path(xls_target_directory_path);//target_xls_dir_path);

		String[] tab_rbs_name= ((new File(xls_path)).getName().split("_"));

		System.out.println("xls_path:  "+ xls_target_directory_path+"\t"+tab_rbs_name.length);
		
		String rbs_name= tab_rbs_name[0]+"_"+tab_rbs_name[1];

		String XML_res_dir_name = rbs_name+XML_res_dir_suffix;
	
		
		// *************************************************************************************************************
        //  lecture fichier utran.xls
        
        Workbook wrk_relations_ADD = null;
      		try {
      			wrk_relations_ADD = Workbook.getWorkbook(new File(xls_path));
      		} catch (BiffException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		} catch (IOException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		}
                   
                  //Obtain the reference to the first sheet in the workbook
                  Sheet UtranRelation_sheet = wrk_relations_ADD.getSheet(0);
                  
                  if(!((String)((Cell)UtranRelation_sheet.getCell(0,1)).getContents()).isEmpty()){
                	  String op_res = ((Cell)UtranRelation_sheet.getCell(15,1)).getContents();
                			  //((Cell)(((Cell[])UtranRelation_sheet.getColumn(16))[1])).getContents();
                	  
		// Creation of the targeted Xml result directory: 
		File f = new File(xls_target_directory_path+XML_res_dir_name);
		f.mkdirs(); 
		
		String path_xml_result = (f.getPath()).replace("\\", "/");
       
		System.out.println("operation file type: \t"+op_res);
		
		
        String xml_file_fullpath = path_xml_result+"/"+rbs_name+"_"+op_res+XML_res_File_name_suffix; //"_Utran_Relations_Add"+".xml";
		
        try (
        		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(xml_file_fullpath), "UTF-8"))) {
		    } catch (IOException ex){}  
		
		
        // Creation of the XML File components:
        XMLOutputFactory xof = XMLOutputFactory.newInstance();
              
        FileOutputStream output = new FileOutputStream(xml_file_fullpath.replace("\\", "/"));
        
        XMLStreamWriter xtw  = xof.createXMLStreamWriter(output, "UTF-8");//new FileWriter(fileName));
        
        xtw = new IndentingXMLStreamWriter(xtw);
       
        // Starting Documents
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

        for (int row=1; row < UtranRelation_sheet.getRows();row++){
            	  
            	  String row_state = ((Cell)UtranRelation_sheet.getCell(0,row )).getContents();
            	  if(row_state.isEmpty()==false){
              	
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
      		        xtw.writeAttribute("modifier",OPERATION);
      		        
      		        xtw.writeStartElement("Utran","un:attributes");
                    xtw.writeStartElement("Utran","un:adjacentCell");
 String FDN_info = "SubNetwork=ONRM_ROOT_MO_R,SubNetwork="+Neighbour_RNC+",MeContext="+Neighbour_RNC+",ManagedElement=1,RncFunction=1,UtranCell="+Neighbour_cell;
      		        xtw.writeCharacters(FDN_info);
      		        xtw.writeEndElement();
      		        xtw.writeEndElement();
      		        
       		        xtw.writeStartElement("Utran","xn:VsDataContainer");
      		        xtw.writeAttribute("id",WRANrelation_ID);
      		        xtw.writeAttribute("modifier",OPERATION);
      		        
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
              
              }
              
              
              
              xtw.writeEndElement();
              xtw.writeEndElement();
              xtw.writeEmptyElement("Utran","fileFooter");
              xtw.writeAttribute("dateTime","2004-01-10T12:00:00+02:00");
              xtw.writeEndElement();
              xtw.flush();
              xtw.close();
              System.out.println("Writing XML Utran Relation First Carrier Operation is Successfully Done");

                  }      
		
	}
	
	public static void Add_RncGsmCells_conversion(String target_xls_dir_path,String XML_res_dir_suffix,String XML_res_File_name_suffix)throws Exception{
		// *************************************************************************************************************
        //		                   NodeB parameters preparationfor RncGsmCells
		// *************************************************************************************************************

		String xls_target_directory_path =  target_xls_dir_path.replace("\\", "/");//+"/";
		
		// call of Method to find the XLS Relations_ADD file :
		String xls_path = get_xls_relations_add_path(xls_target_directory_path);//target_xls_dir_path);

		String[] tab_rbs_name= ((new File(xls_path)).getName().split("_"));

		String rbs_name= tab_rbs_name[0]+"_"+tab_rbs_name[1];

		String XML_res_dir_name = rbs_name+XML_res_dir_suffix;
	
		
		// *************************************************************************************************************
        //                      lecture  ExternalGsmCell_sheet du fichier utran.xls
		// *************************************************************************************************************
	        
        Workbook wrk_relations_ADD = null;
      		try {
      			wrk_relations_ADD = Workbook.getWorkbook(new File(xls_path));
      		} catch (BiffException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		} catch (IOException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		}
                   
                  //Obtain the reference to the first sheet in the workbook
                  Sheet ExternalGsmCell_sheet = wrk_relations_ADD.getSheet(2);
              if(!((String)((Cell)ExternalGsmCell_sheet.getCell(0,1)).getContents()).isEmpty()) {
                  
            	  
            	  String op_res = ((Cell)ExternalGsmCell_sheet.getCell(11,1)).getContents();
              	
            	  System.out.println("number of rows in ExternalGsmCell_sheet :   "+ExternalGsmCell_sheet.getRows());
                  
		// Creation of the targeted Xml result directory: 
		File f = new File(xls_target_directory_path+XML_res_dir_name);
		f.mkdirs(); 
		
		String path_xml_result = (f.getPath()).replace("\\", "/");
        //System.out.println("path: "+path_xml_result);
		
		
        String xml_file_fullpath = path_xml_result+"/"+rbs_name+"_"+op_res+XML_res_File_name_suffix; //"_Utran_Relations_Add"+".xml";
		
        try (
        		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(xml_file_fullpath), "UTF-8"))) {
		    } catch (IOException ex){}  
		
		
        // Creation of the XML File components:
        XMLOutputFactory xof = XMLOutputFactory.newInstance();
              
        FileOutputStream output = new FileOutputStream(xml_file_fullpath.replace("\\", "/"));
        
        XMLStreamWriter xtw  = xof.createXMLStreamWriter(output, "UTF-8");//new FileWriter(fileName));
        
        xtw = new IndentingXMLStreamWriter(xtw);
       
        // Starting Documents
        xtw.writeStartDocument("UTF-8", "1.0");
                
          //xtw.
			//xtw.writeComment("all elements here are explicitly in the HTML namespace");
        xtw.setDefaultNamespace("RncGsmCells");//,"bulkCmConfigDataFile");
        xtw.writeStartElement("RncGsmCells","bulkCmConfigDataFile");
        xtw.writeAttribute("xmlns:un","utranNrm.xsd");
        xtw.writeAttribute("xmlns:es","EricssonSpecificAttributes.12.26.xsd");
        xtw.writeAttribute("xmlns:xn","genericNrm.xsd");
        xtw.writeAttribute("xmlns:gn","geranNrm.xsd");
        xtw.writeAttribute("xmlns","configData.xsd");
        
        xtw.writeEmptyElement("RncGsmCells","fileHeader");
        xtw.writeAttribute("fileFormatVersion","32.615 V4.5");
        xtw.writeAttribute("senderName","DC=www.ericsson.com,SubNetwork=Ericsson,IRPAgent=1");
        xtw.writeAttribute("vendorName","Ericsson");
        //xtw.writeEndElement();
        
        xtw.writeStartElement("RncGsmCells","configData");
        xtw.writeAttribute("dnPrefix","DC=www.ericsson.com");
       
        xtw.writeStartElement("RncGsmCells","xn:SubNetwork");
        xtw.writeAttribute("id","ONRM_ROOT_MO_R");
       
        
      
                  for (int row=1; row < ExternalGsmCell_sheet.getRows();row++){
                	  
                	  String row_state = ((Cell)ExternalGsmCell_sheet.getCell(0,row )).getContents();
                	  if(row_state.isEmpty()==false){
                  	
    		            //Read the contents of the Cell using getContents() method, which will return
    		            //it as a String
    		            String GSM_CELL_ID = ((Cell)ExternalGsmCell_sheet.getCell(0,row )).getContents();
    		            String USERLABEL = ((Cell)ExternalGsmCell_sheet.getCell(1, row)).getContents();
    		            String cl = ((Cell)ExternalGsmCell_sheet.getCell(2, row)).getContents();
    		            String ncc = ((Cell)ExternalGsmCell_sheet.getCell(3, row)).getContents();
    		            String bcc = ((Cell)ExternalGsmCell_sheet.getCell(4, row)).getContents();
    		            String bcchArfcn = ((Cell)ExternalGsmCell_sheet.getCell(5, row)).getContents();
    		            String lAC = ((Cell)ExternalGsmCell_sheet.getCell(6, row)).getContents();
    		            String maxTxPowerUl = ((Cell)ExternalGsmCell_sheet.getCell(7, row)).getContents();
    		            String qRxLevMin = ((Cell)ExternalGsmCell_sheet.getCell(8, row)).getContents();
    		            String individualOffset = ((Cell)ExternalGsmCell_sheet.getCell(9, row)).getContents();
    		            String bandIndicator = ((Cell)ExternalGsmCell_sheet.getCell(10, row)).getContents();
    		            String Operation = ((Cell)ExternalGsmCell_sheet.getCell(11, row)).getContents();
    		            String Vendor = ((Cell)ExternalGsmCell_sheet.getCell(12, row)).getContents();
    		              
    		            //Display the cell contents
    		            System.out.println("GSM CELL ID: "+GSM_CELL_ID);
    		            System.out.println("USER LABEL: "+USERLABEL);
    		            System.out.println("cl: "+cl);
    		            System.out.println("ncc: "+ncc);
    		            System.out.println("bcc: "+bcc);
    		            System.out.println("bcchArfcn: "+bcchArfcn);
    		            System.out.println("LAC: "+lAC);
    		            System.out.println("maxTxPowerUl: "+maxTxPowerUl);
    		            System.out.println("Load qRxLevMin: "+qRxLevMin);
    		            System.out.println("Individual Offset: "+individualOffset);
    		            System.out.println("bandIndicator: "+bandIndicator);
    		            System.out.println("Operation: "+Operation);
    		            System.out.println("Vendor: "+Vendor);
    		            
    		            
    		            
    		            xtw.writeStartElement("RncGsmCells","gn:ExternalGsmCell");
          		        xtw.writeAttribute("id",GSM_CELL_ID);
          		        xtw.writeAttribute("modifier","create");
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:attributes");
          		            		        
          		        xtw.writeStartElement("RncGsmCells","gn:userLabel");
          		        xtw.writeCharacters(USERLABEL);
          		        xtw.writeEndElement();

          		        xtw.writeStartElement("RncGsmCells","gn:cellIdentity");
          		        xtw.writeCharacters(cl);
          		        xtw.writeEndElement();
 
          		        xtw.writeStartElement("RncGsmCells","gn:ncc");
          		        xtw.writeCharacters(ncc);
          		        xtw.writeEndElement();
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:bcc");
          		        xtw.writeCharacters(bcc);
          		        xtw.writeEndElement();
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:bcchFrequency");
          		        xtw.writeCharacters(bcchArfcn);
          		        xtw.writeEndElement();
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:lac");
          		        xtw.writeCharacters(lAC);
          		        xtw.writeEndElement();
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:mcc");
          		        xtw.writeCharacters("604");
          		        xtw.writeEndElement();
          		        
          		        xtw.writeStartElement("RncGsmCells","gn:mnc");
          		        xtw.writeCharacters("01");
          		        xtw.writeEndElement();
          		        
          		        xtw.writeEndElement();
          		        
          		      xtw.writeStartElement("RncGsmCells","xn:VsDataContainer");
          		      xtw.writeAttribute("id",GSM_CELL_ID);
        		      xtw.writeAttribute("modifier",Operation);

        		      xtw.writeStartElement("RncGsmCells","xn:attributes");
        		        
        		        xtw.writeStartElement("RncGsmCells","xn:vsDataType");
        		        xtw.writeCharacters("vsDataExternalGsmCell");
        		        xtw.writeEndElement();

        		        xtw.writeStartElement("RncGsmCells","xn:vsDataFormatVersion");
        		        xtw.writeCharacters("EricssonSpecificAttributes.12.26");
        		        xtw.writeEndElement();

        		        xtw.writeStartElement("RncGsmCells","es:vsDataExternalGsmCell");
        		        
        		        xtw.writeStartElement("RncGsmCells","es:mncLength");
        		        xtw.writeCharacters("2");
        		        xtw.writeEndElement();
        		        
        		        xtw.writeStartElement("RncGsmCells","es:qRxLevMin");
        		        xtw.writeCharacters(qRxLevMin);
        		        xtw.writeEndElement();
        		        
        		        xtw.writeStartElement("RncGsmCells","es:maxTxPowerUl");
        		        xtw.writeCharacters(maxTxPowerUl);
        		        xtw.writeEndElement();
        		        
        		        xtw.writeStartElement("RncGsmCells","es:individualOffset");
        		        xtw.writeCharacters(individualOffset);
        		        xtw.writeEndElement();
        		        
        		        xtw.writeStartElement("RncGsmCells","es:bandIndicator");
        		        xtw.writeCharacters(bandIndicator);
        		        xtw.writeEndElement();
        		        
        		        xtw.writeEndElement();
        		        xtw.writeEndElement();
        		        xtw.writeEndElement();
        		        xtw.writeEndElement();
        		        
                   }

                  }
                  
                  xtw.writeEndElement();
                  xtw.writeEndElement();
                  xtw.writeEmptyElement("RncGsmCells","fileFooter");
                  xtw.writeAttribute("dateTime","2003-11-10T12:00:00+02:00");
                  xtw.writeEndElement();
                  xtw.flush();
                  xtw.close();
                  System.out.println("Writing XML (External Gsm Cell) Operation is Successfully Done");
              }
		
	}
	
	public static void Add_RncGsmRelations_conversion(String target_xls_dir_path,String XML_res_dir_suffix,String XML_res_File_name_suffix)throws Exception{
		// *************************************************************************************************************
        //		                   NodeB parameters preparation for RncGsmRelations
		// *************************************************************************************************************

		String xls_target_directory_path =  target_xls_dir_path.replace("\\", "/");//+"/";
		
		// call of Method to find the XLS Relations_ADD file :
		String xls_path = get_xls_relations_add_path(xls_target_directory_path);//target_xls_dir_path);

		String[] tab_rbs_name= ((new File(xls_path)).getName().split("_"));

		String rbs_name= tab_rbs_name[0]+"_"+tab_rbs_name[1];

		String XML_res_dir_name = rbs_name+XML_res_dir_suffix;
	
		
		// *************************************************************************************************************
        //               lecture fichier utran.xls
		// *************************************************************************************************************

		Workbook wrk_relations_ADD = null;
      		try {
      			wrk_relations_ADD = Workbook.getWorkbook(new File(xls_path));
      		} catch (BiffException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		} catch (IOException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		}
                   
                  //Obtain the reference to the first sheet in the workbook
                  Sheet GsmRelation_sheet = wrk_relations_ADD.getSheet(1);

                  if(!((String)((Cell)GsmRelation_sheet.getCell(0,1)).getContents()).isEmpty()) {
                	  
                	  String op_res = ((Cell)GsmRelation_sheet.getCell(8,1)).getContents();
                  	

		// Creation of the targeted Xml result directory: 
		File f = new File(xls_target_directory_path+XML_res_dir_name);
		f.mkdirs(); 
		
		String path_xml_result = (f.getPath()).replace("\\", "/");
        //System.out.println("path: "+path_xml_result);
		
		
        String xml_file_fullpath = path_xml_result+"/"+rbs_name+"_"+op_res+XML_res_File_name_suffix; //"_Utran_Relations_Add"+".xml";
		
        try (
        		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(xml_file_fullpath), "UTF-8"))) {
		    } catch (IOException ex){}  
		
		
        // Creation of the XML File components:
        XMLOutputFactory xof = XMLOutputFactory.newInstance();
              
        FileOutputStream output = new FileOutputStream(xml_file_fullpath.replace("\\", "/"));
        
        XMLStreamWriter xtw  = xof.createXMLStreamWriter(output, "UTF-8");//new FileWriter(fileName));
        
        xtw = new IndentingXMLStreamWriter(xtw);
       
        // Starting Documents
        xtw.writeStartDocument("UTF-8", "1.0");
                
          //xtw.
			//xtw.writeComment("all elements here are explicitly in the HTML namespace");
        xtw.setDefaultNamespace("GSMRelation");//,"bulkCmConfigDataFile");
        xtw.writeStartElement("GSMRelation","bulkCmConfigDataFile");
        xtw.writeAttribute("xmlns:un","utranNrm.xsd");
        xtw.writeAttribute("xmlns:es","EricssonSpecificAttributes.12.26.xsd");
        xtw.writeAttribute("xmlns:xn","genericNrm.xsd");
        xtw.writeAttribute("xmlns:gn","geranNrm.xsd");
        xtw.writeAttribute("xmlns","configData.xsd");
        
        xtw.writeEmptyElement("GSMRelation","fileHeader");
        xtw.writeAttribute("fileFormatVersion","32.615 V4.5");
        xtw.writeAttribute("senderName","DC=www.ericsson.com,SubNetwork=Ericsson,IRPAgent=1");
        xtw.writeAttribute("vendorName","Ericsson");
        //xtw.writeEndElement();
        
        xtw.writeStartElement("GSMRelation","configData");
        xtw.writeAttribute("dnPrefix","DC=www.ericsson.com");
       
        xtw.writeStartElement("GSMRelation","xn:SubNetwork");
        xtw.writeAttribute("id","ONRM_ROOT_MO_R");
       
                      
                  for (int row=1; row < GsmRelation_sheet.getRows();row++){
                	  
                	  String row_state = ((Cell)GsmRelation_sheet.getCell(0,row )).getContents();
                	  if(row_state.isEmpty()==false){
                  	
    		            //Read the contents of the Cell using getContents() method, which will return
    		            //it as a String
    		            String FDN = ((Cell)GsmRelation_sheet.getCell(0,row )).getContents();
    		            String GSM_Relation_ID = ((Cell)GsmRelation_sheet.getCell(1, row)).getContents();
    		            String Source_cell = ((Cell)GsmRelation_sheet.getCell(2, row)).getContents();
    		            String Source_RNC = ((Cell)GsmRelation_sheet.getCell(3, row)).getContents();
    		            String Neighbour_GSM_cell = ((Cell)GsmRelation_sheet.getCell(4, row)).getContents();
    		            String selectionPriority = ((Cell)GsmRelation_sheet.getCell(5, row)).getContents();
    		            String mobilityRelationType = ((Cell)GsmRelation_sheet.getCell(6, row)).getContents();
    		            String Qoffset1 = ((Cell)GsmRelation_sheet.getCell(7, row)).getContents();
    		            String Operation = ((Cell)GsmRelation_sheet.getCell(8, row)).getContents();
    		            String Validations = ((Cell)GsmRelation_sheet.getCell(9, row)).getContents();
    		            String Vendor = ((Cell)GsmRelation_sheet.getCell(10, row)).getContents();
    		            
    		              
    		            //Display the cell contents
    		            System.out.println("FDN: "+FDN);
    		            System.out.println("GSM Relation ID: "+GSM_Relation_ID);
    		            System.out.println("Source cell: "+Source_cell);
    		            System.out.println("Source RNC: "+Source_RNC);
    		            System.out.println("Neighbour GSM cell: "+Neighbour_GSM_cell);
    		            System.out.println("selection Priority: "+selectionPriority);
    		            System.out.println("mobility Relation Type: "+mobilityRelationType);
    		            System.out.println("Qoffset1: "+Qoffset1);
    		            System.out.println("Operation: "+Operation);
    		            System.out.println("Validations: "+Validations);
    		            System.out.println("Vendor: "+Vendor);
    		            
    		            
    		            xtw.writeStartElement("GSMRelation","xn:SubNetwork");
          		        xtw.writeAttribute("id",Source_RNC);
          		        
          		        xtw.writeStartElement("GSMRelation","xn:MeContext");
          		        xtw.writeAttribute("id",Source_RNC);
          		        
          		        xtw.writeStartElement("GSMRelation","xn:ManagedElement");
          		        xtw.writeAttribute("id","1");
          		        
          		        xtw.writeStartElement("GSMRelation","un:RncFunction");
          		        xtw.writeAttribute("id","1");
          		        
          		        xtw.writeStartElement("GSMRelation","un:UtranCell");
          		        xtw.writeAttribute("id",Source_cell);
          		        
           		        xtw.writeStartElement("GSMRelation","gn:GsmRelation");
          		        xtw.writeAttribute("id",GSM_Relation_ID);
          		        xtw.writeAttribute("modifier",Operation);
          		        
          		        xtw.writeStartElement("GSMRelation","gn:attributes");
                        xtw.writeStartElement("GSMRelation","gn:adjacentCell");
                     String FDN_info = "SubNetwork=ONRM_ROOT_MO_R,ExternalGsmCell="+Neighbour_GSM_cell;//+",MeContext="+Neighbour_RNC+",ManagedElement=1,RncFunction=1,UtranCell="+Neighbour_cell;
          		        xtw.writeCharacters(FDN_info);
          		        xtw.writeEndElement();
          		        xtw.writeEndElement();
          		        
           		        xtw.writeStartElement("GSMRelation","xn:VsDataContainer");
          		        xtw.writeAttribute("id",GSM_Relation_ID);
          		        xtw.writeAttribute("modifier","create");
          		        
          		        xtw.writeStartElement("GSMRelation","xn:attributes");
          		        
          		        xtw.writeStartElement("GSMRelation","xn:vsDataType");
          		        xtw.writeCharacters("vsDataGsmRelation");
         		        xtw.writeEndElement();
          		        
         		        xtw.writeStartElement("GSMRelation","xn:vsDataFormatVersion");
          		        xtw.writeCharacters("EricssonSpecificAttributes.12.26");
         		        xtw.writeEndElement();
         		        
         		        xtw.writeStartElement("GSMRelation", "es:vsDataGsmRelation");
         		        
         		        xtw.writeStartElement("GSMRelation", "es:qOffset1sn");
         		        xtw.writeCharacters(Qoffset1);
       		            xtw.writeEndElement();
         		        
       		            xtw.writeStartElement("GSMRelation", "es:selectionPriority");
         		        xtw.writeCharacters(selectionPriority);
       		            xtw.writeEndElement();
       		            
         		        xtw.writeStartElement("GSMRelation", "es:mobilityRelationType");
         		        xtw.writeCharacters(mobilityRelationType);
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
                  }
 
                  xtw.writeEndElement();
                  xtw.writeEndElement();
                  xtw.writeEmptyElement("GSMRelation","fileFooter");
                  xtw.writeAttribute("dateTime","2003-11-10T12:00:00+02:00");
                  xtw.writeEndElement();
                  xtw.flush();
                  xtw.close();
                  System.out.println("Writing XML  (RNC GSM Relations) Operation is Successfully Done");
                  }
	}
	
	public static void Add_UtranRelations_Second_Carrier_conversion(String target_xls_dir_path,String XML_res_dir_suffix,String XML_res_File_name_suffix)throws Exception{
		
		// *************************************************************************************************************
        //		                   NodeB parameters preparation
		// *************************************************************************************************************

		String xls_target_directory_path =  target_xls_dir_path.replace("\\", "/");//+"/";
		
		// call of Method to find the XLS Relations_ADD file :
		String xls_path = get_xls_relations_Second_Carrier_add_path(xls_target_directory_path);//target_xls_dir_path);

		String[] tab_rbs_name= ((new File(xls_path)).getName().split("_"));

		String[] temp = (tab_rbs_name[1]).split("-");
		
		String rbs_name= tab_rbs_name[0]+"_"+temp[0];

		String XML_res_dir_name = rbs_name+XML_res_dir_suffix;
	
		
		// *************************************************************************************************************
        //                          lecture fichier utran.xls
        // *************************************************************************************************************
        Workbook wrk_relations_ADD = null;
      		try {
      			wrk_relations_ADD = Workbook.getWorkbook(new File(xls_path));
      		} catch (BiffException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		} catch (IOException e) {
      			// TODO Auto-generated catch block
      			e.printStackTrace();
      		}
                   
                  //Obtain the reference to the first sheet in the workbook
                  Sheet UtranRelation_sheet = wrk_relations_ADD.getSheet(0);
                  
                  if(!((String)((Cell)UtranRelation_sheet.getCell(0,1)).getContents()).isEmpty()){
                	  
               String op_res = ((Cell)UtranRelation_sheet.getCell(15,1)).getContents();
                  	
		// Creation of the targeted Xml result directory: 
		File f = new File(xls_target_directory_path+XML_res_dir_name);
		f.mkdirs(); 
		
		String path_xml_result = (f.getPath()).replace("\\", "/");
        //System.out.println("path: "+path_xml_result);
		
		
        String xml_file_fullpath = path_xml_result+"/"+rbs_name+"_"+op_res+XML_res_File_name_suffix; //"_Utran_Relations_Add"+".xml";
		
        try (
        		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(xml_file_fullpath), "UTF-8"))) {
		    } catch (IOException ex){}  
		
		
        // Creation of the XML File components:
        XMLOutputFactory xof = XMLOutputFactory.newInstance();
              
        FileOutputStream output = new FileOutputStream(xml_file_fullpath.replace("\\", "/"));
        
        XMLStreamWriter xtw  = xof.createXMLStreamWriter(output, "UTF-8");//new FileWriter(fileName));
        
        xtw = new IndentingXMLStreamWriter(xtw);
       
        // Starting Documents
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

        for (int row=1; row < UtranRelation_sheet.getRows();row++){
            	  
            	  String row_state = ((Cell)UtranRelation_sheet.getCell(0,row )).getContents();
            	  if(row_state.isEmpty()==false){
              	
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
      		        xtw.writeAttribute("modifier",OPERATION);
      		        
      		        xtw.writeStartElement("Utran","un:attributes");
                    xtw.writeStartElement("Utran","un:adjacentCell");
 String FDN_info = "SubNetwork=ONRM_ROOT_MO_R,SubNetwork="+Neighbour_RNC+",MeContext="+Neighbour_RNC+",ManagedElement=1,RncFunction=1,UtranCell="+Neighbour_cell;
      		        xtw.writeCharacters(FDN_info);
      		        xtw.writeEndElement();
      		        xtw.writeEndElement();
      		        
       		        xtw.writeStartElement("Utran","xn:VsDataContainer");
      		        xtw.writeAttribute("id",WRANrelation_ID);
      		        xtw.writeAttribute("modifier",OPERATION);
      		        
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
              
              }
              
              
              
              xtw.writeEndElement();
              xtw.writeEndElement();
              xtw.writeEmptyElement("Utran","fileFooter");
              xtw.writeAttribute("dateTime","2004-01-10T12:00:00+02:00");
              xtw.writeEndElement();
              xtw.flush();
              xtw.close();
              System.out.println("Writing XML Second Carrier Operation is Successfully Done");

                  }
		
	}

  		
	public static String get_xls_relations_Second_Carrier_add_path(String XLS_directory_path){
		String resulting_path="";
		File[] tab_xls_files = new File(XLS_directory_path).listFiles();
		
		for(int i=0; i<tab_xls_files.length;i++){
			if(tab_xls_files[i].isFile() && tab_xls_files[i].getName().endsWith(".xls")){
			      if(tab_xls_files[i].getName().contains("2ndCarrier")){
			    	  resulting_path =  tab_xls_files[i].getPath().replace("\\", "/");
			      }
		       }
			}
		
		return resulting_path;
	}
	public static String get_xls_relations_add_path(String XLS_directory_path){
		String resulting_path="";
		File[] tab_xls_files = new File(XLS_directory_path).listFiles();
		File ftemp=null;
		for(int i=0; i<tab_xls_files.length;i++){
			ftemp=(File)tab_xls_files[i];
			System.out.println(ftemp.getName());
			if(ftemp.isFile() && ftemp.getName().endsWith(".xls")){
			      if(tab_xls_files[i].getName().contains("_Relations_1stCarrier")|| tab_xls_files[i].getName().contains("_Relations") ){
			    	  if(!tab_xls_files[i].getName().contains("2ndCarrier")){
			    	  resulting_path =  tab_xls_files[i].getPath().replace("\\", "/");
			    	  }
			      }
		       }
			}
		
		return resulting_path;
	}
	
	public static void main(String[] args) {
		String rbs_name ="3G_elfaid";//"3G_BOUTROUCH";// "3G_Tichla";//"XML_Neighbs_3g_aitbieda";//"3g_lamhadi";
				String target_xls_dir_path = (System.getProperty("user.home") + "\\Bureau\\"+rbs_name+"\\").replace("\\", "/");
	//"C:/Documents and Settings/admin/Bureau/Neighbs_3g_lamhadi/";

				System.out.print((System.getProperty("user.home") + "\\Bureau\\3g_lamhadi\\").replace("\\", "/"));	

				
				String XML_res_dir_suffix = "_XML_result/";
				String XML_res_File_name_suffix_UtranRelation = "_UtranRelations"+".xml";
				String XML_res_File_name_suffix_UtranRelation_2ndCarrier = "_UtranRelations_2ndCarrier"+".xml";
				String XML_res_File_name_suffix_RncGsmCells = "_RncGsmCells"+".xml";
				String XML_res_File_name_suffix_RncGsmRelations = "_RncGsmRelations"+".xml";

				
				// *******************************************************************************
				//        Conversion Step from XLS  to XML file for Utran_relations_ADD  
				// *******************************************************************************
				try {
					Add_UtranRelations_conversion(target_xls_dir_path,XML_res_dir_suffix,XML_res_File_name_suffix_UtranRelation);
				} catch (Exception e) {   e.printStackTrace();  }

				
				
				// *******************************************************************************
				//        Conversion Step from XLS  to XML file for RncGsmCells_ADD
				// *******************************************************************************
				
				try {
					Add_RncGsmCells_conversion(target_xls_dir_path,XML_res_dir_suffix,XML_res_File_name_suffix_RncGsmCells);
				} catch (Exception e) {   e.printStackTrace();  }

				
				
				
				// *******************************************************************************
				//        Conversion Step from XLS  to XML file for RncGsmRelations_ADD  
				// *******************************************************************************
				try {
					Add_RncGsmRelations_conversion(target_xls_dir_path,XML_res_dir_suffix,XML_res_File_name_suffix_RncGsmRelations);
				} catch (Exception e) {   e.printStackTrace();  }

				
				
				
				// ****************************************************************************************************
				//        Conversion Step from XLS  to XML file for Second Carrier Utran_relations_ADD  
				// ****************************************************************************************************
				try {
					Add_UtranRelations_Second_Carrier_conversion(target_xls_dir_path,XML_res_dir_suffix,XML_res_File_name_suffix_UtranRelation_2ndCarrier);
				} catch (Exception e) {   e.printStackTrace();  }

				

				
	}

}



/*
String xls_target_directory_path =  target_xls_dir_path.replace("\\", "/");//+"/";

String xls_path = get_xls_relations_add_path(xls_target_directory_path);//target_xls_dir_path);

String[] tab_rbs_name= ((new File(xls_path)).getName().split("_"));

String rbs_name= tab_rbs_name[0]+"_"+tab_rbs_name[1];

String XML_res_dir_name = rbs_name+XML_res_dir_suffix;
*/
//String Xml_file_full_Path_Utran = xls_target_directory_path+ XML_res_dir_name + rbs_name + XML_res_File_name;



//System.out.println("Processed NodeB = "+rbs_name);

/*System.out.println("Relations ADD file path = "+get_xls_relations_add_path(xls_target_directory_path));
System.out.println("Relations ADD directory path = "+xls_target_directory_path);
System.out.println("Relations XML directory path = "+Xml_file_full_Path_Utran);*/
