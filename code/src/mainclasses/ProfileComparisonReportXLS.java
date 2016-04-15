/**
 * 
 * This class compares two Profile xml files and generates the comparison report in excel sheet. 
 * 
 * */
package mainclasses;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ProfileComparisonReportXLS {
	

  public static void main(String argv[]) {
	  try{
	  Workbook profileComparisonWorkBook = new HSSFWorkbook();
	  //Create DOM object for both files  
	  Document doc1 = parseFile(argv[0]+".profile");
	  Document doc2 = parseFile(argv[1]+".profile");
	  //Part 1 - FLS Comparison
	  Map<String,List<String>> mapFieldPermissionsProfile1 = 	retrieveFieldPermissions(doc1);
	  Map<String,List<String>> mapFieldPermissionsProfile2 = 	retrieveFieldPermissions(doc2);
	  Map<String,List<String>> comparisonReportFLS = compareFLS(mapFieldPermissionsProfile1, mapFieldPermissionsProfile2);
	  GenerateFLSSheet(comparisonReportFLS, argv, profileComparisonWorkBook);
		//Part 2 Layout Assignments
	  Map<String,String> pageLayoutAssignmentProf1 = retrievePageLayoutAssignment(doc1);
	  Map<String,String> pageLayoutAssignmentProf2 = retrievePageLayoutAssignment(doc2);
	  Map<String,List<String>> mapLayoutComparisonReport = compareLayoutAssignments(pageLayoutAssignmentProf1, pageLayoutAssignmentProf2);
	  GenerateLayoutSheet(mapLayoutComparisonReport, profileComparisonWorkBook, argv);
	  
	  //Part 3 Object Permissions
	  Map<String, objectPermissions> mapObjectPermissionsProf1 = retrieveObjectPermissions(doc1);
	  Map<String, objectPermissions> mapObjectPermissionsProf2 = retrieveObjectPermissions(doc2);
	  Map<String, List<objectPermissions>> compareObjectPermissions = compareObjectPermissions(mapObjectPermissionsProf1, mapObjectPermissionsProf2);
	  GenerateObjectPermissionSheet(compareObjectPermissions, argv, profileComparisonWorkBook);
	  
	  //Part 4 Retrieve Application Visibility
	  Map<String, applicationVisibilities> mapAppPermissionsProf1  = retrieveApplicationVisibilities(doc1);
	  Map<String, applicationVisibilities> mapAppPermissionsProf2  = retrieveApplicationVisibilities(doc2);
	  Map<String, List<applicationVisibilities>> compareAppPermissions = compareApplicationVisibility(mapAppPermissionsProf1, mapAppPermissionsProf2);
	  //System.out.print("compareAppPermissions = " + compareAppPermissions);
	  GenerateAppPermissionSheet(compareAppPermissions, argv, profileComparisonWorkBook);
	  
	  //Part 5 Compare Apex Class Access
	  Set<String> setApexClassAccessProf1 = retrieveApexClassAccess(doc1);
	  Set<String> setApexClassAccessProf2 = retrieveApexClassAccess(doc2);
	  Integer numberOfDifferencesApex = compareApexClassAccessAndCreateSheet(setApexClassAccessProf1, setApexClassAccessProf2, argv, profileComparisonWorkBook);

	  //Part 6 Compare VF Page Access
	  Set<String> setVFPageAccessProf1 = retrieveVFPageAccess(doc1);
	  Set<String> setVFPageAccessProf2 = retrieveVFPageAccess(doc2);
	  Integer numberOfDifferencesVF = compareVFPageAccessAndCreateSheet(setVFPageAccessProf1, setVFPageAccessProf2, argv, profileComparisonWorkBook);
	  
	  //Compare Record Type Visibility
	  Map<String, recordTypeVisibilities> mapRecordTypePermissionsProf1  = retrieveRecordTypeVisibilities(doc1);
	  Map<String, recordTypeVisibilities> mapRecordTypePermissionsProf2  = retrieveRecordTypeVisibilities(doc2);
	  Map<String, List<recordTypeVisibilities>> compareRecordTypePermissions = compareRecordTypeAccess(mapRecordTypePermissionsProf1, mapRecordTypePermissionsProf2);
	  //System.out.print("compareRecordTypePermissions = " + compareRecordTypePermissions);
	  GenerateRecordTypeAccessSheet(compareRecordTypePermissions, argv, profileComparisonWorkBook);
	  
	  //Compare Tab Visibility
	  Map<String, tabVisibilities> mapTabPermissionsProf1  = retrieveTabVisibilities(doc1);
	  Map<String, tabVisibilities> mapTabPermissionsProf2  = retrieveTabVisibilities(doc2);
	  Map<String, List<tabVisibilities>> compareTabPermissions = compareTabAccess(mapTabPermissionsProf1, mapTabPermissionsProf2);
	  //System.out.print("compareRecordTypePermissions = " + compareRecordTypePermissions);
	  GenerateTabAccessSheet(compareTabPermissions, argv, profileComparisonWorkBook);	  
	  
	  //Compare User Permissions
	  Map<String, userPermissions> mapUserPermissionsProf1  = retrieveUserPermissions(doc1);
	  Map<String, userPermissions> mapUserPermissionsProf2  = retrieveUserPermissions(doc2);
	  Map<String, List<userPermissions>> compareUserPermissions = compareUserPermissions(mapUserPermissionsProf1, mapUserPermissionsProf2);
	  //System.out.print("compareRecordTypePermissions = " + compareRecordTypePermissions);
	  GenerateUserPermAccessSheet(compareUserPermissions, argv, profileComparisonWorkBook, "User Permissions");	 	

	  //Compare Custom Permissions
	  Map<String, userPermissions> mapCustomPermissionsProf1  = retrieveCustomPermissions(doc1);
	  Map<String, userPermissions> mapCustomPermissionsProf2  = retrieveCustomPermissions(doc2);
	  Map<String, List<userPermissions>> compareCustomPermissions = compareUserPermissions(mapCustomPermissionsProf1, mapCustomPermissionsProf2);
	  //System.out.print("compareRecordTypePermissions = " + compareRecordTypePermissions);
	  GenerateUserPermAccessSheet(compareCustomPermissions, argv, profileComparisonWorkBook, "Custom Permissions");	 	
	  
	  
	  //External Data Source Access
	  Map<String, externalDataSourceAccesses> mapExtDataSourcePermProf1  = retrieveExternalDataSourceAccess(doc1);
	  Map<String, externalDataSourceAccesses> mapExtDataSourcePermProf2  = retrieveExternalDataSourceAccess(doc2);
	  Map<String, List<externalDataSourceAccesses>> compareExtDataSourcePermissions = compareExtDataPerms(mapExtDataSourcePermProf1, mapExtDataSourcePermProf2);
	  //System.out.print("compareRecordTypePermissions = " + compareRecordTypePermissions);
	  GenerateExtDataSourcePermAccessSheet(compareExtDataSourcePermissions, argv, profileComparisonWorkBook);	 	  
	  
	  //Generate Comparison Summary
	  generateSummary(profileComparisonWorkBook, comparisonReportFLS.size(), mapLayoutComparisonReport.size(), compareObjectPermissions.size(), compareAppPermissions.size(),
			          numberOfDifferencesApex, numberOfDifferencesVF, compareRecordTypePermissions.size(), compareTabPermissions.size(), compareUserPermissions.size()
			          ,compareCustomPermissions.size(), compareExtDataSourcePermissions.size());
	  
	  // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream("ProfileComparison.xls");
	    profileComparisonWorkBook.write(fileOut);
	    fileOut.close();		  
	    System.out.println("Report Generated Successfully");
	  }catch(Exception e){
		  
		  e.printStackTrace();
	  }
  }

  private static Document parseFile(String fileName){
		Document doc = null;
	  try{
			File fXmlFile = new File(fileName);
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();
	  }catch(Exception e){
		  e.printStackTrace();
	  }	
	  return doc;
  }
  
  //This method Generates map of field permissions. 
  private static Map<String,List<String>> retrieveFieldPermissions(Document doc){
	  Map<String,List<String>> mapFieldPermissions = new HashMap<String, List<String>>(); 
	  NodeList nList = doc.getElementsByTagName("fieldPermissions");
		
		for (int temp = 0; temp < nList.getLength(); temp++) {

			Node nNode = nList.item(temp);
					
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				List<String> lstFieldPermission = new ArrayList<String>();
				
				Element eElement = (Element) nNode;
				lstFieldPermission.add(eElement.getElementsByTagName("readable").item(0).getTextContent());
				lstFieldPermission.add(eElement.getElementsByTagName("editable").item(0).getTextContent());
				mapFieldPermissions.put(eElement.getElementsByTagName("field").item(0).getTextContent(), lstFieldPermission);
			}
		}	  
	  return mapFieldPermissions;
  }
  
  private static Map<String,String> retrievePageLayoutAssignment(Document doc){
	  	Map<String,String> mapLayoutAssignments = new HashMap<String,String>();
		NodeList layoutAssignments = doc.getElementsByTagName("layoutAssignments");
		for (int temp = 0; temp < layoutAssignments.getLength(); temp++) {

			Node nNode = layoutAssignments.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				if(eElement.getElementsByTagName("recordType").item(0)==null){
					String[] layout = eElement.getElementsByTagName("layout").item(0).getTextContent().split("-");
					mapLayoutAssignments.put(layout[0], eElement.getElementsByTagName("layout").item(0).getTextContent());
				}else{
					String layoutType = eElement.getElementsByTagName("layout").item(0).getTextContent().split("-")[0];
					String key = eElement.getElementsByTagName("recordType").item(0).getTextContent();
					if(layoutType.equals("CaseClose")){
						key = key + "CaseClose"; //Added just to distinguish Close Case layouts from non close case. 
					}
					mapLayoutAssignments.put(key, eElement.getElementsByTagName("layout").item(0).getTextContent());
				}
			}
		}	
		return mapLayoutAssignments;
  }
  
  private static Map<String,objectPermissions> retrieveObjectPermissions(Document doc){
	  	Map<String, objectPermissions> mapObjectPermissions = new HashMap<String, objectPermissions>();
		NodeList objectPermissions = doc.getElementsByTagName("objectPermissions");
		for (int temp = 0; temp < objectPermissions.getLength(); temp++) {

			Node nNode = objectPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				objectPermissions objPermission = new objectPermissions();
				objPermission.object = eElement.getElementsByTagName("object").item(0).getTextContent();
				objPermission.allowCreate = eElement.getElementsByTagName("allowCreate").item(0).getTextContent();
				objPermission.allowEdit = eElement.getElementsByTagName("allowEdit").item(0).getTextContent();
				objPermission.allowRead = eElement.getElementsByTagName("allowRead").item(0).getTextContent();
				objPermission.allowDelete = eElement.getElementsByTagName("allowDelete").item(0).getTextContent();
				objPermission.modifyAllRecords = eElement.getElementsByTagName("modifyAllRecords").item(0).getTextContent();
				objPermission.viewAllRecords = eElement.getElementsByTagName("viewAllRecords").item(0).getTextContent();
				mapObjectPermissions.put(objPermission.object, objPermission);
			}
		}	
		return mapObjectPermissions;	  
  }
  
  private static Map<String, applicationVisibilities> retrieveApplicationVisibilities(Document doc){
	  Map<String, applicationVisibilities> mapAppPermissions = new HashMap<String, applicationVisibilities>();
	  NodeList applicationPermissions = doc.getElementsByTagName("applicationVisibilities");
		for (int temp = 0; temp < applicationPermissions.getLength(); temp++) {

			Node nNode = applicationPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				applicationVisibilities appVisibility = new applicationVisibilities();
				appVisibility.isdefault = eElement.getElementsByTagName("default").item(0).getTextContent();
				appVisibility.isVisible = eElement.getElementsByTagName("visible").item(0).getTextContent();
				mapAppPermissions.put(eElement.getElementsByTagName("application").item(0).getTextContent(), appVisibility);
			}
		}		  
	  return mapAppPermissions;
  }
  
  private static Set<String> retrieveApexClassAccess(Document doc){
	  Set<String> setApexClassAccess = new HashSet<String>(); 
	  NodeList classAccess = doc.getElementsByTagName("classAccesses");
		for (int temp = 0; temp < classAccess.getLength(); temp++) {

			Node nNode = classAccess.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				if(eElement.getElementsByTagName("enabled").item(0).getTextContent().equalsIgnoreCase("TRUE")){
					setApexClassAccess.add(eElement.getElementsByTagName("apexClass").item(0).getTextContent());
				}
			}
		}		  
	  return setApexClassAccess;
  }  
  
  private static Set<String> retrieveVFPageAccess(Document doc){
	  Set<String> setVFPageAccess = new HashSet<String>(); 
	  NodeList pageAccess = doc.getElementsByTagName("pageAccesses");
		for (int temp = 0; temp < pageAccess.getLength(); temp++) {

			Node nNode = pageAccess.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				if(eElement.getElementsByTagName("enabled").item(0).getTextContent().equalsIgnoreCase("TRUE")){
					setVFPageAccess.add(eElement.getElementsByTagName("apexPage").item(0).getTextContent());
				}
			}
		}		  
	  return setVFPageAccess;
  }  
    
  private static Map<String, recordTypeVisibilities> retrieveRecordTypeVisibilities(Document doc){
	  Map<String, recordTypeVisibilities> mapAppPermissions = new HashMap<String, recordTypeVisibilities>();
	  NodeList applicationPermissions = doc.getElementsByTagName("recordTypeVisibilities");
		for (int temp = 0; temp < applicationPermissions.getLength(); temp++) {

			Node nNode = applicationPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				recordTypeVisibilities appVisibility = new recordTypeVisibilities();
				appVisibility.isdefault = eElement.getElementsByTagName("default").item(0).getTextContent();
				appVisibility.isVisible = eElement.getElementsByTagName("visible").item(0).getTextContent();
				mapAppPermissions.put(eElement.getElementsByTagName("recordType").item(0).getTextContent(), appVisibility);
			}
		}		  
	  return mapAppPermissions;
  }  
  
  private static Map<String, tabVisibilities> retrieveTabVisibilities(Document doc){
	  Map<String, tabVisibilities> mapTabPermissions = new HashMap<String, tabVisibilities>();
	  NodeList tabPermissions = doc.getElementsByTagName("tabVisibilities");
		for (int temp = 0; temp < tabPermissions.getLength(); temp++) {

			Node nNode = tabPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				tabVisibilities tabVisibility = new tabVisibilities();
				tabVisibility.visibility = eElement.getElementsByTagName("visibility").item(0).getTextContent();
				mapTabPermissions.put(eElement.getElementsByTagName("tab").item(0).getTextContent(), tabVisibility);
			}
		}		  
	  return mapTabPermissions;
  }    
  
  private static Map<String, userPermissions> retrieveUserPermissions(Document doc){
	  Map<String, userPermissions> mapUserPermissions = new HashMap<String, userPermissions>();
	  NodeList userPermissions = doc.getElementsByTagName("userPermissions");
		for (int temp = 0; temp < userPermissions.getLength(); temp++) {

			Node nNode = userPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				userPermissions userPerm = new userPermissions();
				userPerm.enabled = eElement.getElementsByTagName("enabled").item(0).getTextContent();
				mapUserPermissions.put(eElement.getElementsByTagName("name").item(0).getTextContent(), userPerm);
			}
		}		  
	  return mapUserPermissions;
  }
  
  private static Map<String, userPermissions> retrieveCustomPermissions(Document doc){
	  Map<String, userPermissions> mapUserPermissions = new HashMap<String, userPermissions>();
	  NodeList userPermissions = doc.getElementsByTagName("customPermissions");
		for (int temp = 0; temp < userPermissions.getLength(); temp++) {

			Node nNode = userPermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				userPermissions userPerm = new userPermissions();
				userPerm.enabled = eElement.getElementsByTagName("enabled").item(0).getTextContent();
				mapUserPermissions.put(eElement.getElementsByTagName("name").item(0).getTextContent(), userPerm);
			}
		}		  
	  return mapUserPermissions;
  }  
  
  private static Map<String, externalDataSourceAccesses> retrieveExternalDataSourceAccess(Document doc){
	  Map<String, externalDataSourceAccesses> mapExtdataAccessPermissions = new HashMap<String, externalDataSourceAccesses>();
	  NodeList extDataSourcePermissions = doc.getElementsByTagName("externalDataSourceAccesses");
		for (int temp = 0; temp < extDataSourcePermissions.getLength(); temp++) {

			Node nNode = extDataSourcePermissions.item(temp);
			if (nNode.getNodeType() == Node.ELEMENT_NODE) {
				Element eElement = (Element) nNode;
				externalDataSourceAccesses extDataSourcePerm = new externalDataSourceAccesses();
				extDataSourcePerm.enabled = eElement.getElementsByTagName("enabled").item(0).getTextContent();
				mapExtdataAccessPermissions.put(eElement.getElementsByTagName("externalDataSource").item(0).getTextContent(), extDataSourcePerm);
			}
		}		  
	  return mapExtdataAccessPermissions;
  }  
  
  private static Map<String,List<String>> compareFLS(Map<String,List<String>> mapFieldPermissionsProfile1, Map<String,List<String>> mapFieldPermissionsProfile2){
	  Map<String,List<String>> comparisonReport = new HashMap<String,List<String>>();
	  for(String key:mapFieldPermissionsProfile1.keySet()){
		  List<String> prof1Permission = mapFieldPermissionsProfile1.get(key);
		  List<String> prof2Permission = mapFieldPermissionsProfile2.get(key);

		  if(prof2Permission==null || !prof1Permission.get(0).equalsIgnoreCase(prof2Permission.get(0)) || !prof1Permission.get(1).equalsIgnoreCase(prof2Permission.get(1))){
			  List<String> combinedPermissions = new ArrayList<String>();
			  if(prof2Permission==null){
				  prof2Permission = new ArrayList<String>();
				  prof2Permission.add("false");
				  prof2Permission.add("false");
			  }
			  
			  combinedPermissions.addAll(prof1Permission);
			  combinedPermissions.addAll(prof2Permission);
			  comparisonReport.put(key, combinedPermissions);
		  }
	  }
	  return comparisonReport;
  }
  
  private static Map<String,List<String>> compareLayoutAssignments(Map<String,String> mapPageLayoutAssignmentProf1, Map<String,String> mapPageLayoutAssignmentProf2){
	  Map<String,List<String>> ComparisonReport = new HashMap<String,List<String>>();
	  for(String key:mapPageLayoutAssignmentProf1.keySet()){
		  String layout1 = mapPageLayoutAssignmentProf1.get(key);
		  String layout2 = mapPageLayoutAssignmentProf2.get(key);
		  if(layout2==null || !layout1.equalsIgnoreCase(layout2)){
			  List<String> lstLayouts = new ArrayList<String>();
			  lstLayouts.add(layout1);
			  if(layout2==null)layout2 = "Not Assigned";
			  lstLayouts.add(layout2);
			  ComparisonReport.put(key,lstLayouts);
		  }
	  }
	  //Below code for handling non assigned cases
	  
	  for(String key:mapPageLayoutAssignmentProf2.keySet()){
		  String layout1 = mapPageLayoutAssignmentProf1.get(key);
		  String layout2 = mapPageLayoutAssignmentProf2.get(key);
		  //System.out.println("Key = " + key + "   Layout1="+ layout1 + "   Layout2 =" + layout2);
		  if(layout1==null || !layout1.equalsIgnoreCase(layout2)){
			  List<String> lstLayouts = new ArrayList<String>();
			  if(layout1==null)layout1 = "Not Assigned";
			  lstLayouts.add(layout1);
			  lstLayouts.add(layout2);
			  ComparisonReport.put(key,lstLayouts);
		  }
	  }
	  
	  return ComparisonReport;
  }
  
  private static  Map<String, List<objectPermissions>> compareObjectPermissions(Map<String, objectPermissions> mapObjectPermissionsProf1, Map<String, objectPermissions> mapObjectPermissionsProf2){
	  Map<String, List<objectPermissions>> comparisonReport = new HashMap<String, List<objectPermissions>>();
	  for(String key:mapObjectPermissionsProf1.keySet()){
		  objectPermissions objPermProf1 = mapObjectPermissionsProf1.get(key);
		  objectPermissions objPermProf2 = mapObjectPermissionsProf2.get(key);
		  if(objPermProf2==null || (!objPermProf1.allowCreate.equalsIgnoreCase(objPermProf2.allowCreate) 
		   || !objPermProf1.allowEdit.equalsIgnoreCase(objPermProf2.allowEdit)
		   || !objPermProf1.allowDelete.equalsIgnoreCase(objPermProf2.allowDelete)
		   || !objPermProf1.allowRead.equalsIgnoreCase(objPermProf2.allowRead)
		   || !objPermProf1.modifyAllRecords.equalsIgnoreCase(objPermProf2.modifyAllRecords)
		   || !objPermProf1.viewAllRecords.equalsIgnoreCase(objPermProf2.viewAllRecords))){
			  List<objectPermissions> lstObjectPermissions = new ArrayList<objectPermissions>();
			  lstObjectPermissions.add(objPermProf1);
			  if(objPermProf2==null)objPermProf2 = new objectPermissions();
			  lstObjectPermissions.add(objPermProf2);
			  comparisonReport.put(key, lstObjectPermissions);
		  }
	  }
	  for(String key:mapObjectPermissionsProf2.keySet()){
		  objectPermissions objPermProf1 = mapObjectPermissionsProf1.get(key);
		  objectPermissions objPermProf2 = mapObjectPermissionsProf2.get(key);
		  if(objPermProf1==null){
			  List<objectPermissions> lstObjectPermissions = new ArrayList<objectPermissions>();
			  objPermProf1 = new objectPermissions();
			  lstObjectPermissions.add(objPermProf1);
			  lstObjectPermissions.add(objPermProf2);
			  comparisonReport.put(key, lstObjectPermissions);
		  }
	  }
	  return comparisonReport;
  }
  
  private static  Map<String, List<applicationVisibilities>> compareApplicationVisibility(Map<String, applicationVisibilities> mapAppPermissionsProf1, Map<String, applicationVisibilities> mapAppPermissionsProf2){
	  Map<String, List<applicationVisibilities>> comparisonReport = new HashMap<String, List<applicationVisibilities>>();
	  for(String key:mapAppPermissionsProf1.keySet()){
		  applicationVisibilities appPermProf1 = mapAppPermissionsProf1.get(key);
		  applicationVisibilities appPermProf2 = mapAppPermissionsProf2.get(key);
		  if(appPermProf2==null || !appPermProf1.isdefault.equalsIgnoreCase(appPermProf2.isdefault) 
		   || !appPermProf1.isVisible.equalsIgnoreCase(appPermProf2.isVisible)){
			  if(appPermProf2==null)appPermProf2 = new applicationVisibilities();
			  List<applicationVisibilities> lstAppPermissions = new ArrayList<applicationVisibilities>();
			  lstAppPermissions.add(appPermProf1);
			  lstAppPermissions.add(appPermProf2);
			  comparisonReport.put(key, lstAppPermissions);
		  }
	  }
	  return comparisonReport;
  }  
  
  private static int compareApexClassAccessAndCreateSheet(Set<String> setApexClassAccessProf1, Set<String> setApexClassAccessProf2, String argv[], Workbook profileComparisonWorkBook){
	  Set<String> setExtraClassesProf1 = new HashSet<String>();
	  Set<String> setExtraClassesProf2 = new HashSet<String>();
	  Integer numberOfRows = 0;
	  for(String className:setApexClassAccessProf1){
		  if(!setApexClassAccessProf2.contains(className)){
			  setExtraClassesProf1.add(className);
		  }
	  }
	  for(String className:setApexClassAccessProf2){
		  if(!setApexClassAccessProf1.contains(className)){
			  setExtraClassesProf2.add(className);
		  }
	  }
	  Sheet apexSheet = profileComparisonWorkBook.createSheet("Apex Class Comparison");		
		try {
			Row row = apexSheet.createRow((short)0);
			row.createCell(0).setCellValue(argv[0]);
			row.createCell(1).setCellValue(argv[1]);
			Integer i=1;
			//Write a new student object list to the CSV file
			for (String key:setExtraClassesProf1) {
				Row valueRow = apexSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				i++;
			}
			i=1;
			for (String key:setExtraClassesProf2) {
				Row valueRow = apexSheet.getRow(i);
				if(valueRow==null)valueRow = apexSheet.createRow(i);
				valueRow.createCell(1).setCellValue(key);
				i++;
			}			
			numberOfRows = apexSheet.getLastRowNum();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}	 
		return numberOfRows;
  }
  
  private static Integer compareVFPageAccessAndCreateSheet(Set<String> setVFPageAccessProf1, Set<String> setVFPageAccessProf2, String argv[], Workbook profileComparisonWorkBook){
	  Set<String> setExtraPagesProf1 = new HashSet<String>();
	  Set<String> setExtraPagesProf2 = new HashSet<String>();
	  Integer numberOfRows = 0;
	  for(String pageName:setVFPageAccessProf1){
		  if(!setVFPageAccessProf2.contains(pageName)){
			  setExtraPagesProf1.add(pageName);
		  }
	  }
	  for(String pageName:setVFPageAccessProf2){
		  if(!setVFPageAccessProf1.contains(pageName)){
			  setExtraPagesProf2.add(pageName);
		  }
	  }
	  Sheet VFSheet = profileComparisonWorkBook.createSheet("VF Page Comparison");		
		try {
			Row row = VFSheet.createRow((short)0);
			row.createCell(0).setCellValue(argv[0]);
			row.createCell(1).setCellValue(argv[1]);
			Integer i=1;
			//Write a new student object list to the CSV file
			for (String key:setExtraPagesProf1) {
				Row valueRow = VFSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				i++;
			}
			i=1;
			for (String key:setExtraPagesProf2) {
				Row valueRow = VFSheet.getRow(i);
				if(valueRow==null)valueRow = VFSheet.createRow(i);
				valueRow.createCell(1).setCellValue(key);
				i++;
			}			
			numberOfRows = VFSheet.getLastRowNum();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
		return numberOfRows;
  }  
  
  private static  Map<String, List<recordTypeVisibilities>> compareRecordTypeAccess(Map<String, recordTypeVisibilities> mapRecordTypePermissionsProf1, Map<String, recordTypeVisibilities> mapRecordTypePermissionsProf2){
	  Map<String, List<recordTypeVisibilities>> comparisonReport = new HashMap<String, List<recordTypeVisibilities>>();
	  for(String key:mapRecordTypePermissionsProf1.keySet()){
		  recordTypeVisibilities appPermProf1 = mapRecordTypePermissionsProf1.get(key);
		  recordTypeVisibilities appPermProf2 = mapRecordTypePermissionsProf2.get(key);
		  
		  if(!appPermProf1.isdefault.equalsIgnoreCase(appPermProf2.isdefault) 
		   || !appPermProf1.isVisible.equalsIgnoreCase(appPermProf2.isVisible)){
			  List<recordTypeVisibilities> lstAppPermissions = new ArrayList<recordTypeVisibilities>();
			  lstAppPermissions.add(appPermProf1);
			  lstAppPermissions.add(appPermProf2);
			  comparisonReport.put(key, lstAppPermissions);
		  }
	  }
	  return comparisonReport;
  }  
  
  private static  Map<String, List<tabVisibilities>> compareTabAccess(Map<String, tabVisibilities> mapTabPermissionsProf1, Map<String, tabVisibilities> mapTabPermissionsProf2){
	  Map<String, List<tabVisibilities>> comparisonReport = new HashMap<String, List<tabVisibilities>>();
	  for(String key:mapTabPermissionsProf1.keySet()){
		  tabVisibilities tabPermProf1 = mapTabPermissionsProf1.get(key);
		  tabVisibilities tabPermProf2 = mapTabPermissionsProf2.get(key);
		  if(tabPermProf2==null || !tabPermProf1.visibility.equalsIgnoreCase(tabPermProf2.visibility)){
			  List<tabVisibilities> lstTabPermissions = new ArrayList<tabVisibilities>();
			  if(tabPermProf2==null)tabPermProf2 = new tabVisibilities();
			  lstTabPermissions.add(tabPermProf1);
			  lstTabPermissions.add(tabPermProf2);
			  comparisonReport.put(key, lstTabPermissions);
		  }
	  }
	  for(String key:mapTabPermissionsProf2.keySet()){
		  tabVisibilities tabPermProf1 = mapTabPermissionsProf1.get(key);
		  tabVisibilities tabPermProf2 = mapTabPermissionsProf2.get(key);
		  if(tabPermProf1==null){
			  List<tabVisibilities> lstTabPermissions = new ArrayList<tabVisibilities>();
			  tabPermProf1 = new tabVisibilities();
			  lstTabPermissions.add(tabPermProf1);
			  lstTabPermissions.add(tabPermProf2);
			  comparisonReport.put(key, lstTabPermissions);
		  }
	  }	  
	  return comparisonReport;
  }    
  
  private static  Map<String, List<userPermissions>> compareUserPermissions(Map<String, userPermissions> mapUserPermissionsProf1, Map<String, userPermissions> mapUserPermissionsProf2){
	  Map<String, List<userPermissions>> comparisonReport = new HashMap<String, List<userPermissions>>();
	  for(String key:mapUserPermissionsProf1.keySet()){
		  userPermissions userPermProf1 = mapUserPermissionsProf1.get(key);
		  userPermissions userPermProf2 = mapUserPermissionsProf2.get(key);
  
		  if(userPermProf2==null || !userPermProf1.enabled.equalsIgnoreCase(userPermProf2.enabled)){
			  List<userPermissions> lstUserPermissions = new ArrayList<userPermissions>();
			  if(userPermProf2==null)userPermProf2 = new userPermissions();
			  lstUserPermissions.add(userPermProf1);
			  lstUserPermissions.add(userPermProf2);
			  comparisonReport.put(key, lstUserPermissions);
		  }
	  }
	  for(String key:mapUserPermissionsProf2.keySet()){
		  userPermissions userPermProf1 = mapUserPermissionsProf1.get(key);
		  userPermissions userPermProf2 = mapUserPermissionsProf2.get(key);
		  if(userPermProf1==null){
			  List<userPermissions> lstUserPermissions = new ArrayList<userPermissions>();
			  userPermProf1 = new userPermissions();
			  lstUserPermissions.add(userPermProf1);
			  lstUserPermissions.add(userPermProf2);
			  comparisonReport.put(key, lstUserPermissions);
		  }
	  }	  
	  return comparisonReport;
  }   
  
  private static  Map<String, List<externalDataSourceAccesses>> compareExtDataPerms(Map<String, externalDataSourceAccesses> mapExtDataSourcePermProf1, Map<String, externalDataSourceAccesses> mapExtDataSourcePermProf2){
	  Map<String, List<externalDataSourceAccesses>> comparisonReport = new HashMap<String, List<externalDataSourceAccesses>>();
	  for(String key:mapExtDataSourcePermProf1.keySet()){
		  externalDataSourceAccesses ExtDataSourcePermProf1 = mapExtDataSourcePermProf1.get(key);
		  externalDataSourceAccesses ExtDataSourcePermProf2 = mapExtDataSourcePermProf2.get(key);

		  if(ExtDataSourcePermProf2==null || !ExtDataSourcePermProf1.enabled.equalsIgnoreCase(ExtDataSourcePermProf2.enabled)){
			  List<externalDataSourceAccesses> lstUserPermissions = new ArrayList<externalDataSourceAccesses>();
			  if(ExtDataSourcePermProf2==null)ExtDataSourcePermProf2 = new externalDataSourceAccesses();
			  lstUserPermissions.add(ExtDataSourcePermProf1);
			  lstUserPermissions.add(ExtDataSourcePermProf2);
			  comparisonReport.put(key, lstUserPermissions);
		  }
	  }
	  for(String key:mapExtDataSourcePermProf2.keySet()){
		  externalDataSourceAccesses ExtDataSourcePermProf1 = mapExtDataSourcePermProf1.get(key);
		  externalDataSourceAccesses ExtDataSourcePermProf2 = mapExtDataSourcePermProf2.get(key);
		  if(ExtDataSourcePermProf1==null){
			  List<externalDataSourceAccesses> lstUserPermissions = new ArrayList<externalDataSourceAccesses>();
			  ExtDataSourcePermProf1 = new externalDataSourceAccesses();
			  lstUserPermissions.add(ExtDataSourcePermProf1);
			  lstUserPermissions.add(ExtDataSourcePermProf2);
			  comparisonReport.put(key, lstUserPermissions);
		  }
	  }	  
	  return comparisonReport;
  }     
  
  private static void GenerateFLSSheet(Map<String,List<String>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		//CSV file header
	  Sheet flsSheet = profileComparisonWorkBook.createSheet("FLS Comparison");		
		try {
			Row row = flsSheet.createRow((short)0);
			row.createCell(0).setCellValue("Object");
			row.createCell(1).setCellValue("Field");
			row.createCell(2).setCellValue("Profile");
			row.createCell(3).setCellValue("Visible");
			row.createCell(4).setCellValue("Editable");
			row.createCell(5).setCellValue("Profile");
			row.createCell(6).setCellValue("Visible");
			row.createCell(7).setCellValue("Editable");
			Integer i=1;
			//Write a new student object list to the CSV file
			for (String key:comparisonReport.keySet()) {
				Row valueRow = flsSheet.createRow(i);

				String[] objectField = key.split("\\.");
				valueRow.createCell(0).setCellValue(objectField[0]);
				valueRow.createCell(1).setCellValue(objectField[1]);
				valueRow.createCell(2).setCellValue(argv[0]);
				List<String> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(3).setCellValue(lstPermissions.get(0));
				valueRow.createCell(4).setCellValue(lstPermissions.get(1));
				valueRow.createCell(5).setCellValue(argv[1]);
				valueRow.createCell(6).setCellValue(lstPermissions.get(2));
				valueRow.createCell(7).setCellValue(lstPermissions.get(3));
				i++;
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
  }		

  private static void GenerateLayoutSheet(Map<String,List<String>> mapLayoutComparisonReport, Workbook profileComparisonWorkBook, String argv[]){
		try {
			Sheet layoutSheet = profileComparisonWorkBook.createSheet("Layout Assignment");
			Row headerrow = layoutSheet.createRow((short)0);
			headerrow.createCell(0).setCellValue("Object");
			headerrow.createCell(1).setCellValue("Record Type");
			headerrow.createCell(2).setCellValue("Profile 1");
			headerrow.createCell(3).setCellValue("Layout");
			headerrow.createCell(4).setCellValue("Profile 2");
			headerrow.createCell(5).setCellValue("Layout");
			Integer i=1;
			//Write a new student object list to the CSV file
			for (String key : mapLayoutComparisonReport.keySet()) {
				Row valueRow = layoutSheet.createRow(i);
				String[] objectField = key.split("\\.");
				valueRow.createCell(0).setCellValue(objectField[0]);
				if(objectField.length>1){
					valueRow.createCell(1).setCellValue(objectField[1]);
				}else{
					valueRow.createCell(1).setCellValue("Master");				
				}
				valueRow.createCell(2).setCellValue(argv[0]);
				List<String> layouts = mapLayoutComparisonReport.get(key);
				valueRow.createCell(3).setCellValue(layouts.get(0));
				valueRow.createCell(4).setCellValue(argv[1]);
				valueRow.createCell(5).setCellValue(layouts.get(1));
				i++;
			}
			
		} catch (Exception e) {

		} finally {
		}
  }		
  
  private static void GenerateObjectPermissionSheet(Map<String,List<objectPermissions>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet("Object Permissions");			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Object");
			row.createCell(1).setCellValue("Profile 1");
			row.createCell(2).setCellValue("allowCreate");
			row.createCell(3).setCellValue("allowDelete");
			row.createCell(4).setCellValue("allowEdit");
			row.createCell(5).setCellValue("allowRead");
			row.createCell(6).setCellValue("modifyAllRecords");
			row.createCell(7).setCellValue("viewAllRecords");			
			row.createCell(8).setCellValue("Profile 2");
			row.createCell(9).setCellValue("allowCreate");
			row.createCell(10).setCellValue("allowDelete");
			row.createCell(11).setCellValue("allowEdit");
			row.createCell(12).setCellValue("allowRead");
			row.createCell(13).setCellValue("modifyAllRecords");
			row.createCell(14).setCellValue("viewAllRecords");			

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				valueRow.createCell(1).setCellValue(argv[0]);
				List<objectPermissions> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(2).setCellValue(lstPermissions.get(0).allowCreate);
				valueRow.createCell(3).setCellValue(lstPermissions.get(0).allowDelete);
				valueRow.createCell(4).setCellValue(lstPermissions.get(0).allowEdit);
				valueRow.createCell(5).setCellValue(lstPermissions.get(0).allowRead);
				valueRow.createCell(6).setCellValue(lstPermissions.get(0).modifyAllRecords);
				valueRow.createCell(7).setCellValue(lstPermissions.get(0).viewAllRecords);
				valueRow.createCell(8).setCellValue(argv[1]);
				valueRow.createCell(9).setCellValue(lstPermissions.get(1).allowCreate);
				valueRow.createCell(10).setCellValue(lstPermissions.get(1).allowDelete);
				valueRow.createCell(11).setCellValue(lstPermissions.get(1).allowEdit);
				valueRow.createCell(12).setCellValue(lstPermissions.get(1).allowRead);
				valueRow.createCell(13).setCellValue(lstPermissions.get(1).modifyAllRecords);
				valueRow.createCell(14).setCellValue(lstPermissions.get(1).viewAllRecords);
				i++;
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
  }		  
  
  private static void GenerateAppPermissionSheet(Map<String,List<applicationVisibilities>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet("App Permissions");			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Application");
			row.createCell(1).setCellValue("Profile 1");
			row.createCell(2).setCellValue("Default");
			row.createCell(3).setCellValue("Visible");
			row.createCell(4).setCellValue("Profile 2");
			row.createCell(5).setCellValue("Default");
			row.createCell(6).setCellValue("Visible");

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				valueRow.createCell(1).setCellValue(argv[0]);
				List<applicationVisibilities> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(2).setCellValue(lstPermissions.get(0).isdefault);
				valueRow.createCell(3).setCellValue(lstPermissions.get(0).isVisible);
				valueRow.createCell(4).setCellValue(argv[1]);
				valueRow.createCell(5).setCellValue(lstPermissions.get(1).isdefault);
				valueRow.createCell(6).setCellValue(lstPermissions.get(1).isVisible);
				i++;
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
  }	  
  
  private static void GenerateRecordTypeAccessSheet(Map<String,List<recordTypeVisibilities>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet("Record Type Permissions");			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Object");
			row.createCell(1).setCellValue("Record Type");
			row.createCell(2).setCellValue("Profile 1");
			row.createCell(3).setCellValue("Default");
			row.createCell(4).setCellValue("Visible");
			row.createCell(5).setCellValue("Profile 2");
			row.createCell(6).setCellValue("Default");
			row.createCell(7).setCellValue("Visible");

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				String[] objRecordtype = key.split("\\."); 
				valueRow.createCell(0).setCellValue(objRecordtype[0]);
				valueRow.createCell(1).setCellValue(objRecordtype[1]);
				valueRow.createCell(2).setCellValue(argv[0]);
				List<recordTypeVisibilities> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(3).setCellValue(lstPermissions.get(0).isdefault);
				valueRow.createCell(4).setCellValue(lstPermissions.get(0).isVisible);
				valueRow.createCell(5).setCellValue(argv[1]);
				valueRow.createCell(6).setCellValue(lstPermissions.get(1).isdefault);
				valueRow.createCell(7).setCellValue(lstPermissions.get(1).isVisible);
				i++;
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
  }	  
  
  private static void GenerateTabAccessSheet(Map<String,List<tabVisibilities>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet("Tab Permissions");			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Tab");
			row.createCell(1).setCellValue("Profile 1");
			row.createCell(2).setCellValue("Visibility");
			row.createCell(3).setCellValue("Profile 2");
			row.createCell(4).setCellValue("Visibility");

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				List<tabVisibilities> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(1).setCellValue(argv[0]);
				valueRow.createCell(2).setCellValue(lstPermissions.get(0).visibility);
				valueRow.createCell(3).setCellValue(argv[1]);
				valueRow.createCell(4).setCellValue(lstPermissions.get(1).visibility);
				i++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
  }  
  
  private static void GenerateUserPermAccessSheet(Map<String,List<userPermissions>> comparisonReport, String argv[], Workbook profileComparisonWorkBook, String sheetName){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet(sheetName);			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Permission Name");
			row.createCell(1).setCellValue("Profile 1");
			row.createCell(2).setCellValue("Enabled");
			row.createCell(3).setCellValue("Profile 2");
			row.createCell(4).setCellValue("Enabled");

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				List<userPermissions> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(1).setCellValue(argv[0]);
				valueRow.createCell(2).setCellValue(lstPermissions.get(0).enabled);
				valueRow.createCell(3).setCellValue(argv[1]);
				valueRow.createCell(4).setCellValue(lstPermissions.get(1).enabled);
				i++;
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
}    
  
  private static void GenerateExtDataSourcePermAccessSheet(Map<String,List<externalDataSourceAccesses>> comparisonReport, String argv[], Workbook profileComparisonWorkBook){
		try {
			Sheet objectSheet = profileComparisonWorkBook.createSheet("Ext Data Source");			
			Row row = objectSheet.createRow((short)0);
			row.createCell(0).setCellValue("Data Source Name");
			row.createCell(1).setCellValue("Profile 1");
			row.createCell(2).setCellValue("Enabled");
			row.createCell(3).setCellValue("Profile 2");
			row.createCell(4).setCellValue("Enabled");

			Integer i = 1;
			//Write a new student object list to the CSV file
			for (String key : comparisonReport.keySet()) {
				Row valueRow = objectSheet.createRow(i);
				valueRow.createCell(0).setCellValue(key);
				List<externalDataSourceAccesses> lstPermissions = comparisonReport.get(key);
				valueRow.createCell(1).setCellValue(argv[0]);
				valueRow.createCell(2).setCellValue(lstPermissions.get(0).enabled);
				valueRow.createCell(3).setCellValue(argv[1]);
				valueRow.createCell(4).setCellValue(lstPermissions.get(1).enabled);
				i++;
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
}     
  
  private static void generateSummary(Workbook profileComparisonWorkBook, Integer FLS, Integer layouts, Integer objects,
		  							  Integer Apps, Integer Apex, Integer VF, Integer RecordTypes, Integer Tabs, Integer UserPerms, Integer customPerms, Integer extDataSourcePerm){
	  
		try {
			Sheet summarySheet = profileComparisonWorkBook.createSheet("Summary");			
			Row row = summarySheet.createRow((short)0);
			row.createCell(0).setCellValue("Permission Type");
			row.createCell(1).setCellValue("Number Of Differences");
			Row FLSRow = summarySheet.createRow(1);
			FLSRow.createCell(0).setCellValue("FLS");
			FLSRow.createCell(1).setCellValue(FLS);
			Row layoutRow = summarySheet.createRow(2);
			layoutRow.createCell(0).setCellValue("Layouts");
			layoutRow.createCell(1).setCellValue(layouts);
			Row objectsRow = summarySheet.createRow(3);
			objectsRow.createCell(0).setCellValue("Objects");
			objectsRow.createCell(1).setCellValue(objects);
			Row appsRow = summarySheet.createRow(4);
			appsRow.createCell(0).setCellValue("Applications");
			appsRow.createCell(1).setCellValue(Apps);
			Row apexRow = summarySheet.createRow(5);
			apexRow.createCell(0).setCellValue("Apex Classes");
			apexRow.createCell(1).setCellValue(Apex);
			Row VFRow = summarySheet.createRow(6);
			VFRow.createCell(0).setCellValue("Visualforce Pages");
			VFRow.createCell(1).setCellValue(VF);
			Row recordTypeRow = summarySheet.createRow(7);
			recordTypeRow.createCell(0).setCellValue("Record Type");
			recordTypeRow.createCell(1).setCellValue(RecordTypes);
			Row tabsRow = summarySheet.createRow(8);
			tabsRow.createCell(0).setCellValue("Tabs");
			tabsRow.createCell(1).setCellValue(Tabs);
			Row userPermsRow = summarySheet.createRow(9);
			userPermsRow.createCell(0).setCellValue("User Permissions");
			userPermsRow.createCell(1).setCellValue(UserPerms);
			Row customPermsRow = summarySheet.createRow(10);
			customPermsRow.createCell(0).setCellValue("Custom Permissions");
			customPermsRow.createCell(1).setCellValue(customPerms);
			Row extDataSource = summarySheet.createRow(11);
			extDataSource.createCell(0).setCellValue("External Data Source Access");
			extDataSource.createCell(1).setCellValue(extDataSourcePerm);

			
			  profileComparisonWorkBook.setSheetOrder("Summary",0);
			  profileComparisonWorkBook.setActiveSheet(0);			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
	  
  }
  static class objectPermissions{
	  String allowCreate;
	  String allowDelete;
      String allowEdit;
      String allowRead;
      String modifyAllRecords;
      String object;
      String viewAllRecords; 
      objectPermissions(){
    	  allowCreate = "false";
    	  allowDelete = "false";
    	  allowEdit = "false";
    	  allowRead = "false";
    	  modifyAllRecords = "false";
    	  viewAllRecords = "false";
      }
  }
  
  static class applicationVisibilities{
	  String isdefault;
	  String isVisible;
	  applicationVisibilities(){
		  isdefault = "false";
		  isVisible = "false";
	  }
  }
  
  static class recordTypeVisibilities{
	  String isdefault;
	  String isVisible;
	  recordTypeVisibilities(){
		  isdefault = "false";
		  isVisible = "false";
	  }
  }  
  static class tabVisibilities{
	  String visibility;
	  tabVisibilities(){
		  visibility = "Tab Hidden";
	  }
  }  
  
  static class userPermissions{
	  String enabled;
	  userPermissions(){
		  enabled = "FALSE";
	  }
  }    
  
  static class externalDataSourceAccesses{
	  String enabled;
	  externalDataSourceAccesses(){
		  enabled = "FALSE";
	  }
  }    
  
}
