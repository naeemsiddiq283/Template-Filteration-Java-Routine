import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.vd.xtime.automation.AbstractComponent;

public class ReadLocalizationDataAndCreateMap extends AbstractComponent {
	protected static int mergeTermsRowCounter = 1;
	protected static int localizationStringsRowCounter = 1;
	protected static int okTemplates = 0;
	protected static int discrepantTemplates = 0;

	public static void readLocalizationDataAndCreateMap(String localizationFilePath, String glossaryFilePath)
			throws Exception {
		/* Reading Localization sheet */
		FileInputStream fileInputStream = new FileInputStream(new File(localizationFilePath));
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet localizationSheet = workbook.getSheetAt(2);
		XSSFSheet templateSheet = workbook.getSheetAt(1);
		fileInputStream.close();
		/**/

		/* Reading Glossary sheet for mergeterms */
		FileInputStream glossaryFileInputStream = new FileInputStream(new File(glossaryFilePath));
		XSSFWorkbook glossaryWorkbook = new XSSFWorkbook(glossaryFileInputStream);
		XSSFSheet mergeTermsSheet = glossaryWorkbook.getSheetAt(0);
		glossaryFileInputStream.close();

		/**/
		Map<String, Map<String, String>> tagMap = createTagMap(localizationSheet);

		Map<String, Map<String, Map<String, String>>> languageOuterMap = createLanguageMap(localizationSheet);

		Map<String, String> mergeTermsMap = createMergeTermsMap(mergeTermsSheet);
		replaceLocalizationTagWithValueAndN6MergeTermWithN7InTemplate(templateSheet, tagMap, languageOuterMap,
				mergeTermsMap);

		FileOutputStream outputStream = new FileOutputStream(localizationFilePath);
		workbook.write(outputStream);
		workbook.close();
		outputStream.flush();
		outputStream.close();

		glossaryWorkbook.close();

	}

	private static Map<String, String> createMergeTermsMap(XSSFSheet mergeTermsSheet) throws Exception {

		Map<String, String> mergeTermMap = new HashMap<String, String>();
		Iterator<Row> rowIterator = mergeTermsSheet.iterator();
		rowIterator.next();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Cell N6MergeTermCell = row.getCell(0);
			Cell N7MergeTermCell = row.getCell(1);
			String N6MergeTerm = "";
			String N7MergeTerm = "";
			if (N6MergeTermCell != null)
				N6MergeTerm = N6MergeTermCell.getStringCellValue().trim();
			if (N7MergeTermCell != null)
				N7MergeTerm = N7MergeTermCell.getStringCellValue().trim();
			if (!mergeTermMap.containsKey(N6MergeTerm))
				mergeTermMap.put(N6MergeTerm, N7MergeTerm);

		}
//		System.out.println("MergTerms Map Size: " + mergeTermMap.size() + "\nMergTerms Kap : " + mergeTermMap.get("$enginetype"));

		return mergeTermMap;
	}

	private static Map<String, Map<String, String>> createTagMap(XSSFSheet localizationSheet) {
		Map<String, Map<String, String>> tagMap = new HashMap<String, Map<String, String>>();
		Map<String, String> orgKeyMap;

		Iterator<Row> rowIterator = localizationSheet.iterator();
		rowIterator.next();
		int counter = 0;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			String tagName = row.getCell(6).getStringCellValue().trim();
			String value = row.getCell(7).getStringCellValue().trim();
			Cell orgKeyCell = row.getCell(10);
			String orgKey = "";
			if (orgKeyCell != null) {
				orgKey = orgKeyCell.getStringCellValue().trim();
			} else
				orgKey = "GLOBAL";

			if (tagMap.containsKey(tagName)) {
				orgKeyMap = tagMap.get(tagName);
				if (!orgKeyMap.containsKey(orgKey)) {
					orgKeyMap.put(orgKey, value);
					tagMap.put(tagName, orgKeyMap);
				}
			} else {
				orgKeyMap = new ConcurrentHashMap<String, String>();
				if (!orgKeyMap.containsKey(orgKey)) {
					orgKeyMap.put(orgKey, value);
					tagMap.put(tagName, orgKeyMap);
				}
			}
			counter++;
//			if (counter != 0)
//				break;
		}
//		System.out.println("Total Rows : " + counter);
//		System.out.println("Outer Map : " + tagMap);
//		System.out.println("length : " + tagMap.size());
		return tagMap;
	}

	private static Map<String, Map<String, Map<String, String>>> createLanguageMap(XSSFSheet localizationSheet) {

		Map<String, Map<String, Map<String, String>>> languageOuterMap = new HashMap<String, Map<String, Map<String, String>>>();
		Map<String, Map<String, String>> languageMiddleMap;
		Map<String, String> languageInnerMap;

		Iterator<Row> rowIterator1 = localizationSheet.iterator();
		rowIterator1.next();
		int counter11 = 0;
		while (rowIterator1.hasNext()) {
			Row row = rowIterator1.next();
			String outerKey = row.getCell(6).getStringCellValue().trim();
			String value = row.getCell(7).getStringCellValue().trim();
			Cell orgKeyCell = row.getCell(10);
			Cell languageCell = row.getCell(8);
			String middleKey = "";
			if (orgKeyCell != null) {
				middleKey = orgKeyCell.getStringCellValue().trim();
			} else
				middleKey = "GLOBAL";
			String innerKey = "";
			if (languageCell != null) {
				innerKey = languageCell.getStringCellValue().trim();
			} else
				innerKey = "GLOBAL";
			if (languageOuterMap.containsKey(outerKey)) {
				languageMiddleMap = languageOuterMap.get(outerKey);
				if (languageMiddleMap.containsKey(middleKey)) {
					languageInnerMap = languageMiddleMap.get(middleKey);
					if (!languageInnerMap.containsKey(innerKey)) {
						languageInnerMap.put(innerKey, value);
						languageMiddleMap.put(middleKey, languageInnerMap);
					}
				} else {
					languageInnerMap = new ConcurrentHashMap<String, String>();
					languageInnerMap.put(innerKey, value);
					languageMiddleMap.put(middleKey, languageInnerMap);
				}
			} else {
				languageInnerMap = new ConcurrentHashMap<String, String>();
				languageInnerMap.put(innerKey, value);
				languageMiddleMap = new ConcurrentHashMap<String, Map<String, String>>();
				languageMiddleMap.put(middleKey, languageInnerMap);
				languageOuterMap.put(outerKey, languageMiddleMap);

			}
			counter11++;
//			if (counter11 == 2)
//				break;
		}
//		System.out.println("Map  \n\n" + languageOuterMap);
//		if (languageOuterMap.containsKey("notification.general.template1.HTMLblock1")){
//			System.out.println("TestValue : " + languageOuterMap.get("notification.general.template1.HTMLblock1").get("XTM20151023153049").get("GLOBAL"));
//		}
		return languageOuterMap;
	}

	private static void replaceLocalizationTagWithValueAndN6MergeTermWithN7InTemplate(XSSFSheet templateSheet,
			Map<String, Map<String, String>> tagMap, Map<String, Map<String, Map<String, String>>> languageOuterMap,
			Map<String, String> mergeTermsMap) throws Exception {
		String discrepanciesFile = "C:\\Users\\Sakhi\\Desktop\\xTime\\DiscrepanciesInLocalizationAndMergeTerms.xlsx";
		FileInputStream fileInputStream = new FileInputStream(new File(discrepanciesFile));
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet localizationDiscrepanciesSheet = workbook.getSheetAt(0);
		XSSFSheet mergeTermDiscrepanciesSheet = workbook.getSheetAt(1);
//		XSSFSheet SubjectlocalizationDiscrepanciesSheet = workbook.getSheetAt(2);
//		XSSFSheet SubjectMergeTermDiscrepanciesSheet = workbook.getSheetAt(3);
		fileInputStream.close();
		ArrayList<String> dealersList = new ArrayList<>();
		Map<String, Integer> total_templates = new HashMap<String, Integer>();
		Map<String, Integer> templates_ready_to_work_on = new HashMap<String, Integer>();

		int counter1 = 0;
		Iterator<Row> templateSheetRowIterator = templateSheet.iterator();
		Row row = templateSheetRowIterator.next();
		while (templateSheetRowIterator.hasNext()) {
			row = templateSheetRowIterator.next();

			Cell templateCell = row.getCell(27);
			Cell orgCell = row.getCell(6);
			Cell templateSubjectCell = row.getCell(18);
			String templateString = "";
			if (templateCell != null) {
				templateString = templateCell.getStringCellValue();
			}
			String templateSubject = "";
			if (templateSubjectCell != null) {
				templateSubject = templateSubjectCell.getStringCellValue().trim();
			}

			String rawTemplate = templateString;
			String orgStringValue = orgCell.getStringCellValue().trim();

			if (!dealersList.contains(orgStringValue))
				dealersList.add(orgStringValue);

			if (!total_templates.containsKey(orgStringValue))
				total_templates.put(orgStringValue, 0);

			if (!templates_ready_to_work_on.containsKey(orgStringValue))
				templates_ready_to_work_on.put(orgStringValue, 0);

			Cell excludedFile = row.getCell(37);
			String excludedFileName = "";
			if (excludedFile != null)
				excludedFileName = excludedFile.getStringCellValue().trim();

			if (!excludedFileName.equals("Y - Do Not Migrate")) {

				if (!total_templates.containsKey(orgStringValue))
					total_templates.put(orgStringValue, 1);
				else if (total_templates.containsKey(orgStringValue))
					total_templates.put(orgStringValue, total_templates.get(orgStringValue) + 1);

				counter1++;
				int templateRowNumber = row.getRowNum() + 1;
				System.out.println("Row Number : " + templateRowNumber + "   " + row.getCell(0).getNumericCellValue());
//				
//				/* Update and replace Subject Of Template */
//				templateSubject = replaceLocalizationStringsWithValues(templateSubject, templateRowNumber,
//						SubjectlocalizationDiscrepanciesSheet, orgStringValue, tagMap, languageOuterMap);
//				templateSubject = updateAllN6MergeTermWithN7MergeTerms(templateRowNumber,
//						SubjectMergeTermDiscrepanciesSheet, templateSubject, mergeTermsMap, rawTemplate);
////				System.out.println("Row Number : " + templateRowNumber + "\t " + templateSubject);
//				replaceSubjectStringInTemplateSheet(templateRowNumber, templateSheet, templateSubject); 
//				/************** END *****************/

				/* Update Template */

				templateString = replaceDECODEsyntax(templateString, templateRowNumber, localizationDiscrepanciesSheet,
						orgStringValue, tagMap, languageOuterMap);
				templateString = replaceLocalizationStringsWithValues(templateString, templateRowNumber,
						localizationDiscrepanciesSheet, orgStringValue, tagMap, languageOuterMap);
				templateString = removeContentSaidByClient(templateString);
				templateString = replaceCalendarLinksWithMergeTermsProvidedByClient(templateString, templateRowNumber);
				templateString = updateAllN6MergeTermWithN7MergeTerms(templateRowNumber, mergeTermDiscrepanciesSheet,
						templateString, mergeTermsMap, rawTemplate);
				templateString = updateINITCAPsyntax(templateString);
				templateString = updateIFsyntax(templateString);

				String regex = "(\\$[A-Za-z_0-9]+)";
				Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
				Matcher matcher = pattern.matcher(templateString);

				matcher = pattern.matcher(templateString);
				boolean check1 = false;
				while (matcher.find()) {
					String mergeTerm = matcher.group().trim();
					if (!mergeTerm.equals("$if") && !mergeTerm.equals("$500") && !mergeTerm.equals("$59")
							&& !mergeTerm.equals("$139") && !mergeTerm.equals("$750") && !mergeTerm.equals("$19")
							&& !mergeTerm.equals("$200") && !mergeTerm.equals("$17") && !mergeTerm.equals("$5")
							&& !mergeTerm.equals("$1000") && !mergeTerm.equals("$1") && !mergeTerm.equals("$20")
							&& !mergeTerm.equals("$100") && !mergeTerm.equals("$250") && !mergeTerm.equals("$39")
							&& !mergeTerm.equals("$50") && !mergeTerm.equals("$35") && !mergeTerm.equals("$10")
							&& !mergeTerm.equals("$mailto")) {
						check1 = true;
						break;
					}

					// To create Files containing Dollar Amount like string
//					if (mergeTerm.equals("$500") || mergeTerm.equals("$59") || mergeTerm.equals("$139")
//							|| mergeTerm.equals("$750") || mergeTerm.equals("$19") || mergeTerm.equals("$200")
//							|| mergeTerm.equals("$17") || mergeTerm.equals("$5") || mergeTerm.equals("$1000")
//							|| mergeTerm.equals("$1") || mergeTerm.equals("$20") || mergeTerm.equals("$100")
//							|| mergeTerm.equals("$250") || mergeTerm.equals("$39") || mergeTerm.equals("$50")
//							|| mergeTerm.equals("$35") || mergeTerm.equals("$10")) {
//						check1 = true;
//						break;
//					}

				}

				ArrayList<String> group1DealersList = group1DealerList();
				String dealerGroup = "Group 2";
				if (group1DealersList.contains(orgStringValue))
					dealerGroup = "Group 1";

				if (check1 == false) {
					if (!templateString.contains("<<NOT DEFINED!!!!>>")) {
						if (!templates_ready_to_work_on.containsKey(orgStringValue))
							templates_ready_to_work_on.put(orgStringValue, 1);
						else if (templates_ready_to_work_on.containsKey(orgStringValue))
							templates_ready_to_work_on.put(orgStringValue,
									templates_ready_to_work_on.get(orgStringValue) + 1);
						okTemplates++;
//						createHTMLfile(templateString, templateRowNumber, dealerGroup, orgStringValue);
//						writeTemplateCharactersCountInColumnAH_AndGroupInAI(templateRowNumber, templateSheet,
//								templateString, dealerGroup);
					}
				} else {
					if (!templateString.contains("<<NOT DEFINED!!!!>>")) {
						discrepantTemplates++;
//						createHTMLfileWithdiscrepancies(templateString, templateRowNumber, dealerGroup, orgStringValue);
//						writeTemplateCharactersCountInColumnAH_AndGroupInAI(templateRowNumber, templateSheet,
//								templateString, dealerGroup);
					}
				}
//					System.out.println("Updated Template : \n" + templateString + "\n");

				if (counter1 == 15) {
					break;
				}

				/************** END *****************/
			} else {
				counter1++;
			}

		}
//		System.out.println("check");
		System.out.println("okTemplates : " + okTemplates + "\nDiscrepant Template : " + discrepantTemplates);
//		writeDealersTotalFilesAndFilesReadyToMigrateInExcelFile(dealersList, total_templates,
//				templates_ready_to_work_on);

//		System.out.println("Counter  : " + counter1);
		FileOutputStream outputStream = new FileOutputStream(discrepanciesFile);
		workbook.write(outputStream);
		workbook.close();
		outputStream.flush();
		outputStream.close();
	}

	private static String replaceDECODEsyntax(String templateString, int templateRowNumber,
			XSSFSheet localizationDiscrepanciesSheet, String orgStringValue, Map<String, Map<String, String>> tagMap,
			Map<String, Map<String, Map<String, String>>> languageOuterMap) {

		String decodeRegEx = "(\\$\\~\\[DECODE([\\S\\s])*?.\\)\\])";
		Pattern pattern = Pattern.compile(decodeRegEx, Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String decodeString = matcher.group().trim();
			System.out.println(decodeString);
			String tempDecodeString = decodeString.substring(10, decodeString.length() - 2);
			System.out.println(tempDecodeString);
			String[] decodeArguments = tempDecodeString.split(",");
			int totalArguments = decodeArguments.length;
			String firstArgument = "", secondArgument = "", thirdArgument = "", fourthArgument = "", fifthArgument = "",
					sixthArgument = "";

			firstArgument = decodeArguments[0].trim();
			firstArgument = firstArgument.substring(1, firstArgument.length() - 1);
			secondArgument = decodeArguments[1].trim();
			secondArgument = secondArgument.substring(1, secondArgument.length() - 1);
			thirdArgument = decodeArguments[2].trim();
			thirdArgument = thirdArgument.substring(1, thirdArgument.length() - 1);
			fourthArgument = decodeArguments[3].trim();
			fourthArgument = fourthArgument.substring(1, fourthArgument.length() - 1);
			String callingMethod = "DECODE";
			String valueAgainstLocalizationString = valueAgainstLocalizationString(templateString, firstArgument,
					orgStringValue, tagMap, languageOuterMap, templateRowNumber, localizationDiscrepanciesSheet,
					callingMethod);
			String decodeReplacementString = "";

			if (totalArguments == 4) {
				if (secondArgument.equals("")) {
//					System.out.println("SECONG ARGUMENT IS EMPTY");
					if (!valueAgainstLocalizationString.equals("GLOBAL") && !valueAgainstLocalizationString.equals(""))
						decodeReplacementString = fourthArgument;
					else if (valueAgainstLocalizationString.equals("GLOBAL"))
						decodeReplacementString = thirdArgument;
				} else {
					if (valueAgainstLocalizationString.equals(secondArgument))
						decodeReplacementString = thirdArgument;
					else
						decodeReplacementString = fourthArgument;
				}
			} else if (totalArguments == 6) {
				fifthArgument = decodeArguments[4].trim();
				fifthArgument = fifthArgument.substring(1, fifthArgument.length() - 1);
				sixthArgument = decodeArguments[5].trim();
				sixthArgument = sixthArgument.substring(1, sixthArgument.length() - 1);
				if (valueAgainstLocalizationString.equals(secondArgument))
					decodeReplacementString = thirdArgument;
				else if (valueAgainstLocalizationString.equals(fourthArgument))
					decodeReplacementString = fifthArgument;
				else
					decodeReplacementString = sixthArgument;
			}
			templateString.replace(decodeString, decodeReplacementString);

			System.out.println(firstArgument + " " + firstArgument.length() + "\n" + secondArgument + " "
					+ secondArgument.length() + "\n" + thirdArgument + " " + thirdArgument.length() + "\n"
					+ fourthArgument + " " + fourthArgument.length());

		}

		return templateString;

	}

	private static void writeDealersTotalFilesAndFilesReadyToMigrateInExcelFile(ArrayList<String> dealersList,
			Map<String, Integer> total_templates, Map<String, Integer> templates_ready_to_work_on) throws Exception {

		String dealerListFile = "C:\\Users\\Sakhi\\Desktop\\xTime\\DealerListWithTotalTemplatesAndTemplateReadyToMigrate.xlsx";
		FileInputStream fileInputStream = new FileInputStream(new File(dealerListFile));
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet dealersReadyTemplatesSheet = workbook.getSheetAt(0);
		fileInputStream.close();
		for (int i = 0; i < dealersList.size(); i++) {
			String dealer = dealersList.get(i);
//			System.out.println(
//					dealer + "\t" + total_templates.get(dealer) + "\t" + templates_ready_to_work_on.get(dealer));
			Row newRow = dealersReadyTemplatesSheet.createRow(i + 1);
			Cell dealerName = newRow.createCell(0);
			Cell totalTemplates = newRow.createCell(1);
			Cell templateReadyToMigrate = newRow.createCell(2);

			dealerName.setCellValue(dealer);
			totalTemplates.setCellValue(total_templates.get(dealer));
			templateReadyToMigrate.setCellValue(templates_ready_to_work_on.get(dealer));

		}
		FileOutputStream outputStream = new FileOutputStream(dealerListFile);
		workbook.write(outputStream);
		workbook.close();
		outputStream.flush();
		outputStream.close();
		System.out.println("ALHAMDULILLAH \nDONE");

	}

	private static String removeContentSaidByClient(String templateString) {

//		if (templateString.contains("$isvalet"))
//			templateString = templateString.replaceAll("\\$if \\$isvalet=[A-Za-z_0-9\\ \\=\\'\\']+",
//					"<#if transportationType == 'Valet'>");
//		if (templateString.contains("$valetloaner"))
//			templateString = templateString.replaceAll("\\$if \\$valetloaner=[A-Za-z_0-9\\ \\=\\'\\']+",
//					"<#if transportationType == 'Valet with Loaner'>");
//		

		if (templateString.contains("$if $isvalet='Y'")) {
			templateString = templateString.replace("$if $isvalet='Y'", "<#if transportationType == 'Valet'>");
		}
		if (templateString.contains("$if $valetloaner='Y'")) {
			templateString = templateString.replace("$if $valetloaner='Y'",
					"<#if transportationType == 'Valet with Loaner'>");
		}
		if (templateString.contains("The following appointment has been $apptstate:"))
			templateString = templateString.replace("The following appointment has been $apptstate:", "");
		if (templateString.contains("THE FOLLOWING APPOINTMENT HAS BEEN<br><br>$apptstate:<br><br>"))
			templateString = templateString.replace("THE FOLLOWING APPOINTMENT HAS BEEN<br><br>$apptstate:<br><br>",
					"");
		if (templateString.contains("<title>Hyundai Appointment $apptstate Email</title>"))
			templateString = templateString.replace("<title>Hyundai Appointment $apptstate Email</title>", "");

		if (templateString.contains("Last modified on: $lastmodifiedtime"))
			templateString = templateString.replace("Last modified on: $lastmodifiedtime", "");
		if (templateString.contains("| Last Modified: $lastmodifiedtime"))
			templateString = templateString.replace("| Last Modified: $lastmodifiedtime", "");

		if (templateString.contains("Team: $teamname<br />"))
			templateString = templateString.replace("Team: $teamname<br />", "");
		if (templateString.contains("Team: $teamname<br>"))
			templateString = templateString.replace("Team: $teamname<br>", "");
		if (templateString.contains("<span>$teamname</span>"))
			templateString = templateString.replace("<span>$teamname</span>", "");
		if (templateString.contains("<tr><td>Team:</td><td><b>$teamname</b></td></tr>"))
			templateString = templateString.replace("<tr><td>Team:</td><td><b>$teamname</b></td></tr>", "");

		if (templateString.contains("Trim Information: $trim, $enginetype, $enginesize, $drivetype, $transmissiontype"))
			templateString = templateString
					.replace("Trim Information: $trim, $enginetype, $enginesize, $drivetype, $transmissiontype", "");
		if (templateString.contains("$apptdayth of $apptparitalmonth"))
			templateString = templateString.replace("$apptdayth of $apptparitalmonth", "${apptDateTime}");

//		if (templateString.contains("<b><i>Cancellation Notes:</b></i><br />\r\n" + 
//				"\r\n" + 
//				"$cancellationreasontype  \r\n" + 
//				"\r\n" + 
//				"$cancellationreasonnote,\r\n" + 
//				"\r\n" + 
//				"<br /><br />"))
//			templateString = templateString.replace("<b><i>Cancellation Notes:</b></i><br />\r\n" + 
//					"\r\n" + 
//					"$cancellationreasontype  \r\n" + 
//					"\r\n" + 
//					"$cancellationreasonnote,\r\n" + 
//					"\r\n" + 
//					"<br /><br />", "");

		if (templateString.contains("<b>Cancellation Notes:</b> $cancellationreasonnote"))
			templateString = templateString.replace("<b>Cancellation Notes:</b> $cancellationreasonnote", "");
		if (templateString.contains("<b>Cancellation Reason:</b><br />$cancellationreasonnote<br /><br />"))
			templateString = templateString
					.replace("<b>Cancellation Reason:</b><br />$cancellationreasonnote<br /><br />", "");
		if (templateString.contains("<i><b>Cancellation Notes:</b> $cancellationreasonnote</i>"))
			templateString = templateString.replace("<i><b>Cancellation Notes:</b> $cancellationreasonnote</i>", "");
		if (templateString.contains("<b><i>Cancellation Notes:</i></b> $cancellationreasonnote"))
			templateString = templateString.replace("<b><i>Cancellation Notes:</i></b> $cancellationreasonnote", "");
		if (templateString.contains("Cancellation Notes:&nbsp;&nbsp;<b> $cancellationreasonnote</b>"))
			templateString = templateString.replace("Cancellation Notes:&nbsp;&nbsp;<b> $cancellationreasonnote</b>",
					"");
		if (templateString.contains("$cancellationreasonnote,"))
			templateString = templateString.replace("$cancellationreasonnote,", "");

		if (templateString.contains("<b>Cancellation Type:</b> $cancellationreasontype"))
			templateString = templateString.replace("<b>Cancellation Type:</b> $cancellationreasontype", "");
		if (templateString.contains("<i><b>Reason:</b> $cancellationreasontype<br /></i>"))
			templateString = templateString.replace("<i><b>Reason:</b> $cancellationreasontype<br /></i>", "");
		if (templateString.contains("<b>Cancellation Reason:</b> $cancellationreasontype<br />"))
			templateString = templateString.replace("<b>Cancellation Reason:</b> $cancellationreasontype<br />", "");
		if (templateString.contains("<b>Cancellation Reason Type:</b><br />$cancellationreasontype<br />"))
			templateString = templateString
					.replace("<b>Cancellation Reason Type:</b><br />$cancellationreasontype<br />", "");
		if (templateString.contains("Cancellation Type:&nbsp;&nbsp;<b>$cancellationreasontype</b><br />"))
			templateString = templateString
					.replace("Cancellation Type:&nbsp;&nbsp;<b>$cancellationreasontype</b><br />", "");
		if (templateString.contains("$cancellationreasontype"))
			templateString = templateString.replace("$cancellationreasontype", "");

		return templateString;
	}

	private static String valueAgainstLocalizationString(String templateString, String tagName, String orgStringValue,
			Map<String, Map<String, String>> tagMap, Map<String, Map<String, Map<String, String>>> languageOuterMap,
			int templateRowNumber, XSSFSheet localizationDiscrepanciesSheet, String callingMethod) {

		String tagNameAsLocalizedSheet = tagName.substring(2, tagName.length() - 1); // To remove $[ and ]
		if (tagNameAsLocalizedSheet.equals("terms.valet.ALLCAPS")
				|| tagNameAsLocalizedSheet.equals("terms.VALET.allcaps"))
			tagNameAsLocalizedSheet = "terms.VALET.allcaps";
		if (tagNameAsLocalizedSheet.equals("notif.tci.toyota1.TEXT06nocal")
				|| tagNameAsLocalizedSheet.equals("notif.tci.toyota1.TEXT06NoCal"))
			tagNameAsLocalizedSheet = "notif.tci.toyota1.TEXT06NoCal";
		if (tagNameAsLocalizedSheet.equals("terms.valet.ls") || tagNameAsLocalizedSheet.equals("terms.valet.LS"))
			tagNameAsLocalizedSheet = "terms.valet.LS";
		if (tagNameAsLocalizedSheet.equals("terms.loaner.ls") || tagNameAsLocalizedSheet.equals("terms.loaner.LS"))
			tagNameAsLocalizedSheet = "terms.loaner.LS";

		String localizedTermValueFromLocalizedSheet = null; // from tag
		if (!tagMap.containsKey(tagNameAsLocalizedSheet)) {
			if (!callingMethod.equals("DECODE"))
				writeLocalizedStringDiscripentDataInFile(templateRowNumber, tagNameAsLocalizedSheet, orgStringValue,
						localizationDiscrepanciesSheet, localizationStringsRowCounter++, 1);
		} else {
			if (tagMap.get(tagNameAsLocalizedSheet).containsKey(orgStringValue)) {
				localizedTermValueFromLocalizedSheet = tagMap.get(tagNameAsLocalizedSheet).get(orgStringValue);
			} else if (tagMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
				if (languageOuterMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
					if (languageOuterMap.get(tagNameAsLocalizedSheet).get("GLOBAL").containsKey("GLOBAL")) {
						localizedTermValueFromLocalizedSheet = languageOuterMap.get(tagNameAsLocalizedSheet)
								.get("GLOBAL").get("GLOBAL");
						if (callingMethod.equals("DECODE"))
							localizedTermValueFromLocalizedSheet = "GLOBAL";
					}
				}
			}
		}

		return localizedTermValueFromLocalizedSheet;
	}

	private static String replaceLocalizationStringsWithValues(String templateString, int templateRowNumber,
			XSSFSheet localizationDiscrepanciesSheet, String orgStringValue, Map<String, Map<String, String>> tagMap,
			Map<String, Map<String, Map<String, String>>> languageOuterMap) {

		boolean isLocalizedStringExist = true;
		while (isLocalizedStringExist) {
			String regex = "(\\$\\[.*?\\])";
			Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
			Matcher matcher = pattern.matcher(templateString);
			while (matcher.find()) {
				String tagName = matcher.group().trim();
//				String tagNameAsLocalizedSheet = tagName.substring(2, tagName.length() - 1); // To remove $[ and ]
//				if (tagNameAsLocalizedSheet.equals("terms.valet.ALLCAPS")
//						|| tagNameAsLocalizedSheet.equals("terms.VALET.allcaps"))
//					tagNameAsLocalizedSheet = "terms.VALET.allcaps";
//				if (tagNameAsLocalizedSheet.equals("notif.tci.toyota1.TEXT06nocal")
//						|| tagNameAsLocalizedSheet.equals("notif.tci.toyota1.TEXT06NoCal"))
//					tagNameAsLocalizedSheet = "notif.tci.toyota1.TEXT06NoCal";
//				if (tagNameAsLocalizedSheet.equals("terms.valet.ls")
//						|| tagNameAsLocalizedSheet.equals("terms.valet.LS"))
//					tagNameAsLocalizedSheet = "terms.valet.LS";
//				if (tagNameAsLocalizedSheet.equals("terms.loaner.ls")
//						|| tagNameAsLocalizedSheet.equals("terms.loaner.LS"))
//					tagNameAsLocalizedSheet = "terms.loaner.LS";
//
//				String localizedTermValueFromLocalizedSheet = null; // from tag
//				if (!tagMap.containsKey(tagNameAsLocalizedSheet)) {
//					writeLocalizedStringDiscripentDataInFile(templateRowNumber, tagNameAsLocalizedSheet, orgStringValue,
//							localizationDiscrepanciesSheet, localizationStringsRowCounter++, 1);
//				} else {
//					if (tagMap.get(tagNameAsLocalizedSheet).containsKey(orgStringValue)) {
//						localizedTermValueFromLocalizedSheet = tagMap.get(tagNameAsLocalizedSheet).get(orgStringValue);
//					} else if (tagMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
//						if (languageOuterMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
//							if (languageOuterMap.get(tagNameAsLocalizedSheet).get("GLOBAL").containsKey("GLOBAL")) {
//								localizedTermValueFromLocalizedSheet = languageOuterMap.get(tagNameAsLocalizedSheet)
//										.get("GLOBAL").get("GLOBAL");
//							}
//						}
//					}
//				}
				String callingMethod = "LOCALIZATION";
				String localizedTermValueFromLocalizedSheet = valueAgainstLocalizationString(templateString, tagName,
						orgStringValue, tagMap, languageOuterMap, templateRowNumber, localizationDiscrepanciesSheet,
						callingMethod); // from tag

				if (localizedTermValueFromLocalizedSheet != null) {
					templateString = templateString.replace(tagName, localizedTermValueFromLocalizedSheet);
				} else {
					templateString = templateString.replace(tagName, "<<NOT DEFINED!!!!>>");
				}
			}
			matcher = pattern.matcher(templateString);
			if (!matcher.find()) {
				isLocalizedStringExist = false;
			}
		}
		return templateString;
	}

	private static String replaceCalendarLinksWithMergeTermsProvidedByClient(String templateString,
			int templateRowNumber) {

		String googleLinkRegex = "(http:\\/\\/www\\.google\\.com\\/calendar[A-Za-z0-9\\?\\{\\}\\/\\%\\&\\=\\!\\_\\.\\ \\r\\n\\$]+)";
		String yahooLinkRegex = "(http:\\/\\/calendar\\.yahoo\\.com[A-Za-z0-9\\?\\{\\}\\/\\%\\&\\=\\!\\_\\.\\ \\r\\n\\$]+)";
		String microsoftLiveLinkRegex = "(http:\\/\\/calendar\\.live\\.com\\/calendar\\/calendar\\.aspx[A-Za-z0-9\\?\\{\\}\\/\\%\\&\\=\\!\\_\\.\\ \\r\\n\\$]+)";

		Pattern pattern = Pattern.compile(googleLinkRegex);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String linkText = matcher.group().trim();
			templateString = templateString.replace(linkText, "${googleCalendarLink}");
//			System.out.println("Google Link : " + linkText);
		}

		pattern = Pattern.compile(yahooLinkRegex);
		matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String linkText = matcher.group().trim();
			templateString = templateString.replace(linkText, "${yahooCalendarLink}");
//			System.out.println("Yahoo Link : " + linkText);
		}

		pattern = Pattern.compile(microsoftLiveLinkRegex);
		matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String linkText = matcher.group().trim();
			templateString = templateString.replace(linkText, "");
//			System.out.println("Live Link : " + linkText);
		}

		return templateString;
	}

	private static String updateIFsyntax(String templateString) {
		String regex = "\\$if[A-Za-z_0-9\\{\\}\\ \\=\\'\\'\\$]+";
		Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String oldIfString = matcher.group().trim();
			String tempOldIfString = oldIfString;
			if (tempOldIfString.contains("{")) {
				String regex1 = "(\\$\\{[A-Za-z_0-9]+\\})";
				Pattern pattern1 = Pattern.compile(regex1, Pattern.CASE_INSENSITIVE);
				Matcher matcher1 = pattern1.matcher(tempOldIfString);
				String ifNewString = "";
				while (matcher1.find()) {
					String mergeTerm = matcher1.group().trim();
					String splittedMergeTerm = mergeTerm.substring(2, mergeTerm.length() - 1);
					if (tempOldIfString.contains("="))
						tempOldIfString = tempOldIfString.replace("=", "==");
					else
						splittedMergeTerm = splittedMergeTerm + "?has_content";
					tempOldIfString = tempOldIfString.replace(mergeTerm, splittedMergeTerm);
				}
				ifNewString = "<#" + tempOldIfString.substring(1) + ">";
				templateString = templateString.replace(oldIfString, ifNewString);
			}
		}
		return templateString;
	}

	private static String updateINITCAPsyntax(String templateString) {
		String regex = "(\\$\\~\\[INITCAP([\\S\\s])*?.\\])";
		Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String initCapString = matcher.group().trim();
			String regex1 = "(\\$\\{[A-Za-z_0-9]+\\})";
			Pattern pattern1 = Pattern.compile(regex1, Pattern.CASE_INSENSITIVE);
			Matcher matcher1 = pattern1.matcher(initCapString);
			String capitalizedString = "";
			while (matcher1.find()) {
				String mergeTerm = matcher1.group().trim();
				capitalizedString += mergeTerm.substring(0, mergeTerm.length() - 1) + "?capitalize} ";
			}
			capitalizedString = capitalizedString.substring(0, capitalizedString.length() - 1); // to remove last space
			templateString = templateString.replace(initCapString, capitalizedString);
		}
		return templateString;
	}

	private static void createHTMLfile(String templateString, int templateRowNumber, String dealerGroup,
			String orgStringValue) throws Exception {

		String newDirectoryPath = "C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue;
		createNewFolder(newDirectoryPath);
		File old_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue + "\\AB" + templateRowNumber + ".html");
		old_file.delete();
		File new_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue + "\\AB" + templateRowNumber + ".html");
		try (FileWriter fw = new FileWriter(new_file, true);
				BufferedWriter bw = new BufferedWriter(fw);
				PrintWriter out = new PrintWriter(bw)) {
			out.println(templateString);
		}
	}

	private static void createHTMLfileWithdiscrepancies(String templateString, int templateRowNumber,
			String dealerGroup, String orgStringValue) throws Exception {

		String newDirectoryPath = "C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue;
		createNewFolder(newDirectoryPath);

		String discrepantFileDirectory = newDirectoryPath + "\\Discrepant Files";
		createNewFolder(discrepantFileDirectory);

		File old_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue + "\\Discrepant Files\\AB" + templateRowNumber + ".html");
		old_file.delete();
		File new_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\" + dealerGroup + "\\"
				+ orgStringValue + "\\Discrepant Files\\AB" + templateRowNumber + ".html");

		try (FileWriter fw = new FileWriter(new_file, true);
				BufferedWriter bw = new BufferedWriter(fw);
				PrintWriter out = new PrintWriter(bw)) {
			out.println(templateString);
		}

	}

	private static void createNewFolder(String newDirectoryPath) {
		File theDir = new File(newDirectoryPath);
		// if the directory does not exist, create it
		if (!theDir.exists()) {
			System.out.println("creating directory: " + theDir.getName());
			boolean result = false;

			try {
				theDir.mkdir();
				result = true;
			} catch (SecurityException se) {
				// handle it
			}
			if (result) {
				System.out.println("DIR created");
			}
		}

	}

	private static String updateAllN6MergeTermWithN7MergeTerms(int templateRowNumber,
			XSSFSheet mergeTermDiscrepanciesSheet, String templateString, Map<String, String> mergeTermsMap,
			String rawTemplate) throws Exception {

		String regex = "(\\$[A-Za-z_0-9]+)";
		Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String mergeTerm = matcher.group().trim(); //
//				System.out.println(mergeTerm);
			if (!mergeTerm.equals("$if") && !mergeTerm.equals("$endif")
					&& !mergeTerm.equals("$mailto")/*
													 * && !mergeTerm.equals("$apptstartdatetimevcal") &&
													 * !mergeTerm.equals("$apptenddatetimevcal") &&
													 * !mergeTerm.equals("$googlecalendardealeraddress")
													 */) {
				if (!mergeTermsMap.containsKey(mergeTerm)) {
					if (mergeTerm.equals("$500") || mergeTerm.equals("$59") || mergeTerm.equals("$139")
							|| mergeTerm.equals("$750") || mergeTerm.equals("$19") || mergeTerm.equals("$200")
							|| mergeTerm.equals("$17") || mergeTerm.equals("$5") || mergeTerm.equals("$1000")
							|| mergeTerm.equals("$1") || mergeTerm.equals("$20") || mergeTerm.equals("$100")
							|| mergeTerm.equals("$250") || mergeTerm.equals("$39") || mergeTerm.equals("$50")
							|| mergeTerm.equals("$35") || mergeTerm.equals("$10")) {
//						writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet, mergeTerm,
//								4);
					} else {
						writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet, mergeTerm, 1,
								rawTemplate);
					}
				} else {
					if (!mergeTermsMap.get(mergeTerm).equals("NA") && !mergeTermsMap.get(mergeTerm).contains("??")) {
						templateString = templateString.replace(mergeTerm, mergeTermsMap.get(mergeTerm));
					} else {
						if (mergeTermsMap.get(mergeTerm).equals("NA")) {
							writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet,
									mergeTerm, 2, rawTemplate);
						} else if (mergeTermsMap.get(mergeTerm).contains("??")) {
							writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet,
									mergeTerm, 3, rawTemplate);
						}
					}
				}
			} else {
				if (mergeTerm.equals("$endif"))
					templateString = templateString.replace(mergeTerm, "</#if>");
			}
		}

		return templateString;
	}

	private static void writeMergeTermDiscripentDataInFile(int templateRowNumber, XSSFSheet mergeTermDiscrepanciesSheet,
			String mergeTerm, int reason, String rawTemplate) {
		boolean check = false;

		if (mergeTermsRowCounter == 1) {
			createNewRowInMergeTermSheetPopulateCells(mergeTermDiscrepanciesSheet, templateRowNumber, mergeTerm, reason,
					rawTemplate);
		} else {
			Iterator<Row> rowIterator = mergeTermDiscrepanciesSheet.iterator();
			Row row = rowIterator.next();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				String rowCell = row.getCell(0).getStringCellValue();
				String mergeTermCell = row.getCell(1).getStringCellValue();
				if (/* rowCell.equals("AB" + templateRowNumber)&& */ mergeTermCell.equals(mergeTerm)) {
					check = true;
					break;
				}
			}
		}
		if (check == false) {
			createNewRowInMergeTermSheetPopulateCells(mergeTermDiscrepanciesSheet, templateRowNumber, mergeTerm, reason,
					rawTemplate);
		}
	}

	private static void createNewRowInMergeTermSheetPopulateCells(XSSFSheet mergeTermDiscrepanciesSheet,
			int templateRowNumber, String mergeTerm, int reason, String rawTemplate) {
//		System.out.println(globalRowCounter);
		Row newRow = mergeTermDiscrepanciesSheet.createRow(mergeTermsRowCounter++);
		Cell templateRowNumCell = newRow.createCell(0);
		Cell mergeTermCell = newRow.createCell(1);
		Cell reasonCell = newRow.createCell(2);
		Cell rawTemplateCell = newRow.createCell(3);
		templateRowNumCell.setCellValue("AB" + templateRowNumber);
		mergeTermCell.setCellValue(mergeTerm);
		rawTemplateCell.setCellValue(rawTemplate);
		if (reason == 1) {
			reasonCell.setCellValue("Merge term not found in glosaary file");
		} else if (reason == 2) {
			reasonCell.setCellValue("NA value in N7 Type");
		} else if (reason == 3)
			reasonCell.setCellValue("N7 MergeTerm Contains ?? in it");
		else if (reason == 4)
			reasonCell.setCellValue("dollar ammount in merge term");
	}

	private static void writeTemplateCharactersCountInColumnAH_AndGroupInAI(int templateRowNumber,
			XSSFSheet templateSheet, String templateString, String dealerGroup) {
		System.out.println("Print Row Number : " + templateRowNumber);
		Row row = templateSheet.getRow(templateRowNumber - 1);
//		Cell templateCell = row.createCell(27);
//		templateCell = row.getCell(27);
//		templateCell.setCellValue(templateString);
		Cell templateCharacterCountCell = row.createCell(33);
		templateCharacterCountCell = row.getCell(33);
		templateCharacterCountCell.setCellValue(templateString.length());
		Cell templateGroupCell = row.createCell(34);
		templateGroupCell = row.getCell(34);
		templateGroupCell.setCellValue(dealerGroup);

	}

	private static void replaceSubjectStringInTemplateSheet(int templateRowNumber, XSSFSheet templateSheet,
			String templateSubject) {
		System.out.println("Print Row Number : " + templateRowNumber);
		Row row = templateSheet.getRow(templateRowNumber - 1);
		Cell templateCell = row.createCell(18);
		templateCell = row.getCell(18);
		templateCell.setCellValue(templateSubject);
	}

	private static void writeLocalizedStringDiscripentDataInFile(int templateRowNumber,
			String localizedTermValueFromLocalizedSheet, String orgStringValue,
			XSSFSheet localizationDiscrepanciesSheet, int rowCount, int reason) {

		boolean check = false;

		if (rowCount == 1) {
			createNewRowAndPopulateCells(localizationDiscrepanciesSheet, rowCount, templateRowNumber, orgStringValue,
					localizedTermValueFromLocalizedSheet, reason);
		} else {
			Iterator<Row> rowIterator = localizationDiscrepanciesSheet.iterator();
			Row row = rowIterator.next();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				String rowCell = row.getCell(0).getStringCellValue();
				String tagCell = row.getCell(1).getStringCellValue();
				String orgKeyCell = row.getCell(2).getStringCellValue();
				if (rowCell.equals("AB" + templateRowNumber) && tagCell.equals(localizedTermValueFromLocalizedSheet)
				/* && orgKeyCell.equals(orgStringValue) */) {
					check = true;
					break;
				}
			}
		}
		if (check == false) {
			createNewRowAndPopulateCells(localizationDiscrepanciesSheet, rowCount, templateRowNumber, orgStringValue,
					localizedTermValueFromLocalizedSheet, reason);
		}

	}

	private static void createNewRowAndPopulateCells(XSSFSheet localizationDiscrepanciesSheet, int rowCount,
			int templateRowNumber, String orgStringValue, String localizedTermValueFromLocalizedSheet, int reason) {
		Row newRow = localizationDiscrepanciesSheet.createRow(rowCount);
		Cell templateRowNumCell = newRow.createCell(0);
		Cell templateTagCell = newRow.createCell(1);
		Cell templateOrgKeyCell = newRow.createCell(2);
		Cell reasonCell = newRow.createCell(3);
		templateRowNumCell.setCellValue("AB" + templateRowNumber);
		templateTagCell.setCellValue(localizedTermValueFromLocalizedSheet);
		templateOrgKeyCell.setCellValue(orgStringValue);
		if (reason == 1) {
			reasonCell.setCellValue("Tag name doesn't exist for respective localization string");
		} else if (reason == 2) {
			reasonCell.setCellValue("No entry found against org_key, nor global entry found with language_id empty");
		}

	}

}
