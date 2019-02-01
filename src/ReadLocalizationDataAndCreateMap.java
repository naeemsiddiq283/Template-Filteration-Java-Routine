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
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadLocalizationDataAndCreateMap {
	protected static int globalRowCounter = 1;

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

		String discrepanciesFile = "C:\\Users\\Sakhi\\Desktop\\xTime\\DiscrepanciesInLocalizationAndMergeTerms.xlsx";
		FileInputStream fileInputStream = new FileInputStream(new File(discrepanciesFile));
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet mergeTermDiscrepanciesSheet = workbook.getSheetAt(1);
		fileInputStream.close();

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
//
//			if (N7MergeTerm.equals("NA"))
//				createNewRowInMergeTermSheetPopulateCells(mergeTermDiscrepanciesSheet, 0, N6MergeTerm, 2);
//			if (N7MergeTerm.contains("??"))
//				createNewRowInMergeTermSheetPopulateCells(mergeTermDiscrepanciesSheet, 0, N6MergeTerm, 3);

		}
//		System.out.println("MergTerms Map Size: " + mergeTermMap.size() + "\nMergTerms Kap : " + mergeTermMap.get("$enginetype"));
		FileOutputStream outputStream = new FileOutputStream(discrepanciesFile);
		workbook.write(outputStream);
		workbook.close();
		outputStream.flush();
		outputStream.close();
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
		fileInputStream.close();

		int counter1 = 0;
		int rowCount = 1;
		Iterator<Row> templateSheetRowIterator = templateSheet.iterator();
		Row row = templateSheetRowIterator.next();
		while (templateSheetRowIterator.hasNext()) {
			row = templateSheetRowIterator.next();
			Cell templateCell = row.getCell(27);
			Cell orgCell = row.getCell(6);
			String templateString = "";
			if (templateCell != null) {
				templateString = templateCell.getStringCellValue();
			}
			String rawTemplate = templateString;

			String orgStringValue = orgCell.getStringCellValue().trim();
			counter1++;
			boolean check = false;
			boolean isLocalizedStringExist = true;
//			System.out.println(
//					"#########################################################################################################################################\n\n");
			int templateRowNumber = row.getRowNum() + 1;
			System.out.println("Row Number : " + templateRowNumber + "   " + row.getCell(0).getNumericCellValue());
			while (isLocalizedStringExist) {
				String regex = "(\\$\\[.*?\\])";
				Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
				Matcher matcher = pattern.matcher(templateString);
				while (matcher.find()) {
					String tagName = matcher.group().trim();
					String tagNameAsLocalizedSheet = tagName.substring(2, tagName.length() - 1); // To remove $[ and ]
					String localizedTermValueFromLocalizedSheet = null; // from tag
//					System.out.println(tagName);
					if (!tagMap.containsKey(tagNameAsLocalizedSheet)) {
						writeLocalizedStringDiscripentDataInFile(templateRowNumber, tagNameAsLocalizedSheet,
								orgStringValue, localizationDiscrepanciesSheet, rowCount++, 1);
					} else {
						if (tagMap.get(tagNameAsLocalizedSheet).containsKey(orgStringValue)) {
							localizedTermValueFromLocalizedSheet = tagMap.get(tagNameAsLocalizedSheet)
									.get(orgStringValue);
//							System.out.println("Contains Data : " + localizedTermValueFromLocalizedSheet);
						} else if (tagMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
							if (languageOuterMap.get(tagNameAsLocalizedSheet).containsKey("GLOBAL")) {
								if (languageOuterMap.get(tagNameAsLocalizedSheet).get("GLOBAL").containsKey("GLOBAL")) {
//									System.out.println(
//											"SDJLFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN");
									localizedTermValueFromLocalizedSheet = languageOuterMap.get(tagNameAsLocalizedSheet)
											.get("GLOBAL").get("GLOBAL");
								}
							}
						}
					}
					if (localizedTermValueFromLocalizedSheet != null) {
						templateString = templateString.replace(tagName, localizedTermValueFromLocalizedSheet);
					} else {
						templateString = templateString.replace(tagName, "NOT DEFINED!!!!");
					}
				}
				matcher = pattern.matcher(templateString);
				if (!matcher.find()) {
					isLocalizedStringExist = false;
				}
			}

			templateString = updateAllN6MergeTermWithN7MergeTerms(templateRowNumber, mergeTermDiscrepanciesSheet,
					templateString, mergeTermsMap, rawTemplate);
			templateString = replaceGoogleCalendarLinkWithaMergeTermProvidedByClient(templateString, templateRowNumber);

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
						&& !mergeTerm.equals("$50") && !mergeTerm.equals("$35") && !mergeTerm.equals("$10")) {
					check1 = true;
					break;
				}

				// To create Files containing Dollar Amount like string
//				if (mergeTerm.equals("$500") || mergeTerm.equals("$59") || mergeTerm.equals("$139")
//						|| mergeTerm.equals("$750") || mergeTerm.equals("$19") || mergeTerm.equals("$200")
//						|| mergeTerm.equals("$17") || mergeTerm.equals("$5") || mergeTerm.equals("$1000")
//						|| mergeTerm.equals("$1") || mergeTerm.equals("$20") || mergeTerm.equals("$100")
//						|| mergeTerm.equals("$250") || mergeTerm.equals("$39") || mergeTerm.equals("$50")
//						|| mergeTerm.equals("$35") || mergeTerm.equals("$10")) {
//					check1 = true;
//					break;
//				}

			}
			if (check1 == false) {
				createHTMLfile(templateString, templateRowNumber);
			}
//				System.out.println("Updated Template : \n" + templateString + "\n");
//			replaceTemplateStringInTemplateSheet(templateRowNumber, templateSheet, templateString);
//				System.out.println(
//						"#########################################################################################################################################\n\n");

//			if (counter1 == 1) {
//				System.out.println("Updated Template : \n" + templateString + "\n");
//				break;
//			}
		}
//		System.out.println("Counter  : " + counter1);
		FileOutputStream outputStream = new FileOutputStream(discrepanciesFile);
		workbook.write(outputStream);
		workbook.close();
		outputStream.flush();
		outputStream.close();
	}

	private static String replaceGoogleCalendarLinkWithaMergeTermProvidedByClient(String templateString,
			int templateRowNumber) {

		String googleLink = "http://www.google.com/calendar/event?action=TEMPLATE&text=Service%20appointment%20at%20${dealerName}&dates=$apptstartdatetimevcal/$apptenddatetimevcal&details=Service%20for%20your%20${vehicleMake}%20${vehicleModel}%20at%20$googlecalendardealeraddress%0D%0A%0D%0AThank%20you%20for%20using%20online%20service%20scheduling!&location=$googlecalendardealeraddress&trp=true";
		if (templateString.contains(googleLink)) {
//			System.out.println("YESSSSS!!!!!!!!!!!!!!!!!!!!     " + templateRowNumber);
			templateString = templateString.replace(googleLink, "${googleCalendarLink}");
		}
		return templateString;
	}

	private static String updateIFsyntax(String templateString) {
		String regex = "\\$if[A-Za-z_0-9\\{\\}\\ \\=\\'\\'\\$]+";
		Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(templateString);
		while (matcher.find()) {
			String oldIfString = matcher.group().trim();
//			System.out.println("Old : "+oldIfString);
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
//				System.out.println("New : " + ifNewString);
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
//			System.out.println("Old : " + initCapString);
			String regex1 = "(\\$\\{[A-Za-z_0-9]+\\})";
			Pattern pattern1 = Pattern.compile(regex1, Pattern.CASE_INSENSITIVE);
			Matcher matcher1 = pattern1.matcher(initCapString);
			String capitalizedString = "";
			while (matcher1.find()) {
				String mergeTerm = matcher1.group().trim();
				capitalizedString += mergeTerm.substring(0, mergeTerm.length() - 1) + "?capitalize} ";
			}
			capitalizedString = capitalizedString.substring(0, capitalizedString.length() - 1); // to remove last space
//			System.out.println("New : " + capitalizedString);
			templateString = templateString.replace(initCapString, capitalizedString);
		}
		return templateString;
	}

	private static void createHTMLfile(String templateString, int templateRowNumber) throws Exception {

		File old_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\AB" + templateRowNumber + ".html");
		old_file.delete();
		File new_file = new File("C:\\Users\\Sakhi\\Desktop\\xTime\\Template Files\\AB" + templateRowNumber + ".html");

		try (FileWriter fw = new FileWriter(new_file, true);
				BufferedWriter bw = new BufferedWriter(fw);
				PrintWriter out = new PrintWriter(bw)) {
			out.println(templateString);
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
			if (!mergeTerm.equals("$if") && !mergeTerm.equals("$endif")) {
				if (!mergeTermsMap.containsKey(mergeTerm)) {
					if (mergeTerm.equals("$500") || mergeTerm.equals("$59") || mergeTerm.equals("$139")
							|| mergeTerm.equals("$750") || mergeTerm.equals("$19") || mergeTerm.equals("$200")
							|| mergeTerm.equals("$17") || mergeTerm.equals("$5") || mergeTerm.equals("$1000")
							|| mergeTerm.equals("$1") || mergeTerm.equals("$20") || mergeTerm.equals("$100")
							|| mergeTerm.equals("$250") || mergeTerm.equals("$39") || mergeTerm.equals("$50")
							|| mergeTerm.equals("$35") || mergeTerm.equals("$10")) {
//						writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet, mergeTerm,
//								4);
					} else
						writeMergeTermDiscripentDataInFile(templateRowNumber, mergeTermDiscrepanciesSheet, mergeTerm, 1,
								rawTemplate);
				} else {
					if (!mergeTermsMap.get(mergeTerm).equals("NA") && !mergeTermsMap.get(mergeTerm).contains("??")) {
						templateString = templateString.replace(mergeTerm, mergeTermsMap.get(mergeTerm));
					} else {
//						System.out.println(
//								"Template Number :   AB" + templateRowNumber + "         Merge Term : " + mergeTerm);
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

		if (globalRowCounter == 1) {
			createNewRowInMergeTermSheetPopulateCells(mergeTermDiscrepanciesSheet, templateRowNumber, mergeTerm, reason,
					rawTemplate);
		} else {
			Iterator<Row> rowIterator = mergeTermDiscrepanciesSheet.iterator();
			Row row = rowIterator.next();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				String rowCell = row.getCell(0).getStringCellValue();
				String mergeTermCell = row.getCell(1).getStringCellValue();
				if (/*rowCell.equals("AB" + templateRowNumber)&& */ mergeTermCell.equals(mergeTerm) ) {
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
		Row newRow = mergeTermDiscrepanciesSheet.createRow(globalRowCounter++);
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

	private static void replaceTemplateStringInTemplateSheet(int templateRowNumber, XSSFSheet templateSheet,
			String templateString) {
		System.out.println("Print Row Number : " + templateRowNumber);
		Row row = templateSheet.getRow(templateRowNumber - 1);
		Cell templateCell = row.createCell(27);
		templateCell = row.getCell(27);
		templateCell.setCellValue(templateString);
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
//				String orgKeyCell = row.getCell(2).getStringCellValue();
				if (/* rowCell.equals("AB" + templateRowNumber) && */ tagCell
						.equals(localizedTermValueFromLocalizedSheet)/* && orgKeyCell.equals(orgStringValue) */) {
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
