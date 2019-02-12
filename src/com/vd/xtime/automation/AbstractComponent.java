/**
 * 
 */
package com.vd.xtime.automation;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Naeem Siddiq
 *
 *         ASE Venturedive
 */
public class AbstractComponent {

	protected static ArrayList<String> group1DealerList() throws Exception {
		String dealerGroupsFilePath = "C:\\Users\\Sakhi\\Desktop\\xTime\\dealer-groupings.xlsx";
		FileInputStream fileInputStream = new FileInputStream(new File(dealerGroupsFilePath));
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet dealersGroupingSheet = workbook.getSheetAt(0);
		fileInputStream.close();

		ArrayList<String> group1DealersList = new ArrayList<String>();
		Iterator<Row> rowIterator = dealersGroupingSheet.iterator();
		Row row = rowIterator.next();
		while (rowIterator.hasNext()) {
			row = rowIterator.next();
			Cell group1DealerCell = row.getCell(0);
			String dealer = "";
			if (group1DealerCell != null) {
				dealer = group1DealerCell.getStringCellValue().trim();
//				System.out.println("Row Number " + (row.getRowNum() + 1) + "   :" + dealer);
				group1DealersList.add(dealer);
			}
		}

//		System.out.println("\nGroup 1 Dealers List Count : " + group1DealersList.size());

		workbook.close();
		return group1DealersList;
	}
}
