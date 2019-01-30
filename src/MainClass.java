/**
 * @author Naeem Siddiq
 *
 *         ASE Venturedive
 */
public class MainClass {

	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		String LocalizationFilePath = "C:\\Users\\Sakhi\\Desktop\\xTime\\Phase 2 Notification Template Migration Data - Template and Localizations - Final - 1-28-19.xlsx";
		String glossaryFilePath = "C:\\Users\\Sakhi\\Desktop\\xTime\\Glossary - N6 _ N7 - All (1).xlsx";
		ReadLocalizationDataAndCreateMap.readLocalizationDataAndCreateMap(LocalizationFilePath, glossaryFilePath);
	}
}
