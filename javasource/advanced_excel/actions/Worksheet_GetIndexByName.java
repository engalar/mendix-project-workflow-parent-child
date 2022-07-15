// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package advanced_excel.actions;

import com.mendix.core.Core;
import com.mendix.logging.ILogNode;
import java.io.IOException;
import advanced_excel.proxies.DocumentType;
import advanced_excel.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;
import com.mendix.systemwideinterfaces.core.IMendixObject;

/**
 * Get sheet index by name
 */
public class Worksheet_GetIndexByName extends CustomJavaAction<java.lang.Long>
{
	private java.lang.String SheetName;

	public Worksheet_GetIndexByName(IContext context, java.lang.String SheetName)
	{
		super(context);
		this.SheetName = SheetName;
	}

	@java.lang.Override
	public java.lang.Long executeAction() throws Exception
	{
		// BEGIN USER CODE
		try
		{
			DocumentType docType = Utils.GetDocumentType();
			Workbook workbook = Utils.GetWorkBook();
			Sheet sheet = workbook.getSheet(this.SheetName);
			if (sheet == null) { throw new Exception("Sheet: " + this.SheetName + " not found!"); }
			
			return new Long(workbook.getSheetIndex(sheet));
		} catch (Exception e) {
			logger.error("ERROR in Advanced_Excel.Worksheet_GetIndexByName: " + e.getMessage() + "\n" + e.toString(), e);
			return new Long(-1);
		} 
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "Worksheet_GetIndexByName";
	}

	// BEGIN EXTRA CODE
	protected static ILogNode logger = Core.getLogger("Advanced_Excel");
	// END EXTRA CODE
}
