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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;
import com.mendix.systemwideinterfaces.core.IMendixObject;

/**
 * Reads the value of a cell, defined by its index.
 */
public class Cell_ReadByIndex extends CustomJavaAction<java.lang.String>
{
	private java.lang.String SheetName;
	private java.lang.Long Row;
	private java.lang.Long Column;

	public Cell_ReadByIndex(IContext context, java.lang.String SheetName, java.lang.Long Row, java.lang.Long Column)
	{
		super(context);
		this.SheetName = SheetName;
		this.Row = Row;
		this.Column = Column;
	}

	@java.lang.Override
	public java.lang.String executeAction() throws Exception
	{
		// BEGIN USER CODE
		try
		{
			DocumentType docType = Utils.GetDocumentType();
			Workbook workbook = Utils.GetWorkBook();
			Sheet sheet = workbook.getSheet(this.SheetName);
			if (sheet == null) { throw new Exception("Sheet: " + this.SheetName + " not found!"); }
			Integer rowNum = this.Row.intValue();
			Integer colNum = this.Column.intValue();
			
			Row row = sheet.getRow(rowNum);
			if (row == null) { return ""; }
			Cell cell = row.getCell(colNum);
			if (cell == null) { return ""; }
			
			DataFormatter formatter = new DataFormatter();
			return formatter.formatCellValue(cell);
		} catch (Exception e) {
			logger.error("ERROR in Advanced_Excel.Cell_ReadByIndex: " + e.getMessage() + "\n" + e.toString(), e);
			return "";
		} 
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "Cell_ReadByIndex";
	}

	// BEGIN EXTRA CODE
	protected static ILogNode logger = Core.getLogger("Advanced_Excel");
	// END EXTRA CODE
}
