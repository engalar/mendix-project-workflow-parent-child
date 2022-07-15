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
import java.text.SimpleDateFormat;
import java.util.Date;
import advanced_excel.proxies.CellType;
import advanced_excel.proxies.DocumentType;
import advanced_excel.Utils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.webui.CustomJavaAction;
import com.mendix.systemwideinterfaces.core.IMendixObject;

/**
 * Write a dataset to a range of cells
 */
public class Cell_WriteRange extends CustomJavaAction<java.lang.Boolean>
{
	private java.lang.String SheetName;
	private java.util.List<IMendixObject> __DataSet;
	private java.util.List<advanced_excel.proxies.Data> DataSet;
	private IMendixObject __CellFormat;
	private advanced_excel.proxies.CellFormat CellFormat;

	public Cell_WriteRange(IContext context, java.lang.String SheetName, java.util.List<IMendixObject> DataSet, IMendixObject CellFormat)
	{
		super(context);
		this.SheetName = SheetName;
		this.__DataSet = DataSet;
		this.__CellFormat = CellFormat;
	}

	@java.lang.Override
	public java.lang.Boolean executeAction() throws Exception
	{
		this.DataSet = new java.util.ArrayList<advanced_excel.proxies.Data>();
		if (__DataSet != null)
			for (IMendixObject __DataSetElement : __DataSet)
				this.DataSet.add(advanced_excel.proxies.Data.initialize(getContext(), __DataSetElement));

		this.CellFormat = __CellFormat == null ? null : advanced_excel.proxies.CellFormat.initialize(getContext(), __CellFormat);

		// BEGIN USER CODE
		try
		{
			String cellValue = "";
			CellType cellType;
			DocumentType docType = Utils.GetDocumentType();
			Workbook workbook = Utils.GetWorkBook();
			
			if (docType == DocumentType.XLS) {
				HSSFSheet sheet = ((HSSFWorkbook)workbook).getSheet(this.SheetName);
				if (sheet == null) { throw new Exception("Sheet: " + this.SheetName + " not found!"); }
				
				HSSFCellStyle style = ((HSSFWorkbook)workbook).createCellStyle();
				Utils.CreateStyle(workbook, style, this.CellFormat, docType);
				
				HSSFCellStyle dateStyle = ((HSSFWorkbook)workbook).createCellStyle();
				Utils.CreateStyle(workbook, dateStyle, this.CellFormat, docType);
				dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
				
				Boolean AutofitColumn = false;
				String CustomFormat = null;
				if (this.CellFormat != null)
				{
					AutofitColumn = this.CellFormat.getAutofitColumn();
					CustomFormat = this.CellFormat.getCustomFormat();
				}
				
				Integer rowNum = 0;
				Integer colNum = 0;
				for (advanced_excel.proxies.Data Data : DataSet)
				{
					cellValue = Data.getCellValue();
					cellType = Data.getCellType();
					rowNum = Data.getRow();
					colNum = Data.getColumn();
					HSSFRow row = sheet.getRow(rowNum);
					if (row == null) {
						row = sheet.createRow(rowNum);
					}
					HSSFCell cell = row.getCell(colNum);
					if (cell == null) {
						cell = row.createCell(colNum);
					}
					
					Utils.SetCellValue(workbook, cell, cellType, cellValue);
					if (CustomFormat != null) {
						style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(CustomFormat));
						dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(CustomFormat));
					}
					
					if (cellType == CellType.datetime) {
						cell.setCellStyle(dateStyle);
					} else if (cellValue != null && cellValue.indexOf('\n') != -1) {
						HSSFCellStyle newStyle = ((HSSFWorkbook)workbook).createCellStyle();
						newStyle.cloneStyleFrom(style);
						newStyle.setWrapText(true);
						cell.setCellStyle(newStyle);
					} else {
						cell.setCellStyle(style);
					}
					
					if (AutofitColumn) {
						sheet.autoSizeColumn(colNum);
					}
				}
			} else {
				XSSFSheet sheet = ((XSSFWorkbook)workbook).getSheet(this.SheetName);
				if (sheet == null) { throw new Exception("Sheet: " + this.SheetName + " not found!"); }
				
				XSSFCellStyle style = ((XSSFWorkbook)workbook).createCellStyle();
				Utils.CreateStyle(workbook, style, this.CellFormat, docType);
				
				XSSFCellStyle dateStyle = ((XSSFWorkbook)workbook).createCellStyle();
				Utils.CreateStyle(workbook, dateStyle, this.CellFormat, docType);
				dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
				
				Boolean AutofitColumn = false;
				String CustomFormat = null;
				if (this.CellFormat != null)
				{
					AutofitColumn = this.CellFormat.getAutofitColumn();
					CustomFormat = this.CellFormat.getCustomFormat();
				}
				
				Integer rowNum = 0;
				Integer colNum = 0;
				for (advanced_excel.proxies.Data Data : DataSet)
				{
					cellValue = Data.getCellValue();
					cellType = Data.getCellType();
					rowNum = Data.getRow();
					colNum = Data.getColumn();
					XSSFRow row = sheet.getRow(rowNum);
					if (row == null) {
						row = sheet.createRow(rowNum);
					}
					XSSFCell cell = row.getCell(colNum);
					if (cell == null) {
						cell = row.createCell(colNum);
					}
					
					Utils.SetCellValue(workbook, cell, cellType, cellValue);
					if (CustomFormat != null) {
						style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(CustomFormat));
						dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(CustomFormat));
					}
					
					if (cellType == CellType.datetime) {
						cell.setCellStyle(dateStyle);
					} else if (cellValue != null && cellValue.indexOf('\n') != -1) {
						XSSFCellStyle newStyle = ((XSSFWorkbook)workbook).createCellStyle();
						newStyle.cloneStyleFrom(style);
						newStyle.setWrapText(true);
						cell.setCellStyle(newStyle);
					} else {
						cell.setCellStyle(style);
					}
					
					if (AutofitColumn) {
						sheet.autoSizeColumn(colNum);
					}
				}
			}
			
			return true;
		} catch (Exception e) {
			logger.error("ERROR in Advanced_Excel.Cell_WriteRange: " + e.getMessage() + "\n" + e.toString(), e);
			return false;
		} 
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "Cell_WriteRange";
	}

	// BEGIN EXTRA CODE
	protected static ILogNode logger = Core.getLogger("Advanced_Excel");
	// END EXTRA CODE
}
