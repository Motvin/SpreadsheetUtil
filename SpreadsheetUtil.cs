using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
#if Allow_Office2003_UserModel
using NPOI.HSSF.UserModel;
#endif
using NPOI.XSSF.UserModel;



//[MethodImpl(MethodImplOptions.AggressiveInlining)]???

namespace SpreadsheetUtil
{
	public enum spWorkbookFileFormat
	{
		OfficeOpenXML = 0,
		Office2003 = 1
	}

	public enum spWorksheetVisibility
	{
		Visible = SheetState.Visible,
		Hidden = SheetState.Hidden,
		VeryHidden = SheetState.VeryHidden // very hidden means the user cannot make it visible - it can only be made visible in code
	}

	// Excel api has xlLineStyle enumeration which has some of the below, but there is also a xlBorderWeight that is hairline, medium, thick, thin
	public enum spBorderStyle : short
	{
		None = BorderStyle.None, // in xlLineStyle
		Thin = BorderStyle.Thin,
		Medium = BorderStyle.Medium,
		Dashed = BorderStyle.Dashed, // in xlLineStyle
		Dotted = BorderStyle.Dotted, // in xlLineStyle
		Thick = BorderStyle.Thick,
		Double = BorderStyle.Double, // in xlLineStyle
		Hair = BorderStyle.Hair,
		MediumDashed = BorderStyle.MediumDashed,
		DashDot = BorderStyle.DashDot, // in xlLineStyle
		MediumDashDot = BorderStyle.MediumDashDot,
		DashDotDot = BorderStyle.DashDotDot, // in xlLineStyle
		MediumDashDotDot = BorderStyle.MediumDashDotDot,
		SlantedDashDot = BorderStyle.SlantedDashDot // in xlLineStyle
	}

	// BorderDiagonal
	public enum spDiagonalBorderType
	{
		None = BorderDiagonal.None,
		DiagonalDown = BorderDiagonal.Backward,
		DiagonalUp = BorderDiagonal.Forward,
		Both = BorderDiagonal.Both
	}

	public enum spBorderSide //??? not sure this should be used
	{
		Top = NPOI.XSSF.UserModel.Extensions.BorderSide.TOP,
		Right = NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT,
		Bottom = NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM,
		Left = NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT,
		Diagonal = NPOI.XSSF.UserModel.Extensions.BorderSide.DIAGONAL
	}

	public enum spVerticalAlignment
	{
		None = VerticalAlignment.None,
		Top = VerticalAlignment.Top,
		Center = VerticalAlignment.Center,
		Bottom = VerticalAlignment.Bottom,
		Justify = VerticalAlignment.Justify,
		Distributed = VerticalAlignment.Distributed
	}

	public enum spHorizontalAlignment
	{
		General = HorizontalAlignment.General,
		Left = HorizontalAlignment.Left,
		Center = HorizontalAlignment.Center,
		Right = HorizontalAlignment.Right,
		Fill = HorizontalAlignment.Fill,
		Justify = HorizontalAlignment.Justify,
		CenterSelection = HorizontalAlignment.CenterSelection,
		Distributed = HorizontalAlignment.Distributed
	}

	public enum spFillPattern : short
	{
		NoFill = FillPattern.NoFill,
		SolidForeground = FillPattern.SolidForeground,
		FineDots = FillPattern.FineDots,
		AltBars = FillPattern.AltBars,
		SparseDots = FillPattern.SparseDots,
		ThickHorizontalBands = FillPattern.ThickHorizontalBands,
		ThickVerticalBands = FillPattern.ThickVerticalBands,
		ThickBackwardDiagonals = FillPattern.ThickBackwardDiagonals,
		ThickForwardDiagonals = FillPattern.ThickForwardDiagonals,
		BigSpots = FillPattern.BigSpots,
		Bricks = FillPattern.Bricks,
		ThinHorizontalBands = FillPattern.ThinHorizontalBands,
		ThinVerticalBands = FillPattern.ThinVerticalBands,
		ThinBackwardDiagonals = FillPattern.ThinBackwardDiagonals,
		ThinForwardDiagonals = FillPattern.ThinForwardDiagonals,
		Squares = FillPattern.Squares,
		Diamonds = FillPattern.Diamonds,
		LessDots = FillPattern.LessDots,
		LeastDots = FillPattern.LeastDots
	}

	//public class IndexedColors
	//{
	//	public static readonly IndexedColors Black;
	//	public static readonly IndexedColors PaleBlue;
	//	public static readonly IndexedColors Rose;
	//	public static readonly IndexedColors Lavender;
	//	public static readonly IndexedColors Tan;
	//	public static readonly IndexedColors LightBlue;
	//	public static readonly IndexedColors Aqua;
	//	public static readonly IndexedColors Lime;
	//	public static readonly IndexedColors Gold;
	//	public static readonly IndexedColors LightOrange;
	//	public static readonly IndexedColors Orange;
	//	public static readonly IndexedColors BlueGrey;
	//	public static readonly IndexedColors Grey40Percent;
	//	public static readonly IndexedColors DarkTeal;
	//	public static readonly IndexedColors SeaGreen;
	//	public static readonly IndexedColors DarkGreen;
	//	public static readonly IndexedColors OliveGreen;
	//	public static readonly IndexedColors Brown;
	//	public static readonly IndexedColors Plum;
	//	public static readonly IndexedColors Indigo;
	//	public static readonly IndexedColors Grey80Percent;
	//	public static readonly IndexedColors Automatic;
	//	public static readonly IndexedColors LightGreen;
	//	public static readonly IndexedColors LightTurquoise;
	//	public static readonly IndexedColors LightYellow;
	//	public static readonly IndexedColors LightCornflowerBlue;
	//	public static readonly IndexedColors White;
	//	public static readonly IndexedColors Red;
	//	public static readonly IndexedColors BrightGreen;
	//	public static readonly IndexedColors Blue;
	//	public static readonly IndexedColors Yellow;
	//	public static readonly IndexedColors Pink;
	//	public static readonly IndexedColors Turquoise;
	//	public static readonly IndexedColors DarkRed;
	//	public static readonly IndexedColors Green;
	//	public static readonly IndexedColors SkyBlue;
	//	public static readonly IndexedColors DarkYellow;
	//	public static readonly IndexedColors DarkBlue;
	//	public static readonly IndexedColors Teal;
	//	public static readonly IndexedColors Grey25Percent;
	//	public static readonly IndexedColors Grey50Percent;
	//	public static readonly IndexedColors CornflowerBlue;
	//	public static readonly IndexedColors Maroon;
	//	public static readonly IndexedColors LemonChiffon;
	//	public static readonly IndexedColors Orchid;
	//	public static readonly IndexedColors Coral;
	//	public static readonly IndexedColors RoyalBlue;
	//	public static readonly IndexedColors Violet;

	//	public string HexString { get; }
	//	public byte[] RGB { get; }
	//	public short Index { get; }

	//	public static IndexedColors ValueOf(string colorName);
	//	public static IndexedColors ValueOf(int index);
	//}

    public static class SpreadsheetUtil
	{
 		public static ICellStyle GetPreferredCellStyle(ICell cell)
		{
			// a method to get the preferred cell style for a cell
			// this is either the already applied cell style
			// or if that not present, then the row style (default cell style for this row)
			// or if that not present, then the column style (default cell style for this column)
			ICellStyle cellStyle = cell.CellStyle;
			if (cellStyle.Index == 0) cellStyle = cell.Row.RowStyle;
			if (cellStyle == null) cellStyle = cell.Sheet.GetColumnStyle(cell.ColumnIndex);
			if (cellStyle == null) cellStyle = cell.CellStyle;
			return cellStyle;
		}
	}

	public sealed class Workbook
	{
		internal IWorkbook wb;

		public Workbook(spWorkbookFileFormat format = spWorkbookFileFormat.OfficeOpenXML)
		{
			if (format == spWorkbookFileFormat.OfficeOpenXML)
			{
				wb = new XSSFWorkbook(); //??? there are alot of members of XSSFWorkbook/HSSFWorkbook that aren't a part of IWorkbook - and indexers and worksheet collections - should we include those as well?
				//???wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());
			}
#if Allow_Office2003_UserModel
			else
			{
				wb = new HSSFWorkbook();
			}
#endif
			// ??? the line below works, but maybe there are cases where we want to know a missing cell vs. blank and we don't just want to start creating additional cells
			//wb.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
		}

		public int WorksheetCount => wb.NumberOfSheets;
		//public int WorksheetCount
		//{
		//	get
		//	{
		//		return wb.NumberOfSheets;
		//	}
		//}

		public Worksheet GetWorksheet(int worksheetIndex)
		{
			return new Worksheet(wb.GetSheetAt(worksheetIndex));
		}

		public Worksheet GetWorksheet(string worksheetName)
		{
			return new Worksheet(wb.GetSheet(worksheetName));
		}

		//public int GetWorksheetIndex(Worksheet ws)
		//{
		//	return wb.GetSheetIndex(ws.ws);
		//}

		// returns 0 based worksheet index or -1 if no worksheet found
		public int GetWorksheetIndex(string worksheetName)
		{
			return wb.GetSheetIndex(worksheetName);
		}

		public int GetFirstVisibleWorksheetIndex()
		{
			return wb.FirstVisibleTab;
		}

		public int GetActiveWorksheetIndex()
		{
			return wb.ActiveSheetIndex;
		}

		public void SetActiveWorksheet(int worksheetIndex)
		{
			wb.SetActiveSheet(worksheetIndex);
		}

		public void SetSelectedWorksheet(int worksheetIndex)
		{
			//??? can't there be more than one selected worksheet? and how can this be unselected (select a different tab)? or get the selected tab or selection state of a tab?
			wb.SetSelectedTab(worksheetIndex);
		}

		// ??? not sure how creating the format (and cell styles) from the workbook object works - maybe there are a limited number of these?
		// try to run this function and see how it works - really should be able to set the formatString of a cell object instead of from the workbook object
		//public void SetCellFormat(Worksheet ws, int rowIndex, int colIndex, string formatString)
		//{
		//	IDataFormat format = wb.CreateDataFormat();

		//	ws.GetOrCreateCellInRow(rowIndex, colIndex).CellStyle.DataFormat = format.GetFormat(formatString);
		//}

		public string GetWorksheetName(int worksheetIndex)
		{
			return wb.GetSheetName(worksheetIndex);
		}

		public void SetWorksheetName(int worksheetIndex, string worksheetName)
		{
			wb.SetSheetName(worksheetIndex, worksheetName);
		}

		public void SetWorksheetOrder(string worksheetName, int newWorksheetIndex)
		{
			wb.SetSheetOrder(worksheetName, newWorksheetIndex);
		}

		public Worksheet AddWorksheet(string worksheetName = null)
		{
			Worksheet ws;
			if (worksheetName != null)
			{
				ws = new Worksheet(wb.CreateSheet(worksheetName));
			}
			else
			{
				ws = new Worksheet(wb.CreateSheet());
			}
			return ws;
		}

		public void RemoveWorksheet(int worksheetIndex)
		{
			wb.RemoveSheetAt(worksheetIndex);
		}

		//??? instead of the Is.. Hidden/VeryHidden if there was a GetWorksheetVisiblity - it seems like that isn't part of npoi IWorkbook - not sure why - could add this function with logic from IsSheetHidden/VeryHidden, etc.
		
		//??? remove the following 2 functions
		//[Obsolete]
		//public bool IsWorksheetHidden(int worksheetIndex)
		//{
		//	return wb.IsSheetHidden(worksheetIndex);
		//}

		//[Obsolete]
		//public bool IsWorksheetVeryHidden(int worksheetIndex)
		//{
		//	return wb.IsSheetVeryHidden(worksheetIndex);
		//}

		//??? probably have a Worksheets property - similar to Excel interop

		//??? change some Get/Set functions to properties?
		//public spWorksheetVisibility GetWorksheetVisibility(int worksheetIndex)
		//{
		//	spWorksheetVisibility visibility = spWorksheetVisibility.Visible;

		//	if (wb.IsSheetHidden(worksheetIndex))
		//	{
		//		//??? I think hidden is not a superset of very hidden - it is a different state
		//		visibility = spWorksheetVisibility.Hidden;
		//	}
		//	else if (wb.IsSheetVeryHidden(worksheetIndex))
		//	{
		//		visibility = spWorksheetVisibility.VeryHidden;
		//	}
		//	return visibility;
		//}

		//public void SetWorksheetVisibility(int worksheetIndex, spWorksheetVisibility visibility)
		//{
		//	wb.SetSheetHidden(worksheetIndex, (SheetState)visibility);
		//}

		public void Save()
		{
			//??? need to keep the filename as a member variable when Opened? or saved as - or is this available somehow?
		}

		public void SaveAs(string fileName)
		{
			using (FileStream stream = File.Open(fileName, FileMode.Create, FileAccess.Write))
			{
				wb.Write(stream);
			}
		}

		public void SaveAs(Stream stream)
		{
			wb.Write(stream);
		}

		public void Close()
		{
			wb.Close();
			wb = null;
		}
	}

	public sealed class Worksheet
	{
		//internal ISheet ws; // Note: an ISheet can probably be a worksheet or a chart?, but here we are using it only for a worksheet???
		public ISheet ws; // Note: an ISheet can probably be a worksheet or a chart?, but here we are using it only for a worksheet???
		//???what if a worksheet can be moved to another workbook?  private IWorkbook parentWb;
		// IWorkbook Workbook { get; } gets the internal parent workbook? - maybe this is all we need

		// make this internal so that we can't create worksheets outside of this assembly - is this really the way to do it???
		internal Worksheet(ISheet sheet)
		{
			ws = sheet;
		}

		public int WorksheetIndex
		{
			get => ws.Workbook.GetSheetIndex(ws);
		}

		public void SetActiveWorksheet()
		{
			ws.Workbook.SetActiveSheet(ws.Workbook.GetSheetIndex(ws));
		}

		public void SetSelectedWorksheet()
		{
			//??? can't there be more than one selected worksheet? and how can this be unselected (select a different tab)? or get the selected tab or selection state of a tab?
			ws.Workbook.SetSelectedTab(ws.Workbook.GetSheetIndex(ws));
		}

		public void SetWorksheetNameSafe(string name)
		{
			name = WorkbookUtil.CreateSafeSheetName(name, '_');

			try
			{
				ws.Workbook.SetSheetName(ws.Workbook.GetSheetIndex(ws), name);
			}
			catch //??? i think this throws invalidargument exception if the sheet name is not unique just catch this and the original unique name (like Sheet 1) will remain
			{
				
			}

		}

		public void SetFillBackgroundColor(int rowIndex, int colIndex, byte red, byte green, byte blue)
		{
			ICellStyle cellStyle = SpreadsheetUtil.GetPreferredCellStyle(ws.GetRow(rowIndex).GetCell(colIndex));

			//??? find the CellStyleValue from the cellStyle.Index - does Index work for XSSF workbooks, or do you need UIndex?
			int idx = cellStyle.Index; // if this is 0, you can't modify it because it is the default and it won't get written out to the file

			#if false
			CellStyleValue existingValue;
			if (idx == 0)
			{
				//??? don't decrement ref count of default cell style or try to remove it
				existingValue = defaultCellStyleValue;
			}
			else
			{
				//???
				existingValue = defaultCellStyleValue;
			}
			CellStyleValue changedValue = existingValue;
			changedValue.SetFillBackgroundColor(red, green, blue);
			#endif
		}

		// Note: Thin seems to be the default when setting borders in Excel's Format Cells screen and it looks ok, so use it as the default below
		public void SetBorder(int rowIndex, int colIndex, spBorderSide side, spBorderStyle style = spBorderStyle.Thin)
		{
			ICellStyle cellStyle = SpreadsheetUtil.GetPreferredCellStyle(ws.GetRow(rowIndex).GetCell(colIndex));

			//??? find the CellStyleValue from the cellStyle.Index - does Index work for XSSF workbooks, or do you need UIndex?
			int idx = cellStyle.Index; // if this is 0, you can't modify it because it is the default and it won't get written out to the file

			#if false
			CellStyleValue existingValue;
			if (idx == 0)
			{
				//??? don't decrement ref count of default cell style or try to remove it
				existingValue = defaultCellStyleValue;
			}
			else
			{
				//???
				existingValue = defaultCellStyleValue;
			}
			CellStyleValue changedValue = existingValue;
			changedValue.SetFillBackgroundColor(red, green, blue);
			#endif
		}

		// ??? not sure how creating the format (and cell styles) from the workbook object works - maybe there are a limited number of these?
		// try to run this function and see how it works - really should be able to set the formatString of a cell object instead of from the workbook object
		public void SetCellFormat(int rowIndex, int colIndex, string formatString)
		{
			IDataFormat format = ws.Workbook.CreateDataFormat();
			ICellStyle cellStyle = ws.Workbook.CreateCellStyle();
			cellStyle.DataFormat = format.GetFormat(formatString);

			ICell cell = GetOrCreateCellInRow(rowIndex, colIndex);
			cell.CellStyle = cellStyle;

		}

		public void SetCellFormat2(int rowIndex, int colIndex, string formatString)
		{
			IDataFormat format = ws.Workbook.CreateDataFormat();
			//ICellStyle cellStyle = ws.Workbook.CreateCellStyle();
			//ws.Workbook.MissingCellPolicy._policy = MissingCellPolicy.Policy.RETURN_NULL_AND_BLANK;

			//ICellStyle cellStyle = ws.Workbook.CreateCellStyle();
			//cellStyle.DataFormat = format.GetFormat(formatString);

			//GetOrCreateCellInRow(rowIndex, colIndex).CellStyle.DataFormat = format.GetFormat(formatString);
			ICell cell = GetOrCreateCellInRow(rowIndex, colIndex);

			//??? maybe the first cell style is reserved
			cell.CellStyle.DataFormat = format.GetFormat(formatString);;
		}

//import java.io.*;

//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.*;

//import org.apache.poi.ss.util.CellUtil;

//import java.util.Map;
//import java.util.HashMap;

//public class CarefulCreateCellStyles {

// public CellStyle getPreferredCellStyle(Cell cell) {
//  // a method to get the preferred cell style for a cell
//  // this is either the already applied cell style
//  // or if that not present, then the row style (default cell style for this row)
//  // or if that not present, then the column style (default cell style for this column)
//  CellStyle cellStyle = cell.getCellStyle();
//  if (cellStyle.getIndex() == 0) cellStyle = cell.getRow().getRowStyle();
//  if (cellStyle == null) cellStyle = cell.getSheet().getColumnStyle(cell.getColumnIndex());
//  if (cellStyle == null) cellStyle = cell.getCellStyle();
//  return cellStyle;
// }

// public CarefulCreateCellStyles() throws Exception {

//   Workbook workbook = new XSSFWorkbook();

//   // at first we are creating needed fonts
//   Font defaultFont = workbook.createFont();
//   defaultFont.setFontName("Arial");
//   defaultFont.setFontHeightInPoints((short)14);

//   Font specialfont = workbook.createFont();
//   specialfont.setFontName("Courier New");
//   specialfont.setFontHeightInPoints((short)18);
//   specialfont.setBold(true);

//   // now we are creating a default cell style which will then be applied to all cells
//   CellStyle defaultCellStyle = workbook.createCellStyle();
//   defaultCellStyle.setFont(defaultFont);

//   // maybe sone rows need their own default cell style
//   CellStyle aRowCellStyle = workbook.createCellStyle();
//   aRowCellStyle.cloneStyleFrom(defaultCellStyle);
//   aRowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//   aRowCellStyle.setFillForegroundColor((short)3);


//   Sheet sheet = workbook.createSheet("Sheet1");

//   // apply default cell style as column style to all columns
//   org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol cTCol = 
//      ((XSSFSheet)sheet).getCTWorksheet().getColsArray(0).addNewCol();
//   cTCol.setMin(1);
//   cTCol.setMax(workbook.getSpreadsheetVersion().getLastColumnIndex());
//   cTCol.setWidth(20 + 0.7109375);
//   cTCol.setStyle(defaultCellStyle.getIndex());

//   // creating cells
//   Row row = sheet.createRow(0);
//   row.setRowStyle(aRowCellStyle);
//   Cell cell = null;
//   for (int c = 0; c  < 3; c++) {
//    cell = CellUtil.createCell(row, c, "Header " + (c+1));
//    // we get the preferred cell style for each cell we are creating
//    cell.setCellStyle(getPreferredCellStyle(cell));
//   }

//   System.out.println(workbook.getNumCellStyles()); // 3 = 0(default) and 2 just created

//   row = sheet.createRow(1);
//   cell = CellUtil.createCell(row, 0, "centered");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);

//   System.out.println(workbook.getNumCellStyles()); // 4 = 0 and 3 just created

//   cell = CellUtil.createCell(row, 1, "bordered");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   Map<String, Object> properties = new HashMap<String, Object>();
//   properties.put(CellUtil.BORDER_LEFT, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_RIGHT, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_TOP, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.THICK);
//   CellUtil.setCellStyleProperties(cell, properties);

//   System.out.println(workbook.getNumCellStyles()); // 5 = 0 and 4 just created

//   cell = CellUtil.createCell(row, 2, "other font");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   CellUtil.setFont(cell, specialfont);

//   System.out.println(workbook.getNumCellStyles()); // 6 = 0 and 5 just created

//// until now we have always created new cell styles. but from now on CellUtil will use
//// already present cell styles if they matching the needed properties.

//   row = sheet.createRow(2);
//   cell = CellUtil.createCell(row, 0, "bordered");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   properties = new HashMap<String, Object>();
//   properties.put(CellUtil.BORDER_LEFT, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_RIGHT, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_TOP, BorderStyle.THICK);
//   properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.THICK);
//   CellUtil.setCellStyleProperties(cell, properties);

//   System.out.println(workbook.getNumCellStyles()); // 6 = nothing new created

//   cell = CellUtil.createCell(row, 1, "other font");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   CellUtil.setFont(cell, specialfont);

//   System.out.println(workbook.getNumCellStyles()); // 6 = nothing new created

//   cell = CellUtil.createCell(row, 2, "centered");
//   cell.setCellStyle(getPreferredCellStyle(cell));
//   CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);

//   System.out.println(workbook.getNumCellStyles()); // 6 = nothing new created


//   workbook.write(new FileOutputStream("CarefulCreateCellStyles.xlsx"));
//   workbook.close();  
// }

// public static void main(String[] args) throws Exception {
//  CarefulCreateCellStyles carefulCreateCellStyles = new CarefulCreateCellStyles();
// }
//}
		public short TestCreateCellStyle()
		{
			ICellStyle intCellStyle = ws.Workbook.CreateCellStyle();
			return ws.Workbook.NumCellStyles;
		}

		public void TestCreateDataFormat()
		{
			IDataFormat fmt = ws.Workbook.CreateDataFormat();
		}

		public void TestSetCellStyle(int rowIndex, int colIndex)
		{
			ICellStyle intCellStyle = ws.Workbook.CreateCellStyle();
			intCellStyle.DataFormat = ws.Workbook.CreateDataFormat().GetFormat( "#,##0" );
			ws.GetRow(rowIndex).GetCell(colIndex).CellStyle = intCellStyle;
		}

		public ICellStyle CreateOrGetExistingCellStyle(ICellStyle fromNewCellStyle)
		{
			int colIndex = 0;
			IRow row = null;
			ICell cell = CellUtil.CreateCell(row, colIndex, "");
			CellUtil.SetCellStyleProperty(cell, ws.Workbook, CellUtil.TOP_BORDER_COLOR, null);//??? how does this work?
			// look in dictionary for matching string key of cell style - if present then reuse this cell style, otherwise create new cell style in workbook
			ICellStyle cellStyle = null;

			return cellStyle;
		}

		public int GetLastUsedRowIndex()
		{
			return ws.LastRowNum;
		}

		public int GetLastUsedColumnIndexOfRow(int rowIndex)
		{
			int lastUsedColIndex;
			IRow row = ws.GetRow(rowIndex);
			if (row != null)
			{
				lastUsedColIndex = row.LastCellNum;
			}
			else
			{
				lastUsedColIndex = -1;
			}
			return lastUsedColIndex;
		}

		public spWorksheetVisibility Visibility
		{
			get
			{
				spWorksheetVisibility visibility = spWorksheetVisibility.Visible;

				int worksheetIndex = WorksheetIndex;
				if (ws.Workbook.IsSheetHidden(worksheetIndex))
				{
					//??? I think hidden is not a superset of very hidden - it is a different state
					visibility = spWorksheetVisibility.Hidden;
				}
				else if (ws.Workbook.IsSheetVeryHidden(worksheetIndex))
				{
					visibility = spWorksheetVisibility.VeryHidden;
				}
				return visibility;
			}

			set
			{
				ws.Workbook.SetSheetHidden(WorksheetIndex, (SheetState)value);
			}
		}

		//??? why can't we set this - could the ws get it's wb and then we can set from there? - then could get rid of the wb of these? - at least using name
		public string Name
		{
			get => ws.SheetName;

			// not sure why you can get ws.SheetName, but not set it??? - will this throw an exception if the name is not unique - what about if the name is too long, or has invalid chars
			set => ws.Workbook.SetSheetName(ws.Workbook.GetSheetIndex(ws), value);
		}

		public void FreezePane(int rowIndex, int colIndex)
		{
			//??? not sure why this is col, row and everything else is row, col
			ws.CreateFreezePane(colIndex, rowIndex);
		}

		public void SetActiveCell(int rowIndex, int colIndex)
		{
			ws.SetActiveCell(rowIndex, colIndex);
		}

		public void MergeCells(int rowIndex1, int colIndex1, int rowIndex2, int colIndex2)
		{
			// ??? this returns an int, which looks like it is something that can get passed to GetMergedRegion?
			CellRangeAddress c = new CellRangeAddress(rowIndex1, rowIndex2, colIndex1, colIndex2);
			ws.AddMergedRegion(c);
		}

		//??? do the same for rows - is there something in npoi for this? - I don't think so, but the Excel interop has worksheet.Rows.AutoFit(); to autofit all rows in the worksheet or do this for a range to autofit only a range of rows
		public void AutoSizeColumn(int colIndex, bool useMergedCells = false)
		{
			ws.AutoSizeColumn(colIndex, useMergedCells); //??? this does not work well
		}

		public void AutoSizeColumns(int colIndex1, int colIndex2, bool useMergedCells = false)
		{
			for (int col = colIndex1; col <= colIndex2; col++)
			{
				ws.AutoSizeColumn(col, useMergedCells);
			}
		}

		public void AutoSizeColumns(bool useMergedCells = false)
		{
			// how can we get the used cols?
			//???ws.AutoSizeColumn(column, useMergedCells);
		}

		public int GetColumnWidth(int colIndex)
		{
			return ws.GetColumnWidth(colIndex);
		}

		public float GetColumnWidthInPixels(int colIndex)
		{
			return ws.GetColumnWidthInPixels(colIndex);
			//???ws.AutoSizeColumn(column, useMergedCells);
		}

		public void SetColumnWidth(int colIndex, int width)
		{
			ws.SetColumnWidth(colIndex, width);
		}

		internal IRow GetOrCreateRow(int rowIndex)
		{
			IRow row = ws.GetRow(rowIndex);
			if (row == null)
			{
				row = ws.CreateRow(rowIndex);
			}
			return row;
		}

		internal ICell GetOrCreateCell(IRow row, int colIndex)
		{
			ICell cell = row.GetCell(colIndex);
			if (cell == null)
			{
				cell = row.CreateCell(colIndex);
			}
			return cell;
		}

		internal ICell GetOrCreateCellInRow(int rowIndex, int colIndex)
		{
			IRow row = ws.GetRow(rowIndex);
			if (row == null)
			{
				row = ws.CreateRow(rowIndex);
			}

			//ICell cell = row.GetCell(colIndex);
			ICell cell = row.GetCell(colIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
			//if (cell == null)
			//{
			//	cell = row.CreateCell(colIndex);
			//}
			return cell;
		}

		public void SetRowHeight(int rowIndex, short height)
		{
			GetOrCreateRow(rowIndex).Height = height;
		}

		public short GetRowHeight(int rowIndex)
		{
			return GetOrCreateRow(rowIndex).Height;
		}

		//??? already have SetCellValue functions below, not sure this should be called SetCell - maybe it should set the value and the style in one call?
		//public void SetCell(int zeroBasedRowIndex, int zerofBasedColumnIndex, double val)
		//{
		//	ws.GetRow(zeroBasedRowIndex).GetCell(zerofBasedColumnIndex).SetCellValue(val);
		//}

		//public void SetCell(int zeroBasedRowIndex, int zerofBasedColumnIndex, string val)
		//{
		//	ws.GetRow(zeroBasedRowIndex).GetCell(zerofBasedColumnIndex).SetCellValue(val);
		//}

		//public void SetCell(int zeroBasedRowIndex, int zerofBasedColumnIndex, bool val)
		//{
		//	ws.GetRow(zeroBasedRowIndex).GetCell(zerofBasedColumnIndex).SetCellValue(val);
		//}

		//public void SetCell(int zeroBasedRowIndex, int zerofBasedColumnIndex, DateTime val)
		//{
		//	ws.GetRow(zeroBasedRowIndex).GetCell(zerofBasedColumnIndex).SetCellValue(val);
		//	ICell c;
		//}


		public void SetValue(int rowIndex, int colIndex, bool val)
		{
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellValue(val);
		}

		public void SetValue(int rowIndex, int colIndex, string val)
		{
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellValue(val);
		}

		public void SetValue(int rowIndex, int colIndex, DateTime val)
		{
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellValue(val);
		}

		public void SetValue(int rowIndex, int colIndex, double val)
		{
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellValue(val);
		}

		public void SetValue(int rowIndex, int colIndex, int val)
		{
			SetValue(rowIndex, colIndex, (double)val);
		}

		public void SetValue(int rowIndex, int colIndex, decimal val)
		{
			//??? put this code in a separate function and call it from here and the SetValue in the WorksheetCell
			double valDbl = 0;
			try
			{
				valDbl = (double)val;
			}
			catch
			{
				//??? do something if there is an exception
			}
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellValue(valDbl);
		}

		//??? what about IRichTextString?
		//public void SetValue(int row, int col, double val)
		//{
		//	ws.GetRow(row).GetCell(col).SetCellValue(val);
		//}
		
		public void SetFormula(int rowIndex, int colIndex, string formula)
		{
			GetOrCreateCellInRow(rowIndex, colIndex).SetCellFormula(formula);
		}

		public WorksheetCell GetCell(int rowIndex, int colIndex)
		{
			return new WorksheetCell(GetOrCreateCellInRow(rowIndex, colIndex));
		}

		public void SetRowCellStyle(int rowIndex, WorkbookCellStyle cellStyle)
		{
			//IRow row = ;
			//row.RowStyle = cellStyle;
		}

		public void SetColumnCellStyle(int colIndex, WorkbookCellStyle cellStyle)
		{
			//ICell cell = null;
			//ws.cell
			//cell.RowStyle = cellStyle;
		}

		//setRowStyle works the same as setDefaultColumnStyle: This actually sets the (default) style for cells that are added manually after the workbook has been exported already (i.e. by a human).
		//public void applyStyleToRange(Sheet sheet, CellStyle style, int rowStart, int colStart, int rowEnd, int colEnd) {
  //  for (int r = rowStart; r <= rowEnd; r++) {
  //      for (int c = colStart; c <= colEnd; c++) {
  //          Row row = sheet.getRow(r);

  //          if (row != null) {
  //              Cell cell = row.getCell(c);

  //              if (cell != null) {
  //                  cell.setCellStyle(style);
  //              }
  //          }
  //      }
  //  }
	}

	public sealed class WorksheetCell
	{
		internal ICell cell; // Note: an ISheet can probably be a worksheet or a chart?, but here we are using it only for a worksheet???

		// make this internal so that we can't create WorksheetCells outside of this assembly - is this really the way to do it???
		internal WorksheetCell(ICell cell)
		{
			this.cell = cell;
		}

		public spVerticalAlignment VerticalAlignment
		{
			get => (spVerticalAlignment)cell.CellStyle.VerticalAlignment;

			set => cell.CellStyle.VerticalAlignment = (VerticalAlignment)value;
		}

		public spHorizontalAlignment HorizontalAlignment
		{
			get => (spHorizontalAlignment)cell.CellStyle.Alignment;

			set => cell.CellStyle.Alignment = (HorizontalAlignment)value;
		}

		public bool WrapText
		{
			get => cell.CellStyle.WrapText;

			set => cell.CellStyle.WrapText = value;
		}

		public bool IsMerged => cell.IsMergedCell;

		public void SetValue(double val)
		{
			cell.SetCellValue(val);
		}

		public void SetValue(decimal val)
		{
			//??? put this code in a separate function and call it from here and the SetValue in the WorksheetCell
			double valDbl = 0;
			try
			{
				valDbl = (double)val;
			}
			catch
			{
				//??? do something if there is an exception
			}
			cell.SetCellValue(valDbl);
		}

		public void SetValue(string val)
		{
			cell.SetCellValue(val);
		}

		public void SetValue(bool val)
		{
			cell.SetCellValue(val);
		}

		public void SetValue(DateTime val)
		{
			cell.SetCellValue(val);
		}

		public void SetFormula(string formula)
		{
			cell.SetCellFormula(formula);
		}

		//??? this is somewhat confusing and not very useful, so don't have this function
		//public void SetBorders(CellBorderStyle left, CellBorderStyle top, CellBorderStyle right, CellBorderStyle bottom)
		//{
		//	cell.CellStyle.BorderLeft = (BorderStyle)left;
		//	cell.CellStyle.BorderTop = (BorderStyle)top;
		//	cell.CellStyle.BorderRight = (BorderStyle)right;
		//	cell.CellStyle.BorderBottom = (BorderStyle)bottom;
		//}

		public void SetBorders(spBorderStyle style)
		{
			cell.CellStyle.BorderLeft = (BorderStyle)style;
			cell.CellStyle.BorderTop = (BorderStyle)style;
			cell.CellStyle.BorderRight = (BorderStyle)style;
			cell.CellStyle.BorderBottom = (BorderStyle)style;
		}

		// if setting rgb color, it must be the XSSF type of workbook, so we can cast to XSSFCellStyle below, otherwise this function shouldnt' be called
		public void SetBorder(spBorderSide side, byte red, byte green, byte blue)
		{
			XSSFCellStyle s = (XSSFCellStyle)cell.CellStyle;
			XSSFColor clr = new XSSFColor();

			clr.SetRgb(new byte[] { red, green, blue });
			s.SetBorderColor((NPOI.XSSF.UserModel.Extensions.BorderSide)side, clr);
		}

		public void SetBackgroundColor(byte red, byte green, byte blue)
		{
			XSSFCellStyle s = (XSSFCellStyle)cell.CellStyle;
			XSSFColor clr = new XSSFColor();

			clr.SetRgb(new byte[] { red, green, blue });
			s.SetFillBackgroundColor(clr);
		}

		public void SetForegroundColor(byte red, byte green, byte blue)
		{
			XSSFCellStyle s = (XSSFCellStyle)cell.CellStyle;
			XSSFColor clr = new XSSFColor();

			clr.SetRgb(new byte[] { red, green, blue });
			s.SetFillForegroundColor(clr);
		}

		public short Rotation
		{
			get => cell.CellStyle.Rotation;
			set => cell.CellStyle.Rotation = value;
		}
	}

// Format cell into date time format with "m/d/yy h:mm", 
// which is converted to user's locale in Excel. 
//CellUtil.setCellStyleProperties(cell, Collections.singletonMap(CellUtil.DATA_FORMAT, (short) 22));

	//??? this is a bad name, maybe name this spColor, and change other names to have sp in front, spWorkbook, spWorksheet, spRow, spCell, etc. - don't do this, instead set a namespace which has sp as the last part so you could refer to things like sp.Workbook if there was a name clash
	public struct CellColor
	{
		private uint colorRGB; // use a uint which is faster to compare and hash than 3 bytes?
		//??? add Equals, GetHashCode, etc.

		//??? do we need other things like theme, tint, etc. - not sure how these work - what about gradients?

		private const uint RedBitMask =		0x00FF0000;
		private const uint GreenBitMask =	0x0000FF00;
		private const uint BlueBitMask =	0x000000FF;

		private const uint RedBitMaskInverted =		~RedBitMask;
		private const uint GreenBitMaskInverted =	~GreenBitMask;
		private const uint BlueBitMaskInverted =	~BlueBitMask;

		public CellColor(byte red, byte green, byte blue)
		{
			colorRGB = 0;
			Red = red;
			Green = green;
			Blue = blue;
		}

		public void SetRgb(byte red, byte green, byte blue)
		{
			Red = red;
			Green = green;
			Blue = blue;
		}

		public byte Red
		{
			get
			{
				return (byte)((colorRGB | RedBitMask) >> 16);
			}

			set
			{
				colorRGB = (colorRGB & RedBitMaskInverted) | ((uint)value << 16);
			}
		}

		public byte Green
		{
			get
			{
				return (byte)((colorRGB | GreenBitMask) >> 8);
			}

			set
			{
				colorRGB = (colorRGB & GreenBitMaskInverted) | ((uint)value << 8);
			}
		}

		public byte Blue
		{
			get
			{
				return (byte)((colorRGB | BlueBitMask));
			}

			set
			{
				colorRGB = (colorRGB & BlueBitMaskInverted) | ((uint)value);
			}
		}

		public override bool Equals(object obj)
		{
			CellColor c = (CellColor)obj;
			return Equals(c);
		}

		public override int GetHashCode()
		{
			return (int)colorRGB;
		}

		public bool Equals(CellColor c)
		{
			return colorRGB == c.colorRGB;
		}
		
        public static bool operator ==(CellColor c1, CellColor c2)
        {
            return c1.Equals(c2);
        }

        public static bool operator !=(CellColor c1, CellColor c2)
        {
            return !(c1 == c2);
        }
	}

	public struct CellStyleValue : IEquatable<CellStyleValue>
	{
		private bool isCachedHashCodeValid;
		private int cachedHashCode;
		private uint valueBits; // this is compact (using bits) form of the 4 boolean properties, and all of the enum properties except for FillPattern

		private const uint IsHiddenBit = 0x00000001;
		private const uint WrapTextBit = 0x00000002;
		private const uint ShrinkToFitBit = 0x00000004;
		private const uint IsLockedBit = 0x00000008;

		private const uint IsHiddenBitInverted = ~IsHiddenBit;
		private const uint WrapTextBitInverted = ~WrapTextBit;
		private const uint ShrinkToFitBitInverted = ~ShrinkToFitBit;
		private const uint IsLockedBitInverted = ~IsLockedBit;

		private const uint LeftBorderStyleBitMask =		0x000000F0; // 4 bits
		private const uint TopBorderStyleBitMask =		0x00000F00; // 4 bits
		private const uint RightBorderStyleBitMask =	0x0000F000; // 4 bits
		private const uint BottomBorderStyleBitMask =	0x000F0000; // 4 bits
		private const uint DiagonalBorderStyleBitMask =	0x00F00000; // 4 bits
		private const uint DiagonalBorderTypeBitMask =	0x03000000; // 2 bits
		private const uint HorizontalAlignmentBitMask =	0x03000000; // 3 bits
		private const uint VerticalAlignmentBitMask =	0x03000000; // 3 bits
		
		private const uint LeftBorderStyleBitMaskInverted = ~LeftBorderStyleBitMask;
		private const uint TopBorderStyleBitMaskInverted = ~TopBorderStyleBitMask;
		private const uint RightBorderStyleBitMaskInverted = ~RightBorderStyleBitMask;
		private const uint BottomBorderStyleBitMaskInverted = ~BottomBorderStyleBitMask;
		private const uint DiagonalBorderStyleBitMaskInverted = ~DiagonalBorderStyleBitMask;
		private const uint DiagonalBorderTypeBitMaskInverted = ~DiagonalBorderTypeBitMask;
		private const uint HorizontalAlignmentBitMaskInverted = ~HorizontalAlignmentBitMask;
		private const uint VerticalAlignmentBitMaskInverted = ~VerticalAlignmentBitMask;

		//??? can't have parameterless constructor
		//private CellStyleValue()
		//{
		//	cachedHashCode = CreateHashCode();
		//}

		//-------from ICellStyle
		public bool IsHidden
		{
			get
			{
				return (valueBits & IsHiddenBit) != 0;
			}
			
			set
			{
				if (value)
				{
					valueBits |= IsHiddenBit;
				}
				else
				{
					valueBits &= IsHiddenBitInverted;
				}
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public bool WrapText
		{
			get
			{
				return (valueBits & WrapTextBit) != 0;
			}
			
			set
			{
				if (value)
				{
					valueBits |= WrapTextBit;
				}
				else
				{
					valueBits &= WrapTextBitInverted;
				}
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public bool ShrinkToFit
		{
			get
			{
				return (valueBits & ShrinkToFitBit) != 0;
			}
			
			set
			{
				if (value)
				{
					valueBits |= ShrinkToFitBit;
				}
				else
				{
					valueBits &= ShrinkToFitBitInverted;
				}
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public bool IsLocked
		{
			get
			{
				return (valueBits & IsLockedBit) != 0;
			}
			
			set
			{
				if (value)
				{
					valueBits |= IsLockedBit;
				}
				else
				{
					valueBits &= IsLockedBitInverted;
				}
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spBorderStyle LeftBorderStyle
		{
			get
			{
				return (spBorderStyle)((valueBits & LeftBorderStyleBitMask) >> 4);
			}
			
			set
			{
				valueBits = (valueBits & LeftBorderStyleBitMaskInverted) | ((uint)value << 4);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spBorderStyle TopBorderStyle
		{
			get
			{
				return (spBorderStyle)((valueBits & TopBorderStyleBitMask) >> 8);
			}
			
			set
			{
				valueBits = (valueBits & TopBorderStyleBitMaskInverted) | ((uint)value << 8);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spBorderStyle RightBorderStyle
		{
			get
			{
				return (spBorderStyle)((valueBits & RightBorderStyleBitMask) >> 12);
			}
			
			set
			{
				valueBits = (valueBits & RightBorderStyleBitMaskInverted) | ((uint)value << 12);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spBorderStyle BottomBorderStyle
		{
			get
			{
				return (spBorderStyle)((valueBits & BottomBorderStyleBitMask) >> 16);
			}
			
			set
			{
				valueBits = (valueBits & BottomBorderStyleBitMaskInverted) | ((uint)value << 16);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spBorderStyle DiagonalBorderStyle
		{
			get
			{
				return (spBorderStyle)((valueBits & DiagonalBorderStyleBitMask) >> 20);
			}
			
			set
			{
				valueBits = (valueBits & DiagonalBorderStyleBitMaskInverted) | ((uint)value << 20);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spDiagonalBorderType DiagonalBorderType
		{
			get
			{
				return (spDiagonalBorderType)((valueBits & DiagonalBorderTypeBitMask) >> 24);
			}
			
			set
			{
				valueBits = (valueBits & DiagonalBorderTypeBitMaskInverted) | ((uint)value << 24);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spHorizontalAlignment HorizontalAlignment
		{
			get
			{
				return (spHorizontalAlignment)((valueBits & HorizontalAlignmentBitMask) >> 26);
			}
			
			set
			{
				valueBits = (valueBits & HorizontalAlignmentBitMaskInverted) | ((uint)value << 26);
				isCachedHashCodeValid = false;
			}
		}

		//-------from ICellStyle
		public spVerticalAlignment VerticalAlignment
		{
			get
			{
				return (spVerticalAlignment)((valueBits & VerticalAlignmentBitMask) >> 29);
			}
			
			set
			{
				valueBits = (valueBits & VerticalAlignmentBitMaskInverted) | ((uint)value << 29);
				isCachedHashCodeValid = false;
			}
		}

		public short NumberFormatIndex { get; set; }	//-------from ICellStyle (DataFormat) - this is the index into the table of DataFormats, but it's value in the table is a string

		public short IndentLevel { get; set; }	//-------from ICellStyle (Indention), must be from 0 to 15, this is the number of spaces the text is indented in the cell

		public short FontIndex { get; }	//-------from ICellStyle

		//-------from ICellStyle (called Rotation), this is the degrees, from -90 to 90
		// or xlDownward, xlUpward, xlVertical (this means text where each character isn't rotated, but they flow downward) in Excel API
		public short Orientation { get; set; }

		public const short Orientation_Downward = -90;
		public const short Orientation_Horizontal = 0;
		public const short Orientation_Upward = 90;
		public const short Orientation_Vertical = -4166;//??? will this value work?
		//xlOrientation enumeration
//xlDownward	-4170	Text runs downward. - not sure why this isn't just -90
//xlHorizontal	-4128	Text runs horizontally. - not sure why this isn't just 0
//xlUpward	-4171	Text runs upward. - not sure why this isn't just 90
//xlVertical	-4166	Text runs downward and is centered in the cell. - this can't have a value between -90 to 90

		public spFillPattern FillPattern { get; set; }	//-------from ICellStyle

#if Allow_Office2003_UserModel // the XSSF also has index, but if set, the color is also set (but not the other way around), so it's the color that determines the value, not the index for XSSF
		public short FillBackgroundColorIndex { get; set; }	//-------from ICellStyle (called FillForegroundColor)
		public short FillPatternColorIndex { get; set; }	//-------from ICellStyle (called FillBackgroundColor)

		public short LeftBorderColorIndex { get; set; }	//-------from ICellStyle (called BorderLeftColor)
		public short TopBorderColorIndex { get; set; }	//-------from ICellStyle (called BorderTopColor)
		public short RightBorderColorIndex { get; set; }	//-------from ICellStyle (called BorderRightColor)
		public short BottomBorderColorIndex { get; set; }	//-------from ICellStyle (called BorderBottomColor)
		public short DiagonalBorderColorIndex { get; set; }	//-------from ICellStyle (called BorderDiagonalColor)
#endif

		//public IColor FillBackgroundColorColor { get; }	//-------from ICellStyle, plus set for XSSF
		//public IColor FillForegroundColorColor { get; }	//-------from ICellStyle, plus set for XSSF

		public void SetFillBackgroundColor(byte red, byte green, byte blue)
		{
			FillBackgroundColor = new CellColor(red, green, blue);
		}

		public CellColor LeftBorderColor { get; set; } // ** get RGB
		public CellColor TopBorderColor { get; set; } // ** get RGB
		public CellColor RightBorderColor { get; set; } // ** get RGB
		public CellColor BottomBorderColor { get; set; } // ** get RGB
		public CellColor DiagonalBorderColor { get; set; } // ** get RGB

		private CellColor fillBackgroundColor;
		public CellColor FillBackgroundColor
		{
			get
			{
				return fillBackgroundColor;
			}
			
			set
			{
				fillBackgroundColor = value;
				isCachedHashCodeValid = false;
			}
		}

		public CellColor FillPatternColor { get; set; } // ** get/set RGB
		
		//public void SetX()
		//{
		//	CellUtil.SetCellStyleProperty()
		//}

		//public CellStyle(ICellStyle cellStyle)
		//{
		//	valueBits = 0;
		//	cachedHashCode = 0;

		//	cellStyleIndex = cellStyle.Index;
		//	cachedHashCode = GetHashCode();
		//}

		public int CreateHashCode()
		{
			int hash = 13;

			//??? instead of having these bools, we could maintain a bit string for all of the bools and that would make this faster and the equals comparison faster also
			// there are 4 bools (4 bits), could combine with horizontal alignment (3 bits), and vert. alignment (3 bits), total = 10 bits, each border style is 4 bits * 4 = 16 bits + 10 = 26 bits + 2 bits for BorderDiagonal = 28 bits, 5 bits for FillPattern would be 33 bits, which is too much for a uint

			unchecked // the code below may overflow the hash int and that will cause an exception if compiler is checking for arithmetic overflow - unchecked prevents this
			{
				hash = (hash * 7) + (int)valueBits;
				hash = (hash * 7) + NumberFormatIndex;
				hash = (hash * 7) + FontIndex;

				hash = (hash * 7) + FillBackgroundColor.GetHashCode();
				hash = (hash * 7) + FillPatternColor.GetHashCode();

				hash = (hash * 7) + BottomBorderColor.GetHashCode();
				hash = (hash * 7) + TopBorderColor.GetHashCode();
				hash = (hash * 7) + LeftBorderColor.GetHashCode();
				hash = (hash * 7) + RightBorderColor.GetHashCode();
				hash = (hash * 7) + DiagonalBorderColor.GetHashCode();

				hash = (hash * 7) + Orientation;
				hash = (hash * 7) + IndentLevel;
				hash = (hash * 7) + (int)FillPattern;

#if Allow_Office2003_UserModel
				hash = (hash * 7) + FillBackgroundColorIndex;
				hash = (hash * 7) + FillPatternColorIndex;

				hash = (hash * 7) + LeftBorderColorIndex;
				hash = (hash * 7) + TopBorderColorIndex;
				hash = (hash * 7) + RightBorderColorIndex;
				hash = (hash * 7) + BottomBorderColorIndex;
				hash = (hash * 7) + DiagonalBorderColorIndex;
#endif
			}

			return hash;
		}

		public override bool Equals(object obj)
		{
			CellStyleValue c = (CellStyleValue)obj; // this class is sealed (actually a struct???) so you don't have to worry about the as working with a derived type (which you don't want) as well as WorkbookCellStyle (which you do want)
			return Equals(c);
		}

		public override int GetHashCode()
		{
			if (!isCachedHashCodeValid)
			{
				cachedHashCode = CreateHashCode();
				isCachedHashCodeValid = true;
			}
			
			return cachedHashCode;
		}

		public bool Equals(CellStyleValue c)
		{
			bool equals = false;

			if (this.cachedHashCode != c.cachedHashCode)
			{
				// keep this false (not equals)
			}
			// try to compare items in order of things that are most likely to be different so that the statement short circuits before doing alot of comparisons
			else if (
				valueBits == c.valueBits &&
				NumberFormatIndex == c.NumberFormatIndex &&
				FontIndex == c.FontIndex &&
				FillBackgroundColor == c.FillBackgroundColor && // this uses the XSSFColor members instead of the ColorIndex members (HSSF style), so this won't work with HSSF for now??? - maybe don't support HSSF
				FillPatternColor == c.FillPatternColor &&

				BottomBorderColor == c.BottomBorderColor &&
				TopBorderColor == c.TopBorderColor &&
				LeftBorderColor == c.LeftBorderColor &&
				RightBorderColor == c.RightBorderColor &&
				DiagonalBorderColor == c.DiagonalBorderColor &&

				Orientation == c.Orientation &&
				IndentLevel == c.IndentLevel &&
				FillPattern == c.FillPattern
#if Allow_Office2003_UserModel
				&& FillBackgroundColorIndex == c.FillBackgroundColorIndex
				&& FillPatternColorIndex == c.FillPatternColorIndex

				&& LeftBorderColorIndex == c.LeftBorderColorIndex
				&& TopBorderColorIndex == c.TopBorderColorIndex
				&& RightBorderColorIndex == c.RightBorderColorIndex
				&& BottomBorderColorIndex == c.BottomBorderColorIndex
				&& DiagonalBorderColorIndex == c.DiagonalBorderColorIndex
#endif
				)
			{
				equals = true;
			}

			return equals;
		}
		
        public static bool operator ==(CellStyleValue c1, CellStyleValue c2)
        {
            return c1.Equals(c2);
        }

        public static bool operator !=(CellStyleValue c1, CellStyleValue c2)
        {
            return !(c1 == c2);
        }
	}

	[Flags]
	internal enum StyleChanges : uint // the problem with this method is it doesn't work if a value is changed back or if the default changes??? so remove this method
	{
		BorderLeft =				0x00000001,
		BorderTop =					0x00000002,
		BorderRight =				0x00000004,
		BorderBottom =				0x00000008,
		BorderDiagonal =			0x00000010,
		BorderDiagonalLineStyle =	0x00000020,

		LeftBorderColor =			0x00000040,
		TopBorderColor =			0x00000080,
		RightBorderColor =			0x00000100,
		BottomBorderColor =			0x00000200,
		BorderDiagonalColor =		0x00000400,

		FillForegroundColor =		0x00000800,
		FillBackgroundColor =		0x00001000,
		FillPattern =				0x00002000,

		Rotation =					0x00004000,

		VerticalAlignment =			0x00008000,
		Alignment =					0x00010000, // this is HorizontalAlignment, but it's just called Alignment

		WrapText =					0x00020000,
		IsLocked =					0x00040000,
		IsHidden =					0x00080000,

		DataFormat =				0x00100000,
		ShrinkToFit =				0x00200000,
		IndentLevel =				0x00400000,
	}

	// the cell styles that are created in the workbook list of cell styles
	public sealed class WorkbookCellStyle : IEquatable<WorkbookCellStyle>
	{
		short cellStyleIndexInWorkbook;

		// as style settings are changed, the code needs to re-use an existing matching style, or create a new style
		// this way API users don't have to create new styles for every single cell that needs some style even if many are the same (there are a limited # of styles 4,000 (xls) or 64,000 (xlsx) allowed in a workbook)
		// and we don't have to create styles for every single possible combination of style settings that may be possible
		// 
		private const int StyleListMaxElems = 16;
		private static List<WorkbookCellStyle> styleList = new List<WorkbookCellStyle>(StyleListMaxElems);
		private static HashSet<WorkbookCellStyle> styleSet = null; // keep this null and use the styleList until exceeding the max num elems - when that happens then start using the HashSet instead

		internal Workbook wb;
		internal ICellStyle cellStyle;

		internal WorkbookCellStyle changingWorkbookCellStyle = new WorkbookCellStyle(); // when making changes, change this, then get hashcode, etc. then
		//NPOI.XSSF.Model.StylesTable stylesSource;
		//stylesSource.
		//internal ICellStyle unattachedCellStyle = new XSSFCellStyle(stylesSource);

		private StyleChanges changesFromDefault;// ??? will this work?  What if we change back to the default?

		private static WorkbookCellStyle defaultCellStyle; //??? how do I set this?

		private uint valueBits = 0; // this is compact (using bits) form of the 4 boolean properties, and all of the enum properties except for FillPattern


		// make this internal so that we can't create WorksheetCellStyles outside of this assembly - is this really the way to do it???
		internal WorkbookCellStyle(Workbook wb, ICellStyle cellStyle)
		{
			wb.wb.GetCellStyleAt(0); // this is the default
			ICellStyle wbCellStyle = wb.wb.CreateCellStyle();
			
			this.wb = wb;
			this.cellStyle = cellStyle;
		}

		private WorkbookCellStyle()
		{
		}

		public override bool Equals(object obj)
		{
			WorkbookCellStyle c = obj as WorkbookCellStyle; // this class is sealed so you don't have to worry about the as working with a derived type (which you don't want) as well as WorkbookCellStyle (which you do want)
			return Equals(c);
		}

		public bool Equals(WorkbookCellStyle c)
		{
			bool equals = false;

			if (c != null)
			{
				if (ReferenceEquals(this, c))
				{
					equals = true;
				}
				//else if (this.changesFromDefault != c.changesFromDefault || this.cachedHashCode != c.cachedHashCode)
				//{
				//	// keep this false (not equals)
				//}
				////// try to compare items in order of things that are most likely to be different so that the statement short circuits before doing alot of comparisons
				//else if (
				//	cellStyle.DataFormat == c.cellStyle.DataFormat &&
				//	cellStyle.Alignment == c.cellStyle.Alignment &&
				//	cellStyle.FontIndex == c.cellStyle.FontIndex &&
				//	cellStyle.FillBackgroundColor == c.cellStyle.FillBackgroundColor &&
				//	cellStyle.FillForegroundColor == c.cellStyle.FillForegroundColor &&
				//	cellStyle.WrapText == c.cellStyle.WrapText &&

				//	cellStyle.BorderBottom == c.cellStyle.BorderBottom &&
				//	cellStyle.BorderTop == c.cellStyle.BorderTop &&
				//	cellStyle.BorderLeft == c.cellStyle.BorderLeft &&
				//	cellStyle.BorderRight == c.cellStyle.BorderRight &&
				//	cellStyle.BorderDiagonalLineStyle == c.cellStyle.BorderDiagonalLineStyle &&

				//	cellStyle.BottomBorderColor == c.cellStyle.BottomBorderColor &&
				//	cellStyle.TopBorderColor == c.cellStyle.TopBorderColor &&
				//	cellStyle.LeftBorderColor == c.cellStyle.LeftBorderColor &&
				//	cellStyle.RightBorderColor == c.cellStyle.RightBorderColor &&
				//	cellStyle.BorderDiagonalColor == c.cellStyle.BorderDiagonalColor &&

				//	cellStyle.VerticalAlignment == c.cellStyle.VerticalAlignment &&
				//	cellStyle.Rotation == c.cellStyle.Rotation &&
				//	cellStyle.IsLocked == c.cellStyle.IsLocked &&
				//	cellStyle.IsHidden == c.cellStyle.IsHidden &&
				//	cellStyle.ShrinkToFit == c.cellStyle.ShrinkToFit &&
				//	cellStyle.Indention == c.cellStyle.Indention &&
				//	cellStyle.FillPattern == c.cellStyle.FillPattern
				//	)
				//{
				//	equals = true;
				//}
			}

			return equals;
		}
		
        public static bool operator ==(WorkbookCellStyle c1, WorkbookCellStyle c2)
        {
			bool equals = false;

            if (c1 is null)
            {
                if (c2 is null)
                {
                    equals = true;
                }
            }
			else
			{
                if (!(c2 is null))
                {
					equals = c1.Equals(c2);
				}
			}

            return equals;
        }

        public static bool operator !=(WorkbookCellStyle c1, WorkbookCellStyle c2)
        {
            return !(c1 == c2);
        }
	}

	#if false
	public interface IColor
	{
		short Indexed { get; }
		byte[] RGB { get; }
	}

	public class XSSFColor : IColor
	{
		public XSSFColor();
		public XSSFColor(CT_Color color);
		public XSSFColor(System.Drawing.Color clr);
		public XSSFColor(byte[] rgb);

		public bool IsAuto { get; set; }
		public short Indexed { get; set; }
		public byte[] RGB { get; }
		public int Theme { get; set; }
		public double Tint { get; set; }

		public override bool Equals(object o);
		public byte[] GetARgb();
		public string GetARGBHex();
		public override int GetHashCode();
		public byte[] GetRgb();
		public byte[] GetRgbWithTint();
		public void SetRgb(byte[] rgb);
	}
	public class XSSFCellStyle : ICellStyle
	{
		public XSSFCellStyle(StylesTable stylesSource);
		public XSSFCellStyle(int cellXfId, int cellStyleXfId, StylesTable stylesSource, ThemesTable theme);

		public XSSFColor BottomBorderXSSFColor { get; } // ** get RGB
		public short DataFormat { get; set; }	//-------from ICellStyle
		public short FillBackgroundColor { get; set; }	//-------from ICellStyle
		public IColor FillBackgroundColorColor { get; set; }	//-------from ICellStyle, plus set
		public XSSFColor FillBackgroundXSSFColor { get; set; } // ** get/set RGB
		public short FillForegroundColor { get; set; }	//-------from ICellStyle
		public IColor FillForegroundColorColor { get; set; }	//-------from ICellStyle, plus set
		public XSSFColor FillForegroundXSSFColor { get; set; } // ** get/set RGB
		public FillPattern FillPattern { get; set; }	//-------from ICellStyle
		public short FontIndex { get; }	//-------from ICellStyle
		public bool IsHidden { get; set; }	//-------from ICellStyle
		public short Indention { get; set; }	//-------from ICellStyle
		public short Index { get; }	//-------from ICellStyle
		public short LeftBorderColor { get; set; }	//-------from ICellStyle
		public XSSFColor DiagonalBorderXSSFColor { get; } // ** get RGB
		public XSSFColor LeftBorderXSSFColor { get; } // ** get RGB
		public bool IsLocked { get; set; }	//-------from ICellStyle
		public short RightBorderColor { get; set; }	//-------from ICellStyle
		public XSSFColor RightBorderXSSFColor { get; } // ** get RGB
		public short Rotation { get; set; }	//-------from ICellStyle
		public short TopBorderColor { get; set; }	//-------from ICellStyle
		public XSSFColor TopBorderXSSFColor { get; } // ** get RGB
		public VerticalAlignment VerticalAlignment { get; set; }	//-------from ICellStyle
		public bool WrapText { get; set; }	//-------from ICellStyle
		public bool ShrinkToFit { get; set; }	//-------from ICellStyle
		public short BorderDiagonalColor { get; set; }	//-------from ICellStyle
		public short BottomBorderColor { get; set; }	//-------from ICellStyle
		public BorderStyle BorderDiagonalLineStyle { get; set; }	//-------from ICellStyle
		public BorderStyle BorderTop { get; set; }	//-------from ICellStyle
		public BorderStyle BorderLeft { get; set; }	//-------from ICellStyle
		public BorderStyle BorderRight { get; set; }	//-------from ICellStyle
		public BorderDiagonal BorderDiagonal { get; set; }	//-------from ICellStyle
		public HorizontalAlignment Alignment { get; set; }	//-------from ICellStyle
		public BorderStyle BorderBottom { get; set; }	//-------from ICellStyle
		protected internal int UIndex { get; } //??? not sure what this is - internal anyway


		public object Clone();
		public void CloneStyleFrom(ICellStyle source);
		public override bool Equals(object o);
		public XSSFColor GetBorderColor(BorderSide side);
		public CT_Xf GetCoreXf();
		public CT_Border GetCTBorder();
		public CT_Fill GetCTFill();
		public string GetDataFormatString();
		public IFont GetFont(IWorkbook parentWorkbook);
		public XSSFFont GetFont();
		public override int GetHashCode();
		public CT_Xf GetStyleXf();
		public void SetBorderColor(BorderSide side, XSSFColor color);
		public void SetBottomBorderColor(XSSFColor color);
		public void SetDataFormat(int fmt);
		public void SetDiagonalBorderColor(XSSFColor color);
		public void SetFillBackgroundColor(XSSFColor color);
		public void SetFillForegroundColor(XSSFColor color); //??? this is the same as set on FillForegroundXSSFColor?
		public void SetFont(IFont font);
		public void SetLeftBorderColor(XSSFColor color); // ??? why not have a set on the property instead?
		public void SetRightBorderColor(XSSFColor color); // ??? why not have a set on the property instead?
		public void SetTopBorderColor(XSSFColor color); // ??? why not have a set on the property instead?
		public void SetVerticalAlignment(short align); // ??? not sure what this does differently than VerticalAlignment property
		public void VerifyBelongsToStylesSource(StylesTable src);
	}

	public class HSSFCellStyle : ICellStyle
	{
		public HSSFCellStyle(short index, ExtendedFormatRecord rec, HSSFWorkbook workbook);
		public HSSFCellStyle(short index, ExtendedFormatRecord rec, InternalWorkbook workbook);

		public BorderStyle BorderLeft { get; set; }
		public BorderStyle BorderRight { get; set; }
		public BorderStyle BorderTop { get; set; }
		public BorderStyle BorderBottom { get; set; }
		public short LeftBorderColor { get; set; }
		public short RightBorderColor { get; set; }
		public short TopBorderColor { get; set; }
		public short BottomBorderColor { get; set; }
		public short BorderDiagonalColor { get; set; }
		public BorderStyle BorderDiagonalLineStyle { get; set; }
		public BorderDiagonal BorderDiagonal { get; set; }
		public bool ShrinkToFit { get; set; }
		public short ReadingOrder { get; set; }
		public FillPattern FillPattern { get; set; }
		public short FillBackgroundColor { get; set; }
		public IColor FillBackgroundColorColor { get; }
		public short FillForegroundColor { get; set; }
		public short Indention { get; set; }
		public short Rotation { get; set; }
		public VerticalAlignment VerticalAlignment { get; set; }
		public bool WrapText { get; set; }
		public IColor FillForegroundColorColor { get; }
		public short Index { get; }
		public HSSFCellStyle ParentStyle { get; }
		public short DataFormat { get; set; }
		public string UserStyleName { get; set; }
		public bool IsHidden { get; set; }
		public bool IsLocked { get; set; }
		public HorizontalAlignment Alignment { get; set; }
		public short FontIndex { get; }

		public void CloneStyleFrom(HSSFCellStyle source);
		public void CloneStyleFrom(ICellStyle source);
		public override bool Equals(object obj);
		public string GetDataFormatString(InternalWorkbook workbook);
		public string GetDataFormatString(IWorkbook workbook);
		public string GetDataFormatString();
		public IFont GetFont(IWorkbook parentWorkbook);
		public override int GetHashCode();
		public void SetFont(IFont font);
		public void VerifyBelongsToWorkbook(HSSFWorkbook wb);
	}

	public interface ICellStyle
	{
		x`BorderStyle BorderLeft { get; set; }1
		x`BorderDiagonal BorderDiagonal { get; set; }2
		x`BorderStyle BorderDiagonalLineStyle { get; set; }3
		`short BorderDiagonalColor { get; set; }4
		`short FillForegroundColor { get; set; }5
		`short FillBackgroundColor { get; set; }6
		`FillPattern FillPattern { get; set; }7
		`short BottomBorderColor { get; set; }8
		`short TopBorderColor { get; set; }9
		`short RightBorderColor { get; set; }10
		`short LeftBorderColor { get; set; }11
		`BorderStyle BorderBottom { get; set; }12
		`BorderStyle BorderTop { get; set; }13
		`BorderStyle BorderRight { get; set; }14
		IColor FillBackgroundColorColor { get; }
		IColor FillForegroundColorColor { get; }
		`short Rotation { get; set; }15
		`VerticalAlignment VerticalAlignment { get; set; }16
		`bool WrapText { get; set; }17
		`HorizontalAlignment Alignment { get; set; }18
		`bool IsLocked { get; set; }19
		`bool IsHidden { get; set; }20
		short FontIndex { get; }
		`short DataFormat { get; set; }21
		short Index { get; }
		bool ShrinkToFit { get; set; }22
		short Indention { get; set; }23

		void CloneStyleFrom(ICellStyle source);
		string GetDataFormatString();
		IFont GetFont(IWorkbook parentWorkbook);
		void SetFont(IFont font);
	}
	#endif
}
