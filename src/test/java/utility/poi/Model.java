package utility.poi;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.OptionalInt;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import lombok.Data;
import utility.poi.annotation.ExcelCellStyle;
import utility.poi.annotation.ExcelColumn;
import utility.poi.annotation.ExcelSheet;

@ExcelSheet(sheet_name = "テスト", auto_filter = true, style = @ExcelCellStyle(bold = true, backgroundcolor = IndexedColors.GREY_25_PERCENT))
@Data
public class Model
{
	@ExcelColumn(name = "Code", type = CellType.STRING, width = 20)
	@ExcelCellStyle(top = BorderStyle.THIN, bottom = BorderStyle.THIN, left = BorderStyle.THIN, right = BorderStyle.THIN)
	String code;

	@ExcelColumn(name = "Name", width = 50)
	@ExcelCellStyle(wraptext = true, top = BorderStyle.THIN, bottom = BorderStyle.THIN, left = BorderStyle.THIN, right = BorderStyle.THIN)
	String name;

	@ExcelColumn(name = "Q'ty", width = 26)
	@ExcelCellStyle(format = "#,##0_ ", top = BorderStyle.THIN, bottom = BorderStyle.THIN, left = BorderStyle.THIN, right = BorderStyle.THIN)
	OptionalInt quantity;

	@ExcelColumn(name = "Amount", width = 26)
	@ExcelCellStyle(format = "#,##0_ ", top = BorderStyle.THIN, bottom = BorderStyle.THIN, left = BorderStyle.THIN, right = BorderStyle.THIN)
	BigDecimal amount;

	@ExcelColumn(name = "Modify", width = 20)
	@ExcelCellStyle(format = "yyyy-mm-dd", horizontal_alignment = HorizontalAlignment.CENTER, top = BorderStyle.THIN, bottom = BorderStyle.THIN, left = BorderStyle.THIN, right = BorderStyle.THIN)
	LocalDate modified;
}
