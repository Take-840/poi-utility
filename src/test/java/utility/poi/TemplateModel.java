package utility.poi;

import java.math.BigDecimal;
import java.time.LocalDate;

import lombok.Data;
import utility.poi.annotation.ExcelAddress;
import utility.poi.annotation.ExcelCellStyle;
import utility.poi.annotation.ExcelSheet;

@ExcelSheet(sheet_name = "Sheet1", style = @ExcelCellStyle())
@Data
public class TemplateModel
{
	@ExcelAddress(address = "ISSUE_DATE")
	LocalDate issue_date;

	@ExcelAddress(address = "NAME")
	String name;

	@ExcelAddress(address = "AMOUNT")
	BigDecimal amount;
}
