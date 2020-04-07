package utility.poi;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.OptionalInt;
import java.util.OptionalLong;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import utility.poi.annotation.ExcelCellStyle;
import utility.poi.annotation.ExcelColumn;
import utility.poi.annotation.ExcelSheet;

/**
 * Apache POIを利用してExcelを出力するユーティリティクラス<br>
 * リスト形式のデータをExcelに出力する場合に利用します。
 * @author Takeshi
 *
 */
public class ExcelPoiGenerator<T> implements Closeable, ExcelPoi<T>
{
	private Workbook workbook;
	private boolean closeable = true;

	/**
	 * コンストラクタ
	 * <code>XSSF</code>形式のExcelを作成します。
	 * @throws IOException
	 */
	public ExcelPoiGenerator()
			throws IOException
	{
		this(true);
	}

	/**
	 * コンストラクタ
	 * @param xssf <code>XSSF</code>形式のExcelを作成する場合<code>true</code>
	 * @throws IOException
	 */
	public ExcelPoiGenerator(boolean xssf)
			throws IOException
	{
		this.workbook = WorkbookFactory.create(xssf);
	}

	/**
	 * コンストラクタ
	 * @param workbook <code>Workbook</code>オブジェクト
	 */
	public ExcelPoiGenerator(Workbook workbook)
	{
		this.workbook = workbook;
	}

	/**
	 * 配列データをExcelに出力します。
	 * @param data 配列データ
	 * @param clazz 描画対象のクラス
	 */
	public void writetoExcel(List<T> data, Class<T> clazz)
	{
		// データが存在しない場合は処理しない
		if (data == null || data.size() == 0) return;

		synchronized (workbook)
		{
			// ジェネリクス型に指定されている注釈を取得
			ExcelSheet sheet_info = getClassAnnotation(clazz, ExcelSheet.class);
			Map<String, CellStyle> column_styles = getFieldCellStyle(workbook, clazz);
			Map<String, ExcelColumn> column_infos = getFieldAnnotation(clazz, ExcelColumn.class);
			Map<String, Method> column_getters = getFieldGetterMethod(clazz);

			// ExcelSheet注釈が付いていない場合は処理対象外
			if (sheet_info == null) return;

			// シート作成
			Sheet sheet = workbook.createSheet(sheet_info.sheet_name());

			int current_row = sheet_info.row_start();
			int current_col = sheet_info.column_start();
			int column_num  = column_infos.size();

			// タイトル描画
			if (sheet_info.draw_title())
			{
				// セルスタイル、テキストの設定
				CellStyle title_style = createCellStyle(workbook, sheet_info.style());
				Row title_row = sheet.createRow(current_row++);
				for (String field_name : column_infos.keySet())
				{
					int width = column_infos.get(field_name).width();
					if (width != -1) sheet.setColumnWidth(current_col, width * 256);
					setCellTitle(sheet, column_infos.get(field_name).name(), title_row, current_col++, title_style);
				}

				// フィルターの設定
				if (sheet_info.auto_filter())
				{
					sheet.setAutoFilter(new CellRangeAddress(sheet_info.row_start(), sheet_info.row_start(), sheet_info.column_start(), sheet_info.column_start() + column_num - 1));
				}

				// 固定行の設定
				if (sheet_info.freeze_pane())
				{
					sheet.createFreezePane(sheet_info.column_start(), sheet_info.row_start() + 1);
				}
			}

			// 値の描画
			for (T entity : data)
			{
				// 行の作成
				Row row = sheet.createRow(current_row++);
				current_col = sheet_info.column_start();

				// フィールド単位に出力
				for (String field_name : column_infos.keySet())
				{
					try
					{
						setCell(sheet, entity, row, current_col++, column_getters.get(field_name), column_styles.get(field_name), column_infos.get(field_name));
					}
					catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e)
					{
						e.printStackTrace();
					}
				}
			}

			workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
		}
	}

	/**
	 * <code>Workbook</code>オブジェクトを取得します。
	 * @return <code>Workbook</code>オブジェクト
	 */
	public Workbook getWorkbook()
	{
		this.closeable = false;
		return this.workbook;
	}

	/**
	 * ストリームに出力します。
	 * @param stream 出力ストリーム
	 * @throws IOException
	 */
	public void write(OutputStream stream)
			throws IOException
	{
		workbook.write(stream);
	}

	/**
	 * 終了処理。<code>Workbook</code>オブジェクトを閉じます。
	 */
	@Override
	public void close()
			throws IOException
	{
		// Springで利用する場合、閉じてしまうと出力できないため
		if (this.closeable) this.workbook.close();
	}

	/**
	 * タイトルを設定します。
	 * @param sheet <code>Sheet</code>オブジェクト
	 * @param title タイトル
	 * @param row <code>Row</code>オブジェクト
	 * @param col 列番号
	 * @param style セルスタイル
	 */
	private void setCellTitle(Sheet sheet, String title, Row row, int col, CellStyle style)
	{
		// セルの作成
		Cell cell = row.createCell(col, CellType.STRING);
		if (style != null) cell.setCellStyle(style);
		cell.setCellValue(title);
	}

	/**
	 * フィールドに設定されているセルスタイルを注釈より取得します。
	 * @param workbook <code>Workbook</code>オブジェクト
	 * @param entity_class 描画対象のクラス
	 * @return
	 */
	private Map<String, CellStyle> getFieldCellStyle(Workbook workbook, Class<T> entity_class)
	{
		return Stream.of(entity_class.getDeclaredFields())
				.collect(Collectors.toMap(t -> t.getName(), t -> createCellStyle(workbook, t.getAnnotation(ExcelCellStyle.class))));
	}

	/**
	 * セルスタイルを注釈より作成します。
	 * @param workbook <code>Workbook</code>オブジェクト
	 * @param annotation_style セルスタイル注釈
	 * @return <code>CellStyle</code>オブジェクト
	 */
	private CellStyle createCellStyle(Workbook workbook, ExcelCellStyle annotation_style)
	{
		CellStyle style = workbook.createCellStyle();
		if (annotation_style != null)
		{
			if (annotation_style.backgroundcolor() != IndexedColors.AUTOMATIC)
			{
				style.setFillForegroundColor(annotation_style.backgroundcolor().getIndex());							// 背景色
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
			style.setAlignment(annotation_style.horizontal_alignment());												// 横位置
			style.setVerticalAlignment(annotation_style.vertical_alignment());											// 縦位置
			if (annotation_style.format() != null && !annotation_style.format().trim().equals(""))
			{
				DataFormat format = workbook.createDataFormat();														// 表示書式
				style.setDataFormat(format.getFormat(annotation_style.format()));
			}
			style.setWrapText(annotation_style.wraptext());																// 改行の有無
			if (annotation_style.top() != BorderStyle.NONE)
			{
				style.setTopBorderColor(annotation_style.bordercolor().getIndex());										// 上部分の枠線
				style.setBorderTop(annotation_style.top());
			}
			if (annotation_style.bottom() != BorderStyle.NONE)
			{
				style.setBottomBorderColor(annotation_style.bordercolor().getIndex());									// 下部分の枠線
				style.setBorderBottom(annotation_style.bottom());
			}
			if (annotation_style.left() != BorderStyle.NONE)
			{
				style.setLeftBorderColor(annotation_style.bordercolor().getIndex());									// 左部分の枠線
				style.setBorderLeft(annotation_style.left());
			}
			if (annotation_style.right() != BorderStyle.NONE)
			{
				style.setRightBorderColor(annotation_style.bordercolor().getIndex());									// 右部分の枠線
				style.setBorderRight(annotation_style.right());
			}
			Font font = workbook.createFont();
			font.setFontName(annotation_style.font().getFontname());													// フォント名
			font.setFontHeightInPoints((short)annotation_style.size());													// フォントサイズ
			font.setColor(annotation_style.forecolor().getIndex());														// フォントの色
			font.setBold(annotation_style.bold());																		// 太字
			font.setItalic(annotation_style.italic());																	// 斜体
			font.setStrikeout(annotation_style.strikeout());															// 取消線
			font.setUnderline(annotation_style.underline().getUnderline());												// 下線
			style.setFont(font);
		}

		return style;
	}

	/**
	 * セルを設定します。
	 * @param sheet <code>Sheet</code>オブジェクト
	 * @param entity 描画対象のクラスインスタンス
	 * @param row <code>Row</code>オブジェクト
	 * @param col 列番号
	 * @param getter 項目値を取得するゲッターメソッド
	 * @param style <code>CellStyle</code>オブジェクト
	 * @param column_info 項目描画の注釈
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	private void setCell(Sheet sheet, T entity, Row row, int col, Method getter, CellStyle style, ExcelColumn column_info)
			throws IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		// セルの作成
		Cell cell = (column_info.type() == CellType._NONE) ? row.createCell(col) : row.createCell(col, column_info.type());
		if (style != null) cell.setCellStyle(style);

		Object value = getter.invoke(entity, (Object[])null);
		if (value == null)
		{
			return;
		}
		else if (value.getClass() == Optional.class)
		{
			setCellOptionalValue(cell, (Optional<?>)value);
		}
		else if (value.getClass() == OptionalInt.class)
		{
			setCellOptionalIntValue(cell, (OptionalInt)value);
		}
		else if (value.getClass() == OptionalLong.class)
		{
			setCellOptionalLongValue(cell, (OptionalLong)value);
		}
		else
		{
			setCellValue(cell, value, column_info.trim());
		}
	}
}
