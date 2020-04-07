package utility.poi;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.Optional;
import java.util.OptionalInt;
import java.util.OptionalLong;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import utility.poi.annotation.ExcelAddress;
import utility.poi.annotation.ExcelSheet;

/**
 * テンプレートファイルを利用してExcelを出力します。
 * @author Takeshi
 *
 */
public class ExcelPoiTemplateWriter<T> implements Closeable, ExcelPoi<T>
{
	private Workbook workbook;
	private boolean closeable = true;

	/**
	 * コンストラクタ
	 * @param stream 入力ストリーム
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	public ExcelPoiTemplateWriter(InputStream stream)
			throws EncryptedDocumentException, IOException
	{
		this.workbook = WorkbookFactory.create(stream);
	}

	/**
	 * コンストラクタ
	 * @param file ファイル
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	public ExcelPoiTemplateWriter(File file)
			throws EncryptedDocumentException, IOException
	{
		this.workbook = WorkbookFactory.create(file);
	}

	/**
	 * コンストラクタ
	 * @param workbook <code>Workbook</code>オブジェクト
	 */
	public ExcelPoiTemplateWriter(Workbook workbook)
	{
		this.workbook = workbook;
	}

	/**
	 * Excelに出力します。
	 * @param entity 描画対象のクラスインスタンス
	 * @param clazz 描画対象のクラス
	 */
	public void writetoExcelTemplate(T entity, Class<T> clazz)
	{
		// データが存在しない場合は処理しない
		if (entity == null) return;

		synchronized (workbook)
		{
			// ジェネリクス型に指定されている注釈を取得
			ExcelSheet sheet_info = getClassAnnotation(clazz, ExcelSheet.class);
			Map<String, ExcelAddress> column_addresses = getFieldAnnotation(clazz, ExcelAddress.class);
			Map<String, CellReference> column_references = getFieldCellReferences(workbook, clazz);
			Map<String, Method> column_getters = getFieldGetterMethod(clazz);

			// ExcelSheet注釈が付いていない場合は処理対象外
			if (sheet_info == null) return;

			// シート取得
			Sheet sheet = workbook.getSheet(sheet_info.sheet_name());

			// エンティティのフィールド単位にセット
			for (String field_name : column_getters.keySet())
			{
				try
				{
					if (column_addresses.containsKey(field_name) && column_references.containsKey(field_name))
					{
						setCellValue(workbook, sheet, entity, column_getters.get(field_name)
								, column_references.get(field_name).getRow()
								, column_references.get(field_name).getCol()
								, column_addresses.get(field_name).trim());
					}
				}
				catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e)
				{
					e.printStackTrace();
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
	 * セル参照を注釈より取得します。
	 * @param workbook <code>Workbook</code>オブジェクト
	 * @param entity_class 描画対象のクラス
	 * @return
	 */
	private Map<String, CellReference> getFieldCellReferences(Workbook workbook, Class<T> entity_class)
	{
		return Stream.of(entity_class.getDeclaredFields())
				.collect(Collectors.toMap(t -> t.getName(), t -> createCellReference(workbook, t.getAnnotation(ExcelAddress.class))));
	}

	/**
	 * セルアドレスからセル参照を取得します。<br>
	 * 名前付きセルの場合、アドレス解決を行います。
	 * @param workbook <code>Workbook</code>オブジェクト
	 * @param address <code>ExcelAddress</code>注釈
	 * @return セル参照
	 */
	private CellReference createCellReference(Workbook workbook, ExcelAddress address)
	{
		if (address == null) return null;

		Name cellname = workbook.getName(address.address());
		return new CellReference(cellname == null ? address.address() : cellname.getRefersToFormula());
	}

	/**
	 * セルに値をセットします。
	 * @param workbook <code>Workbook</code>オブジェクト
	 * @param sheet <code>Sheet</code>オブジェクト
	 * @param entity 描画対象のクラスインスタンス
	 * @param getter 項目のゲッターメソッド
	 * @param row 行番号
	 * @param col 列番号
	 * @param trim <code>true</code>の場合、前後の空白を除去します
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	private void setCellValue(Workbook workbook, Sheet sheet, T entity, Method getter, int row, int col, boolean trim)
			throws IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		Row current_row = sheet.getRow(row);
		if (current_row == null) current_row = sheet.createRow(row);
		Cell cell = current_row.getCell(col);
		if (cell == null) cell = current_row.createCell(col);
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
			setCellValue(cell, value, trim);
		}
	}
}
