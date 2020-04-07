package utility.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelSheet
{
	/** シート名 */
	String sheet_name() default "sheet1";

	/** タイトル行の描画有無 */
	boolean draw_title() default true;

	/** タイトルの行の開始位置 */
	int row_start() default 0;

	/** タイトルの列の開始位置 */
	int column_start() default 0;

	/** フィルターの有無 */
	boolean auto_filter() default false;

	/** タイトル行のウィンドウ枠の固定の有無 */
	boolean freeze_pane() default true;

	/** タイトル行のスタイル */
	ExcelCellStyle style();
}
