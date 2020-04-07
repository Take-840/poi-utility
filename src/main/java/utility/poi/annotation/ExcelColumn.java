package utility.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Excel出力時の項目を指定する注釈
 * @author Takeshi
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn
{
	/** 項目のタイトル */
	String name();

	/** 項目のデータ型 */
	CellType type() default CellType.BLANK;

	/**
	 * 項目の幅
	 * 表示する文字数で指定します。
	 */
	int width() default -1;

	/**
	 * 計算式
	 * 先頭の<code>=</code>を除いて指定します。
	 * 計算式が設定された場合、<code>Entity</code>の値は無視されます。
	 */
	String formula() default"";

	/**
	 * 値の出力時に前後の空白を除去するか否かを指定します。<br>
	 * 文字列型の値を持つ場合のみ有効です。
	 */
	boolean trim() default true;
}
