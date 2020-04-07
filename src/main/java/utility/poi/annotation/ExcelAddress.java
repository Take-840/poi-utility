package utility.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * データを出力するアドレス（A1形式か名前を指定）を指定する注釈
 * @author Takeshi
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAddress
{
	/** データを出力するアドレスを指定（A1形式か名前を指定） */
	String address();

	/**
	 * 値の出力時に前後の空白を除去するか否かを指定します。<br>
	 * 文字列型の値を持つ場合のみ有効です。
	 */
	boolean trim() default true;
}
