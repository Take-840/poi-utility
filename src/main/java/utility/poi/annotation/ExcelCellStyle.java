package utility.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import utility.poi.constant.EnumFontType;
import utility.poi.constant.EnumUnderline;

/**
 * Excel出力時のスタイルを指定する注釈
 * @author Takeshi
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCellStyle
{
	/** 背景色 */
	IndexedColors backgroundcolor() default IndexedColors.AUTOMATIC;

	/** 文字の横位置 */
	HorizontalAlignment horizontal_alignment() default HorizontalAlignment.GENERAL;

	/** 文字の縦位置 */
	VerticalAlignment vertical_alignment() default VerticalAlignment.TOP;

	/** 表示形式（Excelの書式指定） */
	String format() default "";

	/** 改行の有無 */
	boolean wraptext() default false;

	/** 枠線の色 */
	IndexedColors bordercolor() default IndexedColors.BLACK;

	/** 上部分の枠線のスタイル */
	BorderStyle top() default BorderStyle.NONE;

	/** 下部分の枠線のスタイル */
	BorderStyle bottom() default BorderStyle.NONE;

	/** 左部分の枠線のスタイル */
	BorderStyle left() default BorderStyle.NONE;

	/** 右部分の枠線のスタイル */
	BorderStyle right() default BorderStyle.NONE;

	/** フォント */
	EnumFontType font() default EnumFontType.Meiryo;

	/** フォントサイズ */
	int size() default 11;

	/** フォントの色 */
	IndexedColors forecolor() default IndexedColors.AUTOMATIC;

	/** 太字 */
	boolean bold() default false;

	/** イタリック */
	boolean italic() default false;

	/** 取消線 */
	boolean strikeout() default false;

	/** 下線 */
	EnumUnderline underline() default EnumUnderline.U_NONE;
}
