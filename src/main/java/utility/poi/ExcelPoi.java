package utility.poi;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;
import java.util.OptionalInt;
import java.util.OptionalLong;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author Takeshi
 *
 */
public interface ExcelPoi<T>
{
	/**
	 * クラスに指定されている注釈を取得します。
	 * @param <S>
	 * @param entity_class 取得対象のクラス
	 * @param annotation_class 注釈クラス
	 * @return
	 */
	default <S extends Annotation> S getClassAnnotation(Class<T> entity_class, Class<S> annotation_class)
	{
		return entity_class.getAnnotation(annotation_class);
	}

	/**
	 * フィールドに指定されている注釈を取得します。
	 * @param <S>
	 * @param entity_class 取得対象のクラス
	 * @param annotation_class 注釈クラス
	 * @return フィールド名をキー、注釈を値とした<code>Map</code>インターフェース
	 */
	default <S extends Annotation> Map<String, S> getFieldAnnotation(Class<T> entity_class, Class<S> annotation_class)
	{
		return Stream.of(entity_class.getDeclaredFields())
			.filter(t -> t.getAnnotation(annotation_class) != null)
			.collect(Collectors.toMap(t -> t.getName(), t -> t.getAnnotation(annotation_class), (a1, a2) -> a1, LinkedHashMap::new));
	}

	/**
	 * フィールドのGetterメソッドを取得します。
	 * @param entity_class 取得対象のクラス
	 * @return フィールド名をキー、メソッドを値とした<code>Map</code>インターフェース
	 */
	default Map<String, Method> getFieldGetterMethod(Class<T> entity_class)
	{
		Map<String, Method> methods = new LinkedHashMap<>();

		for (Field field : entity_class.getDeclaredFields())
		{
			// Getterメソッドの取得
			PropertyDescriptor property = null;
			try
			{
				property = new PropertyDescriptor(field.getName(), entity_class);
			}
			catch (IntrospectionException exp) { }
			if (property == null) continue;

			Method getter = property.getReadMethod();
			if (getter != null) methods.put(field.getName(), getter);
		}

		return methods;
	}

	/**
	 * <code>Optional</code>型の値をセルにセットします。
	 * @param cell セル
	 * @param value 値
	 */
	default void setCellOptionalValue(Cell cell, Optional<?> value)
	{
		// 値がセットされていない場合は何もしない
		if (value.isEmpty()) return;

		setCellValue(cell, value.get());
	}

	/**
	 * <code>OptionalInt</code>型の値をセルにセットします。
	 * @param cell セル
	 * @param value 値
	 */
	default void setCellOptionalIntValue(Cell cell, OptionalInt value)
	{
		// 値がセットされていない場合は何もしない
		if (value.isEmpty()) return;

		setCellValue(cell, value.getAsInt());
	}

	/**
	 * <code>OptionalLong</code>型の値をセルにセットします。
	 * @param cell セル
	 * @param value 値
	 */
	default void setCellOptionalLongValue(Cell cell, OptionalLong value)
	{
		// 値がセットされていない場合は何もしない
		if (value.isEmpty()) return;

		setCellValue(cell, value.getAsLong());
	}

	/**
	 * 値をセルにセットします。
	 * @param cell セル
	 * @param value 値
	 */
	default void setCellValue(Cell cell, Object value)
	{
		setCellValue(cell, value, false);
	}

	/**
	 * 値をセルにセットします。
	 * @param cell セル
	 * @param value 値
	 * @param trim 文字列の値で前後の空白を取り除く場合<code>true</code>
	 */
	default void setCellValue(Cell cell, Object value, boolean trim)
	{
		if (value == null) return;

		// 文字列型の場合
		else if (value.getClass() == String.class)
		{
			cell.setCellValue(trim ? ((String)value).trim() : (String)value);
		}

		// 数値型の場合
		else if (value.getClass() == BigDecimal.class)
		{
			cell.setCellValue(((BigDecimal)value).doubleValue());
		}
		else if (value.getClass() == BigInteger.class)
		{
			cell.setCellValue(((BigInteger)value).longValue());
		}
		else if (value.getClass() == Integer.class || value.getClass() == int.class)
		{
			cell.setCellValue((int)value);
		}
		else if (value.getClass() == Long.class || value.getClass() == long.class)
		{
			cell.setCellValue((long)value);
		}
		else if (value.getClass() == Float.class || value.getClass() == float.class)
		{
			cell.setCellValue((float)value);
		}
		else if (value.getClass() == Double.class || value.getClass() == double.class)
		{
			cell.setCellValue((double)value);
		}

		// boolean型の場合
		else if (value.getClass() == Boolean.class || value.getClass() == boolean.class)
		{
			cell.setCellValue((boolean)value);
		}

		// 日付型の場合
		else if (value.getClass() == Calendar.class)
		{
			cell.setCellValue((Calendar)value);
		}
		else if (value.getClass() == Date.class)
		{
			cell.setCellValue((Date)value);
		}
		else if (value.getClass() == LocalDate.class)
		{
			cell.setCellValue((LocalDate)value);
		}
		else if (value.getClass() == LocalDateTime.class)
		{
			cell.setCellValue((LocalDateTime)value);
		}
	}
}
