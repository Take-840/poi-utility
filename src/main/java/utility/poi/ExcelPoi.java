package utility.poi;

import java.lang.annotation.Annotation;

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
}
