package utility.poi.constant;

/**
 * フォントを指定する列挙体
 * @author Takeshi
 *
 */
public enum EnumFontType
{
	Arial("Arial"),
	Arial_Narrow("Arial Narrow"),
	Arial_Unicode_MS("Arial Unicode MS"),
	Calibri("Calibri"),
	Century("Century"),
	Meiryo("Meiryo UI"),
	Microsoft_Sans_Serif("Microsoft Sans Serif"),
	MS_Gothic("MS Gothic"),
	MS_PGothic("MS PGothic"),
	MS_UIGothic("MS UI Gothic"),
	MS_Mincho("MS Mincho"),
	MS_PMincho("MS PMincho"),
	Segoe_UI("Segoe UI"),
	Tahoma("Tahoma"),
	TimesNewRoman("Times New Roman");

	private final String text;

	private EnumFontType(final String text)
	{
		this.text = text;
	}

	public String getFontname()
	{
		return this.text;
	}
}
