package utility.poi.constant;

/**
 * 河川を指定する列挙体
 * @author Takeshi
 *
 */
public enum EnumUnderline
{
	U_NONE(0),
	U_SINGLE(1),
	U_DOUBLE(2),
	U_SINGLE_ACCOUNTING(0x21),
	U_DOUBLE_ACCOUNTING(0x22);

	private final byte underline;

	private EnumUnderline(final int underline)
	{
		this.underline = (byte)underline;
	}

	public byte getUnderline()
	{
		return this.underline;
	}
}
