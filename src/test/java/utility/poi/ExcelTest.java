package utility.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;
import java.util.OptionalInt;

import org.junit.Test;

public class ExcelTest
{
	@Test
	public void generateExcel()
	{
		Model[] array_data = {
				new Model() {{ setCode("1");  setName("Name_1\nGGG"); setQuantity(OptionalInt.of(1000)); setAmount(new BigDecimal(199));    setModified(LocalDate.now()); }},
				new Model() {{ setCode("2");  setName("Name_2");      setQuantity(OptionalInt.of(1100)); setAmount(new BigDecimal(20100));  setModified(LocalDate.now()); }},
				new Model() {{ setCode("3");  setName("Name_3");      setQuantity(OptionalInt.of(1020)); setAmount(new BigDecimal(1512));   setModified(LocalDate.now()); }},
				new Model() {{ setCode("4");  setName("Name_4");      setQuantity(OptionalInt.of(1003)); setAmount(new BigDecimal(54651));  setModified(LocalDate.now()); }},
				new Model() {{ setCode("5");  setName("Name_5");      setQuantity(OptionalInt.of(4000)); setAmount(new BigDecimal(661515)); setModified(LocalDate.now()); }},
				new Model() {{ setCode("6");  setName("Name_6");      setQuantity(OptionalInt.of(1500)); setAmount(new BigDecimal(5432));   setModified(LocalDate.now()); }},
				new Model() {{ setCode("7");  setName("Name_7");      setQuantity(OptionalInt.of(1060)); setAmount(new BigDecimal(9136));   setModified(LocalDate.now()); }},
				new Model() {{ setCode("8");  setName("Name_8");      setQuantity(OptionalInt.of(1007)); setAmount(new BigDecimal(51135));  setModified(LocalDate.now()); }},
				new Model() {{ setCode("9");  setName("Name_9");      setQuantity(OptionalInt.of(8000)); setAmount(new BigDecimal(216));    setModified(LocalDate.now()); }},
				new Model() {{ setCode("10"); setName(null);          setQuantity(OptionalInt.empty());  setAmount(null);                   setModified(null); }},
		};
		List<Model> data = Arrays.asList(array_data);

		try (
				OutputStream stream = new FileOutputStream(String.format("%s\\generate_test.xlsx", getResourceFolder()));
				ExcelPoiGenerator<Model> writer = new ExcelPoiGenerator<>();
			)
		{
			writer.writetoExcel(data, Model.class);
			writer.write(stream);
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	@Test
	public void templateExcel()
	{
		TemplateModel data = new TemplateModel()
		{{
			setIssue_date(LocalDate.now());
			setName("日本 太郎");
			setAmount(new BigDecimal(14000));
		}};

		try (
				InputStream stream = ExcelTest.class.getClassLoader().getResourceAsStream("template.xlsx");
				ExcelPoiTemplateWriter<TemplateModel> writer = new ExcelPoiTemplateWriter<>(stream);
				OutputStream outstream = new FileOutputStream(String.format("%s\\template_test.xlsx", getResourceFolder()));
			)
		{
			writer.writetoExcelTemplate(data, TemplateModel.class);
			writer.write(outstream);
			outstream.flush();
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	private String getResourceFolder()
	{
		File file = new File(ExcelTest.class.getClassLoader().getResource("template.xlsx").getPath());
		return file.getParent();
	}
}
