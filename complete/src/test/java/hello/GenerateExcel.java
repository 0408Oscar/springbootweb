package hello;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class GenerateExcel {
	
	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	private static final String tempFolder = "temp";
	private static final String sep = File.separator;
	private static String fileName = "excel_" + sdf.format(new Date()) + ".xlsx";
	private static String zipName = "zip_" + sdf.format(new Date()) + ".zip";

	public static void main(String[] args) throws Exception, IOException {

		System.out.println("a: " + System.getProperty("user.dir"));
		System.out.println("b: " + System.getProperty("user.home"));
		System.out.println("c: " + Paths.get("").toAbsolutePath());
		System.out.println();

		Date start = new Date();
		System.out.println("Start at : " + start);

		SXSSFWorkbook wb = new SXSSFWorkbook(1); // keep 100 rows in memory, exceeding rows will be flushed to disk
		SXSSFSheet sh = wb.createSheet();
		
		String[] header = { "A", "B", "C", "D", "E" };

		for (int rownum = 0; rownum < 30000; rownum++) {
			Row row = sh.createRow(rownum);
			for (int cellnum = 0; cellnum < 5; cellnum++) {
				
				Cell cell = row.createCell(cellnum);
				
				if (rownum == 0) 
					SetCellStyleAndValue(wb, row, cell, header[cellnum]);
				else
					cell.setCellValue(sdf.format(new Date()));

				cell.getColumnIndex();
				
			}
			
		}
		
		FileOutputStream out = new FileOutputStream(tempFolder + sep + fileName);
		wb.write(out);
		out.close();
		// dispose of temporary files backing this workbook on disk
		wb.dispose();
		Date end = new Date();
		System.out.println("End at " + end);

		System.out.println("Takes  " + (end.getTime() - start.getTime()));

		Path file = Paths.get(System.getProperty("user.dir") + sep + tempFolder + sep + fileName);

		System.out.println("1: " + (file.getFileName()));
		System.out.println("2: " + (file.getFileSystem()));
		System.out.println("3: " + (file.getName(1)));
		System.out.println("4: " + (file.getParent()));
		System.out.println("5: " + (file.getName(2)));
		System.out.println("6: " + (file.getRoot()));

		System.out.println("7: " + (file.toString()));
		System.out.println("8: " + (file.toAbsolutePath()));
		System.out.println("9: " + (file.toRealPath()));
		System.out.println("10: " + (file.toUri()));
		System.out.println("11: " + (file.spliterator()));
		System.out.println("12: " + (file.getRoot()));

		Map<String, String> env = new HashMap<>();
		// Create the zip file if it doesn't exist
		env.put("create", "true");

		URI uri = URI.create("jar:" + file.getParent().toUri() + zipName);

		try (FileSystem zipfs = FileSystems.newFileSystem(uri, env)) {
			Path sourcefile = file;
			Path pathInZipfile = zipfs.getPath(file.getFileName().toString());
			// Copy a file into the zip file
			Path copy = Files.copy(sourcefile, pathInZipfile, StandardCopyOption.REPLACE_EXISTING);
			
			System.out.println("13: " + (copy.toString()));
			System.out.println("13a: " + (copy.toAbsolutePath()));
			System.out.println("14: " + (sourcefile.toString()));
			System.out.println("14a: " + (sourcefile.toAbsolutePath()));
			System.out.println("15: " + (pathInZipfile.toString()));
			System.out.println("16: " + (zipfs.toString()));
			
			//Files.deleteIfExists(sourcefile);
		}

	}

	private static void SetCellStyleAndValue(SXSSFWorkbook wb, Row row, Cell cell, String value) {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		cell.setCellValue(value==null? "" : value);
	}

}
