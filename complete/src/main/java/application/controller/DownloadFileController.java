package application.controller;

import java.io.File;
import java.io.FileOutputStream;
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
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;

@Controller
public class DownloadFileController {

	private static final String tempFolder = "temp";
	private static final String sep = File.separator;
	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	private static String fileName = "excel_" + sdf.format(new Date()) + ".xlsx";
	private static String zipName = "zip_" + sdf.format(new Date()) + ".zip";

	@RequestMapping(value = "/downloadpage")
	public String download(
			@RequestParam(name = "downloadpage", required = false, defaultValue = "default downloadpage") String name,
			Model model) {
		model.addAttribute("downloadpage", name);
		return "/content/downloadpage";
	}

	@RequestMapping(value = "/downloadExcel", method = RequestMethod.GET)
	public String downloadExcel(HttpServletResponse resp) throws Exception {

		Path zipFile = GenerateExcel();

		resp.setContentType(null);

		// delete the zip file finally
		// Files.deleteIfExists(zipfilePath);

		return "/content/downloadpage";
	}

	private Path GenerateExcel() throws Exception {
		SXSSFWorkbook wb = new SXSSFWorkbook(1); // keep 100 rows in memory, exceeding rows will be flushed to disk
		Sheet sh = wb.createSheet();

		String[] header = { "A", "B", "C", "D", "E" };

		for (int rownum = 0; rownum < 30000; rownum++) {
			Row row = sh.createRow(rownum);
			for (int cellnum = 0; cellnum < 5; cellnum++) {

				Cell cell = row.createCell(cellnum);

				if (rownum == 0) 
					SetCellStyleAndValue(wb, row, cell, header[cellnum]);
				else 
					cell.setCellValue(sdf.format(new Date()));

			}
		}

		FileOutputStream out = new FileOutputStream(tempFolder + sep + fileName);
		wb.write(out);
		out.close();
		// dispose of temporary files backing this workbook on disk
		wb.dispose();

		Path file = Paths.get(System.getProperty("user.dir") + sep + tempFolder + sep + fileName);
		Path zipFilePath = ZipSourceFile(file);

		// return to download
		return zipFilePath;
	}

	private void SetCellStyleAndValue(SXSSFWorkbook wb, Row row, Cell cell, String value) {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		cell.setCellValue(value);
	}

	private Path ZipSourceFile(Path file) throws Exception {
		Map<String, String> env = new HashMap<>();
		// Create the zip file if it doesn't exist
		env.put("create", "true");

		URI uri = URI.create("jar:" + file.getParent().toUri() + zipName);
		Path zipFilePath;
		try (FileSystem zipfs = FileSystems.newFileSystem(uri, env)) {
			Path sourcefile = file;
			Path pathInZipfile = zipfs.getPath(file.getFileName().toString());
			// Copy a file into the zip file
			Files.copy(sourcefile, pathInZipfile, StandardCopyOption.REPLACE_EXISTING);

			Files.deleteIfExists(sourcefile);
			zipFilePath = Paths.get(zipfs.toString());
		}

		return zipFilePath;
	}
}
