package application.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.net.URI;
import java.net.URL;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.ReadableByteChannel;
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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class DownloadFileController {

	private static final String tempFolder = "temp";
	private static final String sep = File.separator;
	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");

	@RequestMapping(value = "/downloadpage")
	public String download(
			@RequestParam(name = "downloadpage", required = false, defaultValue = "default downloadpage") String name,
			Model model) {
		model.addAttribute("downloadpage", name);
		return "/content/downloadpage";
	}

	@RequestMapping(value = "/downloadExcel1", method = RequestMethod.GET)
	public ResponseEntity<Resource> downloadExcel1(HttpServletResponse resp) throws Exception {

		Path zipFile = GenerateExcel();
		ByteArrayResource resource = new ByteArrayResource(Files.readAllBytes(zipFile));
		
		HttpHeaders headers = new HttpHeaders();
		headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + zipFile.getFileName());
		headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
		headers.add("Expires", "0");

		
		//System.out.println(Files.deleteIfExists(zipFile));
		
		return ResponseEntity.ok().headers(headers).contentLength(zipFile.toFile().length())
				.contentType(MediaType.parseMediaType("application/octet-stream")).body(resource);
	}
	
	@RequestMapping(value = "/downloadExcel2", method = RequestMethod.GET)
	@ResponseBody 
	public void downloadExcel2(HttpServletResponse response) throws Exception {

		Path zipFile = GenerateExcel();

		response.setContentType("application/zip");
	    response.setHeader("Content-disposition", "attachment; filename=" + zipFile.toFile().getName());
	    OutputStream out = response.getOutputStream();
	    FileInputStream fis = new FileInputStream(zipFile.toFile());
		byte[] b = new byte[1024];
		int length;
		while((length = fis.read(b)) > 0) {
			out.write(b, 0, length);
		}
		out.flush();
		out.close();
		fis.close();
		
		//return "/content/downloadpage";
	}
	
	@RequestMapping(value = "/downloadExcel3", method = RequestMethod.GET)
	@ResponseBody 
	public void downloadExcel3(HttpServletResponse response) throws Exception {

		Path zipFile = GenerateExcel();

		response.setContentType("application/zip");
	    response.setHeader("Content-disposition", "attachment; filename=" + zipFile.toFile().getName());
	    
	    OutputStream out = response.getOutputStream();
	    out.write(Files.readAllBytes(zipFile));
	    out.flush();
	    out.close();
	    response.flushBuffer();
	    
	    //return "/content/downloadpage";
	}

	private Path GenerateExcel() throws Exception {
		
		String fileName = "excel_" + sdf.format(new Date()) + ".xlsx";
		
		
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
		String zipName = "zip_" + sdf.format(new Date()) + ".zip";
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
