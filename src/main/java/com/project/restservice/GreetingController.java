package com.project.restservice;

import java.io.ByteArrayOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.opencsv.CSVReader;

import static org.springframework.http.ResponseEntity.ok;



@RestController
@RequestMapping("/api") 
public class GreetingController {


	private static final String MIMETYPE_EXCEL = "application/vnd.ms-excel";
	
	private static final String template = "Hello, %s!";
	private final AtomicLong counter = new AtomicLong();

	@GetMapping("/greeting")
	public Greeting greeting(@RequestParam(value = "name", defaultValue = "World") String name) {
		return new Greeting(counter.incrementAndGet(), String.format(template, name));
	}
	
	@GetMapping(value = "/exportar")
	public ResponseEntity<byte[]>  exportar(@RequestParam("name")  String name) {	
		HttpHeaders headers = new HttpHeaders();
		headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
		headers.add("Pragma", "no-cache");
		headers.add("Expires", "0");
		headers.add("Content-Disposition", "attachment; filename=\"exportacion.xls\""); 
		headers.add("Content-type", MIMETYPE_EXCEL);

		byte [] excel = null;
		
		String ruta = ".//src//main//resources//files//";
		String rutaCompleta = ruta + name;
		List<List<String>> records = new ArrayList<List<String>>();

		
		try (CSVReader csvReader = new CSVReader(new FileReader(rutaCompleta));) {
		    String[] values = null;
		    while ((values = csvReader.readNext()) != null) {
		        records.add(Arrays.asList(values));
		    }
		}catch(Exception e) {
			System.out.println(e);
		}

		excel = crearExcelExportacionListado(records);

		return ok().headers(headers).body(excel);
		
		
	
	}

	
	@PostMapping(value="/upload") 
	public String setAdjuntar(
			@RequestParam("files") MultipartFile[] adjuntos) throws IOException {
		//String respuesta = null;
		//if (adjuntos.length > 0) {
		
		String nombreArchivo = null;
		String  upload_folder = ".//src//main//resources//files//";
		try {
			for(MultipartFile file: adjuntos) {
				if(!file.isEmpty()) {
					byte[] bytes = file.getBytes();
					
					nombreArchivo = file.getOriginalFilename();
					Path path = Paths.get(upload_folder + file.getOriginalFilename());
					Files.write(path, bytes);

				}//if
			}//for
			
		}catch(Exception e) {
			 System.out.println("ERROR	" + e);
		}
		 return nombreArchivo;
		
		//return convertServices.setAdjuntar(adjuntos);
						
		//}
		//return null;
		
	}
	
	
	private void crearCabeceraTabla(HSSFSheet sheet, List<String> cabecera) {
		int rowNum = 0;
		int cellNum=-1;

		Row row = sheet.createRow(rowNum);
		HSSFFont font= sheet.getWorkbook().createFont();
		HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		font.setBold(true); 
		cellStyle.setFont(font);

		Cell cell;
		for(int i = 0; i < cabecera.size(); i++) {
			cell = row.createCell(++cellNum);	
			cell.setCellValue(cabecera.get(i));
			cell.setCellStyle(cellStyle);
		}
	}
	

	
	private byte[] crearExcelExportacionListado(List<List<String>> listados) {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();

		try(HSSFWorkbook workbook = new HSSFWorkbook();) {
			HSSFSheet sheet = workbook.createSheet("Listado de carnets");
			ArrayList<String> cabecera = new ArrayList<>();
			cabecera.add("SKU (Ref)");
			cabecera.add("SKU-Codigo almacen");
			cabecera.add("EAN");
			cabecera.add("Codigo almacen");
			cabecera.add("Nombre producto");
			cabecera.add("Tienda, Categoria, Seccion, Familia y Subfamilia");
			cabecera.add("Descripción");
			cabecera.add("Opciones especiales");
			cabecera.add("Observaciones");
			cabecera.add("Precio normal (€) IVA incluido");
			cabecera.add("Precio oferta (€) IVA incluido");
			cabecera.add("Tipo IVA");
			cabecera.add("Peso (kg.)");
			cabecera.add("Stock (unidades)");
			cabecera.add("Compra mínima (unidades)");
			cabecera.add("Compra máxima (unidades)");
			cabecera.add("Etiqueta 1");
			cabecera.add("Etiqueta 2");
			cabecera.add("Etiqueta 3");
			cabecera.add("Etiqueta 4");
			cabecera.add("Etiqueta 5");
			cabecera.add("Enlace foto principal");
			crearCabeceraTabla(sheet, cabecera);
			crearCuerpoTablaExportacion(sheet, listados);
			workbook.write(baos);
			baos.close();
			return baos.toByteArray();
		} catch (java.io.IOException e) {
			
			//log.error(Translator.toLocale("ERROR_EXCEL"), e);
		}
		return "".getBytes();
	}
	
	private void crearCuerpoTablaExportacion(HSSFSheet sheet, List<List<String>> listado) {
		int cellNum;
		Row row;
		Cell cell;
		int rowNum = 0;
		for (int i = 0; i < listado.size(); i ++) {
			List<String> b = listado.get(i);
			String iva = "";
			if(b.get(12).equals("10%") || b.get(12).equals("10")) {
				iva = "IVA Reducido";
			} else if ((b.get(12).equals("4%") || b.get(12).equals("4"))) {
				iva ="IVA Superreducido";	
			}else {
				iva ="Estándar";
			}
			
			cellNum=-1;
			row = sheet.createRow(++rowNum);
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(6));
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(6)+"-"+b.get(9));
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(10));
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(9));
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(7));
			cell = row.createCell(++cellNum);
			cell.setCellValue("TIENDA>"+b.get(1)+">"+b.get(3)+">"+b.get(5));
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(7));
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(11));
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue(iva);
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue("");
			cell = row.createCell(++cellNum);
			cell.setCellValue(b.get(13));
			
    	} 
	}
	
}