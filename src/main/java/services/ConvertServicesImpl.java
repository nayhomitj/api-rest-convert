package services;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;

import java.io.FileReader;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;

import javax.servlet.http.HttpServletRequest;


import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.opencsv.CSVReader;

@Service
public class ConvertServicesImpl  implements ConvertServices {

	@Autowired
	public HttpServletRequest request;

	
	/**
	 * crea la cabecera
	 */
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
	
	@Override
	public byte[] exportar(String name) {

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

		return crearExcelExportacionListado(records);

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
	
	public String  setAdjuntar(MultipartFile[] multipartFile) {
		String nombreArchivo = null;
		String  upload_folder = ".//src//main//resources//files//";
		try {
			for(MultipartFile file: multipartFile) {
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
	}
	
}
