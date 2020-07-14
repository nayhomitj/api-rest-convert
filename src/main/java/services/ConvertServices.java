package services;

import org.springframework.web.multipart.MultipartFile;

public interface ConvertServices {
	
	public byte[] exportar(String ruta);
	
	public String setAdjuntar(MultipartFile[] multipartFile);


}
