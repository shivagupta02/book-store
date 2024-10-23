package com.bookStore.helper;

import org.springframework.web.multipart.MultipartFile;

public class Helper {

	public static boolean checkExcelFormat(MultipartFile file)
	{
		String contentType = file.getContentType();
	
		if(contentType.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
			return true;
		}
		else
		{
			return false;
		}
	}
}
