package com.bookStore.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.bookStore.entity.Book;
import com.bookStore.repository.BookRepository;
import com.fasterxml.jackson.databind.exc.InvalidFormatException;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;

@Service
public class BookService {

	private static final String XSSFColor = null;
	@Autowired
	private BookRepository bRepo;
	
	public void save(Book b)
	{
	bRepo.save(b);
	}
	
	public List<Book> getAllBook()
	{
		return bRepo.findAll();
	}
	public Book getBookById(int id)
	{
		return bRepo.findById(id).get();
	}
	public void deleteById(int id)
	{
		bRepo.deleteById(id);
	}
	public void generateExcel(HttpServletResponse response) throws IOException
	{
		List<Book> books=bRepo.findAll();
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet =workbook.createSheet("BookStock");
		HSSFRow row = sheet.createRow(0);
		
		
		row.createCell(0).setCellValue("ID");
		row.createCell(1).setCellValue("Book_Name");
		row.createCell(2).setCellValue("Author_Name");
		row.createCell(3).setCellValue("Price");
		
		int dataRowIndex = 1;
		
		for(Book book : books)
		{
			HSSFRow dataRow = sheet.createRow(dataRowIndex);
			dataRow.createCell(0).setCellValue(book.getId());
			dataRow.createCell(1).setCellValue(book.getName());
			dataRow.createCell(2).setCellValue(book.getAuthor());
			dataRow.createCell(3).setCellValue(book.getPrice());
			dataRowIndex++;
		
			  try {
		            int price = Integer.parseInt(book.getPrice());
		            if (price > 100) {
		                HSSFCellStyle customStyle = workbook.createCellStyle();
		                customStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		                customStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		                for (int i = 0; i < dataRow.getLastCellNum(); i++) {
		                    dataRow.getCell(i).setCellStyle(customStyle);
		                }
		            }
		        } catch (NumberFormatException e) {
		            // Handle if the price is not a valid integer
		            System.out.println("Invalid price format for book: " + book.getName());
		        }

		        dataRowIndex++;
		    }
          
		ServletOutputStream ops = response.getOutputStream();
		workbook.write(ops);
		workbook.close();
		ops.close();
	}
	
//    public void importData(MultipartFile file) {
//        try {
//            InputStream inputStream = file.getInputStream();
//            XSSFWorkbook wBook = new XSSFWorkbook(inputStream);
//            XSSFSheet eSheet = wBook.getSheetAt(0); // Assuming only one sheet
//            
//            XSSFCellStyle style=wBook.createCellStyle(); 
//            style.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex()); 
//            style.setFillPattern(FillPatternType.DIAMONDS); 
//            
//            XSSFCell cell=row.createCell(1); 
//            cell.setCellValue("welcome"); 
//            cell.setCellStyle(style); 
//            
//            for (Row row : eSheet) {
//	            Book entity = new Book();
//	            entity.setName(row.getCell(0).getStringCellValue()); // Assuming first cell is name
//	            entity.setAuthor(row.getCell(1).getStringCellValue());
//	            entity.setPrice(row.getCell(2).getStringCellValue());// Assuming second cell is author
//	            bRepo.save(entity);
//            }
//            
            

//            Iterable<Row> rowIterator = eSheet.iterator();
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                // Assuming each row has two columns (name and age)
//                String name = row.getCell(0).getStringCellValue();
//                int age = (int) row.getCell(1).getNumericCellValue();
//                // Now, save this data to your database
//                // Example: userRepository.save(new User(name, age));
//            }

//            wBook.close();
//        } catch (IOException | EncryptedDocumentException e) {
//            e.printStackTrace();
//        }
//    }
//   
    
//	 public void importDataFromExcel(InputStream inputStream) throws IOException {
//	        XSSFWorkbook wBook = new XSSFWorkbook(inputStream);
//	        org.apache.poi.ss.usermodel.Sheet eSheet = wBook.getSheetAt(0); // Assuming first sheet
//	        
//	        for (Row row : eSheet) {
//	            Book entity = new Book();
//	            entity.setName(row.getCell(0).getStringCellValue()); // Assuming first cell is name
//	            entity.setAuthor(row.getCell(1).getStringCellValue());
//	            entity.setPrice(row.getCell(2).getStringCellValue());// Assuming second cell is author
//	            bRepo.save(entity);
//	        }
//	        eSheet.forEach(s -> {
//	        	Book entity = new Book();
//	            entity.setName(s.getCell(0).getStringCellValue()); // Assuming first cell is name
//	            entity.setAuthor(s.getCell(1).getStringCellValue());
//	            entity.setPrice(s.getCell(2).getStringCellValue());// Assuming second cell is author
//	            bRepo.save(entity);
//	        });
}

 