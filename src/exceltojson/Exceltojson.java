package exceltojson;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.codehaus.jackson.map.ObjectMapper;
import org.omg.CORBA.portable.InputStream;


public class Exceltojson {
	public String doWork(String filename){
		
		File file = new File(filename);
		Workbook wb = null;
		FileInputStream is = null;
		try {//���� ���� ����
			is = new FileInputStream(file);
			wb= WorkbookFactory.create(new FileInputStream(file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		finally {
			if (is != null)
				try {
					is.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}

        ArrayList table = new  ArrayList();
        for( Row row : wb.getSheetAt(0) ) {// �� ����
        	Object str = new String();
        	ArrayList celllist = new ArrayList<>();	
            for( Cell cell : row ) { // ������
                
            	// ���� Ÿ�� ���� �޾Ƽ� �������� �޵� �� ���� �ϳ��� ��Ʈ������ ����
                switch( cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING :
                        str = cell.getRichStringCellValue().getString();
                        break;

                    case Cell.CELL_TYPE_NUMERIC :
                        if(DateUtil.isCellDateFormatted(cell))
                        	str =cell.getDateCellValue().toString();
                        else
                            str = cell.getNumericCellValue();
                        break;
                        
                    case Cell.CELL_TYPE_BOOLEAN :
                        str = cell.getBooleanCellValue();
                        break;

                    case Cell.CELL_TYPE_FORMULA :
                        str = cell.getCellFormula();
                        break;

                }                

                celllist.add(str);
            }
            table.add(celllist);
        }
        ObjectMapper mapper = new ObjectMapper();
        String buffer = "";
		try {
			buffer = mapper.writeValueAsString(table);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		};
		return buffer;
	}
}
