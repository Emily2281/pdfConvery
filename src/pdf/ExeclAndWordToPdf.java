package pdf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Locale;

import com.aspose.cells.License;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.words.Document;

public class ExeclAndWordToPdf
{

	/**
     * 支持DOC, DOCX, OOXML, RTF, HTML, OpenDocument, PDF, EPUB, XPS, SWF等相互转<br>
     * 
     * @param args
	 * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
    	String fileName = "";
    	/*String filePath = "C:\\Users\\Ai\\Desktop\\临时文件\\外包任务推送操作手册.docx";
    	String filePathWord = wordToPdf(filePath,fileName);
    	System.out.println("word生成pdf文件的路径："+filePathWord);*/
    	String filePath = "C:\\Users\\Ai\\Desktop\\临时文件\\工作簿1.xlsx";
    	String filePathExcel = excelToPdf(filePath,fileName);
    	System.out.println("excel生成pdf文件的路径："+filePathExcel);
    }
    
   

	/**
	 * Excel生成PDF文件
	 * 
	 * @param filePath
	 * @throws Exception
	 */
	public static String excelToPdf(String filePath,String fileName) throws Exception {
		// 验证License
        if (!getExcelLicense()) {
            return "Excel license fault!";
        }
		String filePathBorf = filePath.substring(0, filePath.lastIndexOf("\\")+1);
		if ("".equals(fileName)) {
			fileName = filePath.substring(filePath.lastIndexOf("\\")+1, filePath.indexOf("."));
		}
		
		long old = System.currentTimeMillis();
        Workbook wb = new Workbook(filePath);// 原始excel路径
        String filePathTemp = filePathBorf + fileName+ ".pdf";
        File pdfFile = new File(filePathTemp);// 输出路径
        FileOutputStream fileOS = new FileOutputStream(pdfFile);

        wb.save(fileOS, SaveFormat.PDF);

        long now = System.currentTimeMillis();
        System.out.println("共耗时" + ((now - old) / 1000.0) + "秒");
		return filePathTemp;
	}
	
	/**
	 * word生成PDF文件
	 * 
	 * @param filePath
	 * @throws Exception
	 */
	public static  String wordToPdf(String filePath,String fileName) throws Exception {
		// 验证License
        if (!getWordLicense()) {
            return "Word license fault!";
        }
		long old = System.currentTimeMillis();
		// 打开文档实例
		Document doc = new Document(filePath);
		String filePathBorf = filePath.substring(0, filePath.lastIndexOf("\\")+1);

		if ("".equals(fileName)) {
			fileName = filePath.substring(filePath.lastIndexOf("\\")+1, filePath.indexOf("."));
		}
		String filePathTemp = filePathBorf + fileName+ ".pdf";
		doc.save(filePathTemp, com.aspose.words.SaveFormat.PDF);
		long now = System.currentTimeMillis();
        System.out.println("共耗时" + ((now - old) / 1000.0) + "秒");
		return filePathTemp;
	}
	
    
	/**
	 * 获取license
	 * 
	 * @return
	 */
	public static boolean getExcelLicense() {
		boolean result = false;
		try {
			InputStream is = ExeclAndWordToPdf.class.getClassLoader().getResourceAsStream("\\license.xml");
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
    /**
     * 获取license
     * 
     * @return
     */
    public static boolean getWordLicense() {
        boolean result = false;
        try {
            InputStream is = ExeclAndWordToPdf.class.getClassLoader().getResourceAsStream("\\license.xml");
            com.aspose.words.License aposeLic = new com.aspose.words.License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}