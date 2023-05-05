package it.devlec.exceltopdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import it.devlec.excel.EsempioExcel;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Paths;

public class ExcelToPdf {
    private static final Logger logger = LogManager.getLogger(ExcelToPdf.class);

    public ExcelToPdf() {
    }
    public void excelToPdfTransformer(){
        Document document = new Document();
        logger.debug("Provo a leggere un file excel");
        String excelDiProva = null;
        try {
            excelDiProva = Paths.get(ClassLoader.getSystemResource("excel.xlsx")
                    .toURI()).toString();
        } catch (URISyntaxException e) {
            logger.error("Errore nel trovare nel creare il file");
        }
        FileInputStream file = null;
        File parent = new File(excelDiProva).getParentFile();

        String mioPDF = parent.getAbsolutePath() + File.separator + "fromExcelToPDF.pdf";

        try {
            file = new FileInputStream(excelDiProva);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            FileOutputStream fileOutputStream = new
                    FileOutputStream(
                    mioPDF);
            PdfWriter.getInstance(document,fileOutputStream);
            document.open();
            PdfPTable table = new PdfPTable(3);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i< numberOfRows; i ++){
                Row row = sheet.getRow(i);
                if(i ==0){
                    for (Cell cell : row) {
                        addTableHeader(table, cell.getStringCellValue());
                    }
                }
                String cellValues = new String();
                for (Cell cell : row) {
                    logger.info("Valore " + cell.getStringCellValue());
                    if(cellValues.isEmpty()){
                        cellValues = cell.getStringCellValue();
                    }else{
                        cellValues = ","+cell.getStringCellValue();
                    }
                }
                table.addCell(cellValues);
            }
            workbook.close();
            IOUtils.closeQuietly(file);
            document.add(table);
            document.close();
        } catch (IOException | DocumentException e) {
            logger.error("Errore nel leggere il mio excel", e);
        }
    }

    private void addTableHeader(PdfPTable table, String columnTitle) {

        PdfPCell header = new PdfPCell();
        header.setBackgroundColor(BaseColor.LIGHT_GRAY);
        header.setBorderWidth(2);
        header.setPhrase(new Phrase(columnTitle));
        table.addCell(header);

    }

}
