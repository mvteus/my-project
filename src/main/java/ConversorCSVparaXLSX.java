import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ConversorCSVparaXLSX {
    public static final char DELIMITADOR_CSV = ',';

    public void convertCSVExcel(String arquivoOrigem, String diretorioDestino, String extensaoArquivo)
            throws IllegalArgumentException, IOException {

        Workbook workBook = null;
        FileOutputStream fos = null;

        // Verifica se o arquivo de origem existe.
        File sourceFile = new File(arquivoOrigem);
        if (!sourceFile.exists()) {
            throw new IllegalArgumentException("O arquivo CSV não pode ser encontrado em " + sourceFile);
        }

        // Verifica se a pasta de destino existe para salvar o arquivo Excel.
        File destination = new File(diretorioDestino);
        if (!destination.exists()) {
            throw new IllegalArgumentException(
                    "O diretório destino " + destination + " para o arquivo Excel convertido não existe..");
        }
        if (!destination.isDirectory()) {
            throw new IllegalArgumentException(
                    "O diretório destino " + destination + " para o arquivo Excel não é um diretório/pasta.");
        }

        // Obtendo o objeto BufferedReader
        BufferedReader br = new BufferedReader(new FileReader(sourceFile));

        // Obtendo objetos XSSFWorkbook ou HSSFWorkbook com base na extensão do arquivo Excel repassado
        if (extensaoArquivo.equals(".xlsx")) {
            workBook = new XSSFWorkbook();
        } else {
            workBook = new HSSFWorkbook();
        }

        Sheet sheet = workBook.createSheet("Sheet");

        String nextLine;
        int rowNum = 0;
        while ((nextLine = br.readLine()) != null) {
            Row currentRow = sheet.createRow(rowNum++);
            String rowData[] = nextLine.split(String.valueOf(DELIMITADOR_CSV));
            for (int i = 0; i < rowData.length; i++) {
                if (NumberUtils.isDigits(rowData[i])) {
                    currentRow.createCell(i).setCellValue(Integer.parseInt(rowData[i]));
                } else if (NumberUtils.isNumber(rowData[i])) {
                    currentRow.createCell(i).setCellValue(Double.parseDouble(rowData[i]));
                } else {
                    currentRow.createCell(i).setCellValue(rowData[i]);
                }
            }
        }
        String filename = sourceFile.getName();
        filename = filename.substring(0, filename.lastIndexOf('.'));
        File generatedExcel = new File(diretorioDestino, filename + extensaoArquivo);
        fos = new FileOutputStream(generatedExcel);
        workBook.write(fos);

        try {
            // Fechando os objetos workbook, fos, e br
            workBook.close();
            fos.close();
            br.close();

        } catch (IOException e) {
            System.out.println("Exception ao fechar objetos I/O");
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        long tempoConversao = System.currentTimeMillis();
        boolean convertido = true;

        try {
            ConversorCSVparaXLSX conversor = new ConversorCSVparaXLSX();
            String arquivoOrigem = args[0];
            String diretorioDestino = args[1];
            String extensaoArquivo = args[2];
            conversor.convertCSVExcel(arquivoOrigem, diretorioDestino, extensaoArquivo);
        } catch (Exception e) {
            System.out.println("Exceção inesperada.");
            e.printStackTrace();
            convertido = false;
        }

        if (convertido) {
            System.out.println("Tempo de conversão: " + ((System.currentTimeMillis() - tempoConversao) / 1000) + " segundos.");
        }
    }
}