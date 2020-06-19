import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.text.DateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import javax.faces.context.FacesContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Versão 2.0 de POIUtil. Data: 17/06/2020
 * Principal alteração -> Edição/criação de células/regiões via String. Ex: "A1:B1"
 * Exemplo de uso no main
 * 
 * @author p067613 - https://github.com/joaopmi/PoiUtil
 *
 */
public class POIUtil2 implements Serializable {

	/** Longs */
	private static final long serialVersionUID = 1191392087071562651L;
	/**Integer*/
	private static final int INTERVALO_A_A = 26;
	/**String*/
	private final String colunas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	private static final String REGEX_CELULAS = "^[A-Z]+\\d+$";
	private static final String REGEX_REGIOES = "^[A-Z]+\\d+:[A-Z]+\\d+?$";
	private static final String REGEX_CELULAS_REGIOES = "^[A-Z]+\\d+(:[A-Z]+\\d+)?$";
	private static final String REGEX_APENAS_LETRAS = "[^A-Z]";
	private static final String REGEX_APENAS_NUMEROS = "[^\\d]";
	private static final String CALIBRI = "Calibri";
	/** XSSF */
	private transient XSSFWorkbook workbook;
	/** Maps */
	private transient Map<String, XSSFCellStyle> cellStyles = new HashMap<>();
	private transient Map<String, XSSFFont> fonts = new HashMap<>();

	// CONSTRUTORES

	/** Inicializa um novo workbook vazio. */
	public POIUtil2() {
		this.workbook = new XSSFWorkbook();
	}
	
	/** Inicializa o workbook a partir de um array de bytes de um excel. 
	 * @throws IOException */
	public POIUtil2(final byte[] workbook) throws IOException {
		this.workbook = new XSSFWorkbook(new ByteArrayInputStream(workbook));
	}

	/**
	 * Inicializa um workbook através do path. Utiliza FileInputStream que é fechada
	 * no finally
	 * 
	 * @param path (String) - caminho do excel
	 * @throws IOException
	 */
	public POIUtil2(final String path) throws IOException {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(new File(path));
			this.workbook = new XSSFWorkbook(fis);
		} finally {
			if (null != fis) {
				fis.close();
			}
		}
	}

	// WORKBOOK

	/**
	 * Retorna workbook criado no construtor
	 * 
	 * @return XSSFWorkbook
	 */
	public XSSFWorkbook getWorkbook() {
		return this.workbook;
	}
	
	/**
	 * Escreve excel através do FileOutputStream com o caminho informado.
	 * FileOutputStream e workbook fechados no finally
	 * 
	 * @param path (String) - caminho destino para escrever excel
	 * @throws IOException
	 */
	public void write(final String path) throws IOException {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(path));
			autoSizeColumns().evaluateAllFormulas();
			this.workbook.write(fos);
		} finally {
			if (null != fos) {
				fos.close();
			}
			close();
		}
	}

	/**
	 * Escreve no buffer e fecha o workbook no finally
	 * 
	 * @param ByteArrayOutputStream baos - buffer para escrever os bytes
	 * @throws IOException
	 */
	public void write(final ByteArrayOutputStream baos) throws IOException {
		try {
			autoSizeColumns().evaluateAllFormulas();
			this.workbook.write(baos);
		} finally {
			close();
		}
	}

	/**
	 * Escreve no buffer ByteArrayOutputStream, fecha o buffer e workbook no finally
	 * 
	 * @return byte[]
	 * @throws IOException
	 */
	public byte[] write() throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			autoSizeColumns().evaluateAllFormulas();
			this.workbook.write(baos);
		} finally {
			baos.close();
			close();
		}
		return baos.toByteArray();
	}

	/**
	 * Configura todas as fórmulas criadas nas células. É chamado automaticamente
	 * nos métodos write().
	 * @return PoiUtil
	 */
	public POIUtil2 evaluateAllFormulas() {
		XSSFFormulaEvaluator.evaluateAllFormulaCells(this.workbook);
		return this;
	}
	
	/**
	 * Configura em todas as folhas a largura de todas as colunas para "auto"
	 * @return PoiUtil
	 */
	public POIUtil2 autoSizeColumns() {
		final int qtdSheets = this.workbook.getNumberOfSheets();
		int index = 0;
		while(index < qtdSheets) {
			final XSSFSheet sheet = this.getSheetAt(index);
			int index2 = 0;
			while(index2 < sheet.getLastRowNum()) {
				sheet.autoSizeColumn(index2);
				index2++;
			}
			index++;
		}
		return this;
	}
	
	/**
	 * Download do workbook como .xlsx
	 * @param fileName (String) - nome do arquivo. Caso não possua o .xlsx será inserido.
	 * @throws IOException 
	 */
	public void download(String fileName) throws IOException {
		if(!fileName.contains(".xlsx")) {
			fileName = fileName + ".xlsx";
		}
		final ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ServletOutputStream sos = null;
		try {
			this.write(baos);
			final byte[] bytes = baos.toByteArray();
			final HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext().getResponse();
			response.setHeader("Content-disposition", "attachment;filename="+fileName);
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			response.setContentLength(bytes.length);
			sos = response.getOutputStream();
			sos.write(bytes,0,bytes.length);
			sos.flush();
		}finally {
			FacesContext.getCurrentInstance().responseComplete();
			close(baos, sos);
		}
	}

	/**
	 * Fecha o workbook. É fechado automaticamente nós metodos write()
	 * 
	 * @throws IOException
	 */
	public void close() throws IOException {
		this.workbook.close();
	}

	// SHEET

	/**
	 * Cria uma nova XSSFSheet nomeada
	 * 
	 * @param name (String) - nome da folha
	 */
	public POIUtil2 createSheet(final String name) {
		this.workbook.createSheet(name);
		return this;
	}

	/**
	 * Retorna folha da posição informada.
	 * 
	 * @param index (int) - Posição da folha. Base 0
	 * @return XSSFSheet
	 */
	public XSSFSheet getSheetAt(final int index) {
		return this.workbook.getSheetAt(index);
	}
	
	/**
	 * Configura folha para padrão
	 * @param index (int) - Índice (base 0) da folha
	 * @return PoiUtil
	 */
	public POIUtil2 setActiveSheet(final int index) {
		this.workbook.setActiveSheet(index);
		return this;
	}

	/**
	 * Retorna folha pelo nome informado
	 * 
	 * @param name (String) - nome da folha
	 * @return XSSFSheet
	 */
	public XSSFSheet getSheet(final String name) {
		return this.workbook.getSheet(name);
	}
	
	/**
	 * Remove folha no index informado. Base 0
	 * @param index (int) index base 0
	 * @return POIUtil
	 */
	public POIUtil2 removeSheetAt(final int index) {
		this.workbook.removeSheetAt(index);
		return this;
	}

	/**
	 * Cria uma nova área mesclada
	 * 
	 * @param sheet    (XSSFSheet) - folha a ser alterada
	 * @param regioes (String...) - regiões a serem mescladas (separadas por vírgula). Ex: "A1:F10","B5:G30"...
	 * @throws Exception - regiões inválidas 
	 * @return PoiUtil
	 */
	public POIUtil2 createMergedRegions(final XSSFSheet sheet, final String... regioes) throws Exception {
		validarRegioes(regioes);
		for(int index = 0; index < regioes.length; index++) {
			final String regiao = regioes[index];
			final int[][] linhasColunas = getRowsCols(regiao);
			sheet.addMergedRegion(new CellRangeAddress(linhasColunas[0][0], linhasColunas[1][0], linhasColunas[0][1], linhasColunas[1][1]));			
		}
		return this;
	}

	// CELL

	/**
	 * Cria célula
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha não exista
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						this.getRow(sheet, linhaInicial).createCell(coluna++);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
			}
		}
		return this;
	}
	
	/**
	 * Cria célula configurada com estilo informado
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo adicionado no map de estilos -> createCellStyle().
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String cellStyleName, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						this.getRow(sheet, linhaInicial).createCell(coluna++).setCellStyle(this.cellStyles.get(cellStyleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]).setCellStyle(this.cellStyles.get(cellStyleName));
			}
		}
		return this;
	}
	
	/**
	 * Cria célula configurada com estilo e valor String informados
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo adicionado no map de estilos -> createCellStyle().
	 * @param cellValue (String) - valor em String para a célula
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha ou estilo não existam ou valor não tenha sido informado
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String cellStyleName, final String cellValue, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						final XSSFCell cell = this.getRow(sheet, linhaInicial).createCell(coluna++);
						cell.setCellStyle(this.cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this.cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		return this;
	}
	
	/**
	 * Cria célula configurada com estilo e valor long informados
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo adicionado no map de estilos -> createCellStyle().
	 * @param cellValue (long) - valor em long para a célula
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String cellStyleName, final long cellValue, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						final XSSFCell cell = this.getRow(sheet, linhaInicial).createCell(coluna++);
						cell.setCellStyle(this.cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this.cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		return this;
	}
	
	/**
	 * Cria célula configurada com estilo e valor long informados
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo adicionado no map de estilos -> createCellStyle().
	 * @param cellValue (int) - valor em int para a célula
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String cellStyleName, final int cellValue, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						final XSSFCell cell = this.getRow(sheet, linhaInicial).createCell(coluna++);
						cell.setCellStyle(this.cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this.cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		return this;
	}
	
	/**
	 * Cria célula configurada com estilo e valor long informados
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo adicionado no map de estilos -> createCellStyle().
	 * @param cellValue (double) - valor em double para a célula
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D1 (criará células de A1 até D1 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final String cellStyleName, final double cellValue, final String... celulasRegioes) throws Exception {
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						final XSSFCell cell = this.getRow(sheet, linhaInicial).createCell(coluna++);
						cell.setCellStyle(this.cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this.cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		return this;
	}

	/**
	 * Configura em string o valor da célula
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param value (String) - valor a ser configurado na célula
	 * @param celula (String) - célula a ser configurada. Ex: "A1"
	 * @throws NullPointerException - caso linha ou célula não existam ou valor seja null
	 * @throws Exception - caso célula informada seja inválida 
	 * @return PoiUtil
	 */
	public POIUtil2 setCellValue(final XSSFSheet sheet, final String value, final String celula) throws Exception {
		validarCelulas(new String[] {celula});
		final int[] linhaColuna = getRowCol(celula);
		sheet.getRow(linhaColuna[0]).getCell(linhaColuna[1]).setCellValue(value);
		return this;
	}
	
	/**
	 * Configura em string o valor da célula e seu estilo
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo
	 * @param value (String) - valor a ser configurado na célula
	 * @param celula (String) - célula a ser configurada. Ex: "A1"
	 * @throws NullPointerException - caso linha,célula ou estilo não existam ou valor seja null
	 * @throws Exception - caso célula informada seja inválida
	 * @return PoiUtil
	 */
	public POIUtil2 setCellValue(final XSSFSheet sheet, final String cellStyleName,final String value, final String celula) throws Exception {
		validarCelulas(new String[] {celula});
		final int[] linhaColuna = getRowCol(celula);
		final XSSFCell xssfCell = sheet.getRow(linhaColuna[0]).getCell(linhaColuna[1]);
		xssfCell.setCellValue(value);
		xssfCell.setCellStyle(this.getCellStyle(cellStyleName));
		return this;
	}

	/**
	 * Configura em double o valor da célula
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param value (String) - valor a ser configurado na célula
	 * @param celula (String) - célula a ser configurada. Ex: "A1"
	 * @throws NullPointerException - caso linha ou célula não existam
	 * @throws Exception - caso célula informada seja inválida
	 * @return PoiUtil
	 */
	public POIUtil2 setCellValue(final XSSFSheet sheet, final double value, final String celula) throws Exception {
		validarCelulas(new String[] {celula});
		final int[] linhaColuna = getRowCol(celula);
		sheet.getRow(linhaColuna[0]).getCell(linhaColuna[1]).setCellValue(value);
		return this;
	}
	
	/**
	 * Configura em double o valor da célula e seu estilo
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param cellStyleName (String) - nome do estilo
	 * @param value (String) - valor a ser configurado na célula
	 * @param celula (String) - célula a ser configurada. Ex: "A1"
	 * @throws NullPointerException - caso linha,célula,estilo não existam
	 * @throws Exception - caso célula informada seja inválida
	 * @return PoiUtil
	 */
	public POIUtil2 setCellValue(final XSSFSheet sheet, final String cellStyleName,final double value, final String celula) throws Exception {
		validarCelulas(new String[] {celula});
		final int[] linhaColuna = getRowCol(celula);
		final XSSFCell xssfCell = sheet.getRow(linhaColuna[0]).getCell(linhaColuna[1]);
		xssfCell.setCellValue(value);
		xssfCell.setCellStyle(this.getCellStyle(cellStyleName));
		return this;
	}

	/**
	 * Retorna célula. Retorna nulo caso célula não exista
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param celula (String) - célula a ser configurada. Ex: "A1"
	 * @throws NullPointerException - caso linha não exista
	 * @throws Exception - caso célula informada seja inválida
	 * @return XSSFCell
	 */
	public XSSFCell getCell(final XSSFSheet sheet, final String celula) throws Exception {
		validarCelulas(new String[] {celula});
		final int[] linhaColuna = getRowCol(celula);
		return this.getRow(sheet,linhaColuna[0]).getCell(linhaColuna[1]);
	}

	// ROW

	/**
	 * Cria linha
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row   (int) - número da linha a ser criada. Base 0
	 * @return PoiUtil
	 */
	public POIUtil2 createRow(final XSSFSheet sheet, final int row) {
		sheet.createRow(row);
		return this;
	}
	
	/**
	 * Cria linhas
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param rows  (int[]) - array de linhas a serem criadas. Base 0
	 * @return PoiUtil
	 */
	public POIUtil2 createRows(final XSSFSheet sheet, final int... rows) {
		for(final int row : rows) {
			sheet.createRow(row);
		}
		return this;
	}

	/**
	 * Retorna linha. Retorna nulo caso não exista.
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row   (int) - número da linha a retornar. Base 0
	 * @return XSSFRow
	 */
	public XSSFRow getRow(final XSSFSheet sheet, final int row) {
		return sheet.getRow(row);
	}

	// CELLSTYLE

	/**
	 * Cria um XSSFCellStyle e o armazena no HashMap
	 * 
	 * @param cellStyleName (String) - nome chave do estilo no HashMap
	 * @return PoiUtil
	 */
	public POIUtil2 createCellStyle(final String cellStyleName) {
		this.cellStyles.put(cellStyleName, this.workbook.createCellStyle());
		return this;
	}
	
	/**
	 * Cria e configura XSSFCellStyle no HashMap<String,XSSFCellStyle>
	 * @param name (String) - key no hashmap
	 * @param font (String) - key da fonte no HashMap<String,XSSFFont>
	 * @param hAlign (HorizontalAlignment) - alinhameto horizontal texto
	 * @param vAlign (VerticalAlignment) - posialinhametoção vertical texto
	 * @param fill (FillPatternType) - padrão de preenchimento do background
	 * @param indexedColor (IndexedColors) - índice da cor de preenchimento
	 * @param borders - (BorderStyle...) - array de bordas -> Top, Right, Bottom, Left 
	 * @return
	 */
	public POIUtil2 createCellStyle(final String name, final String font,final boolean wrapText, final HorizontalAlignment hAlign, final VerticalAlignment vAlign,
			final FillPatternType fill, final IndexedColors indexedColor,final BorderStyle... borders) {
		
		createCellStyle(name).editCellStyleFont(name, font).editCellStyleWrapText(name, wrapText).editCellStyleAlignment(name, hAlign, vAlign)
		.editCellStyleFillPattern(name, fill).editCellStyleForegroundColor(name, indexedColor);
		if(null != borders) {
			final BorderStyle[] allBorder = new BorderStyle[4];
			for(int index = 0; index < borders.length; index++) {
				allBorder[index] = borders[index];
			}
			editCellStyleBorder(name, allBorder[0], allBorder[1], allBorder[2], allBorder[3]);
		}
		return this;
	}

	/**
	 * Retorna XSSFCellStyle armazenado no HashMap pela key nome. Retorna nulo caso
	 * não exista
	 * 
	 * @param nome (String) - nome chave do estilo no HashMap
	 * @return XSSFCellStyle
	 */
	public XSSFCellStyle getCellStyle(final String nome) {
		return this.cellStyles.get(nome);
	}

	/**
	 * Edita borda do estilo contido no HashMap.
	 * 
	 * @param cellStyleName (String) - nome chave da célula
	 * @param borderTop     (BorderStyle) - borda topo. Passe nulo para não setar
	 * @param borderRight   (BorderStyle) - borda direita. Passe nulo para não setar
	 * @param borderBottom  (BorderStyle) - borda inferior. Passe nulo para não
	 *                      setar
	 * @param borderLeft    (BorderStyle) - borda esquerda. Passe nulo para não
	 *                      setar
	 * @throws NullPointerException - caso estilo não exista
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleBorder(final String cellStyleName, final BorderStyle borderTop,
			final BorderStyle borderRight, final BorderStyle borderBottom, final BorderStyle borderLeft) {

		final XSSFCellStyle cellStyle = this.getCellStyle(cellStyleName);
		if (null != borderTop) {
			cellStyle.setBorderTop(borderTop);
		}
		if (null != borderRight) {
			cellStyle.setBorderRight(borderRight);
		}
		if (null != borderBottom) {
			cellStyle.setBorderBottom(borderBottom);
		}
		if (null != borderLeft) {
			cellStyle.setBorderLeft(borderLeft);
		}
		return this;
	}

	/**
	 * Edita todas as bordas do estilo contido no HashMap.
	 * 
	 * @param cellStyleName (String) - nome chave da célula
	 * @param borderTop     (BorderStyle) - Tipo de borda
	 * @throws NullPointerException - caso estilo não exista
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleBorderAll(final String cellStyleName, final BorderStyle border) {

		final XSSFCellStyle cellStyle = this.getCellStyle(cellStyleName);
		cellStyle.setBorderTop(border);
		cellStyle.setBorderRight(border);
		cellStyle.setBorderBottom(border);
		cellStyle.setBorderLeft(border);
		return this;
	}

	/**
	 * Configura fonte do estilo
	 * 
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param fontName      (String) - nome da fonte para o documento
	 * @throws NullPointerException - caso estilo não exista
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleFont(final String cellStyleName, final String fontName) {
		if(null != fontName) {
			this.getCellStyle(cellStyleName).setFont(this.getFont(fontName));			
		}
		return this;
	}

	/**
	 * Edita o alinhamento horizontal e vertical da célula
	 * 
	 * @param cellStyleName       (String) - nome chave da fonte no HashMap
	 * @param horizontalAlignment (HorizontalAlignment) - alinhamento horizontal.
	 *                            Passe null para não setar
	 * @param verticalAlignment   (VerticalAlignment) - alinhamento vertical. Passe
	 *                            null para não setar
	 * @throws NullPointerException - caso estilo não exista
	 * @return
	 */
	public POIUtil2 editCellStyleAlignment(final String cellStyleName, final HorizontalAlignment horizontalAlignment,
			final VerticalAlignment verticalAlignment) {

		final XSSFCellStyle estilo = this.getCellStyle(cellStyleName);
		if (null != horizontalAlignment) {
			estilo.setAlignment(horizontalAlignment);
		}
		if (null != verticalAlignment) {
			estilo.setVerticalAlignment(verticalAlignment);
		}
		return this;
	}

	/**
	 * Configura se há quebra de texto no estilo
	 * 
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param wrap          (boolean) - se texto deve quebrar
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleWrapText(final String cellStyleName, final boolean wrap) {
		this.getCellStyle(cellStyleName).setWrapText(wrap);
		return this;
	}
	
	/**
	 * Configura padrão de preenchimento do estilo
	 * 
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param fillPattern   (FillPatternType) - padrão de preenchimento
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleFillPattern(final String cellStyleName, final FillPatternType fillPattern) {
		if(null != fillPattern) {
			this.getCellStyle(cellStyleName).setFillPattern(fillPattern);			
		}
		return this;
	}
	
	/**
	 * Configura cor de fundo do estilo
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param color (IndexedColors) - cor de fundo
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleForegroundColor(final String cellStyleName, final IndexedColors color) {
		if(null != color) {
			this.getCellStyle(cellStyleName).setFillForegroundColor(color.index);			
		}
		return this;
	}

	/**
	 * Configura um formato de valor para o estilo. Ex: R$0.00
	 * 
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param format        (String) - formato de valor
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleFormatValue(final String cellStyleName, final String format) {
		final XSSFCreationHelper helper = this.workbook.getCreationHelper();
		this.getCellStyle(cellStyleName).setDataFormat(helper.createDataFormat().getFormat(format));
		return this;
	}

	// ROW CELL

	/**
	 * Cria linhas e células
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D5 (criará células e linhas de A1 até D5 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createRowCell(final XSSFSheet sheet, final String... celulasRegioes) throws Exception{
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					final XSSFRow row = this.createRow(sheet, linhaInicial).getRow(sheet, linhaInicial);
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						row.createCell(coluna++);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.createRow(sheet, linhaColuna[0]).getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
			}
		}
		return this;
	}

	/**
	 * Cria linhas e células e configura estilo nas células
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param celulasRegioes (String...) - celulas ou regiões a serem criadas. Ex: A1 ou A1:D5 (criará células e linhas de A1 até D5 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso estilo não exista
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 createRowCellStyle(final XSSFSheet sheet, final String styleName, final String... celulasRegioes) throws Exception{
		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					final XSSFRow row = this.createRow(sheet, linhaInicial).getRow(sheet, linhaInicial);
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						row.createCell(coluna++).setCellStyle(this.cellStyles.get(styleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.createRow(sheet, linhaColuna[0]).getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]).setCellStyle(this.cellStyles.get(styleName));
			}
		}
		return this;
	}

	/**
	 * Edita estilo das células
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param celulasRegioes (String...) - celulas ou regiões a serem editadas. Ex: A1 ou A1:D5 (criará células e linhas de A1 até D5 inclusa) Separar por vírgula ("A1","A2:B3"...).
	 * @throws NullPointerException - caso estilo não exista
	 * @throws Exception - células/regiões inválidas
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleInRowsCells(final XSSFSheet sheet, final String cellStyleName, final String... celulasRegioes) throws Exception{

		validarCelulasRegioes(celulasRegioes);
		final String doisPontos = ":";
		for(int i = 0; i < celulasRegioes.length; i++) {
			if(celulasRegioes[i].indexOf(doisPontos) >= 0) {
				final int[][] linhasColunas = getRowsCols(celulasRegioes[i]);
				int linhaInicial = linhasColunas[0][0];
				final int linhaFinal = linhasColunas[1][0];
				final int colunaInicial = linhasColunas[0][1];
				final int colunaFinal = linhasColunas[1][1];
				while(linhaInicial <= linhaFinal) {
					int coluna = colunaInicial;
					while(coluna <= colunaFinal) {
						this.getRow(sheet, linhaInicial).getCell(coluna).setCellStyle(this.cellStyles.get(cellStyleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.getRow(sheet, linhaColuna[0]).getCell(linhaColuna[1]).setCellStyle(this.cellStyles.get(cellStyleName));
			}
		}
		return this;
	}

	// FONTS

	/**
	 * Cria e armazena XSSFFont no HashMap<String,XSSFFont>
	 * 
	 * @param name (String) - nome chave da fonte
	 * @return PoiUtil
	 */
	public POIUtil2 createFont(final String name) {
		this.fonts.put(name, this.workbook.createFont());
		return this;
	}
	
	/**
	 * Cria e configura XSSFFont no HashMap<String,XSSFFont>
	 * @param name (String) - key do hashmap
	 * @param fontName (String) - nome da fonte no excel
	 * @param bold (boolean) - Configura negrito
	 * @param size (double) - Tamanho da fonte
	 * @return
	 */
	public POIUtil2 createFont(final String name, final String fontName, final boolean bold, final double size) {
		createFont(name).editFontName(name, fontName).editFontBold(name, bold).editFontSize(name, size);
		return this;
	}

	/**
	 * Configura negrito da fonte
	 * 
	 * @param name (String) - nome chave da fonte no HashMap
	 * @param bold (boolean) - se é negrito ou não
	 * @return PoiUtil
	 */
	public POIUtil2 editFontBold(final String name, final boolean bold) {
		this.getFont(name).setBold(bold);
		return this;
	}
	
	/**
	 * Configura cor da Fonte. (IndexedColors)
	 * @param name (String) - nome chave da fonte no HashMap
	 * @param color (IndexedColors) - cor da fonte
	 * @return PoiUtil
	 */
	public POIUtil2 editFontColor(final String name, final IndexedColors color) {
		this.getFont(name).setColor(color.getIndex());
		return this;
	}

	/**
	 * Edita nome da fonte do documento
	 * 
	 * @param name     (String) - nome chave da fonte no HashMap
	 * @param fontName (String) - nome da fonte no documento. Ex: Calibri
	 * @return PoiUtil
	 */
	public POIUtil2 editFontName(final String name, final String fontName) {
		this.getFont(name).setFontName(fontName);
		return this;

	}

	public POIUtil2 editFontSize(final String name, final double size) {
		this.getFont(name).setFontHeight(size);
		return this;
	}

	/**
	 * Retorna fonte pela key nome
	 * 
	 * @param name (String) - nome chave da fonte
	 * @return XSSFFont
	 */
	public XSSFFont getFont(final String name) {
		return this.fonts.get(name);
	}

	// IMAGES

	/**
	 * Insere imagem nos pontos XY definidos. Recebe o caminho da imagem (imagePath)
	 * e lê/escreve os bytes com FileInputStream e ByteArrayOutputStream. Ambos são
	 * fechados no finally.
	 * 
	 * @param sheet      (XSSFSheet) - folha a ser alterada
	 * @param regiao (String) - linha e células que imagem ocupará. Ex: "A1:H5"
	 * @param scaleX     (double) - Escala (tamanho) em X da imagem de acordo com
	 *                   row1, row2, cell1, cell2
	 * @param scaleY     (double) - Escala (tamanho) em Y da imagem de acordo com
	 *                   row1, row2, cell1, cell2
	 * @param dx1        (int) - Posição X do canto superior esquerdo em relação ao
	 *                   col1
	 * @param dx2        (int) - Posição X do canto inferior direito em relação ao
	 *                   col2
	 * @param dy1        (int) - Posição Y do canto superior esquerdo em relação ao
	 *                   row1
	 * @param dy2        (int) - Posição Y do canto inferior direito em relação ao
	 *                   row2
	 * @param anchorType (AnchorType) - Comportamento da imagem
	 * @param imagePath  (String) - Caminho da imagem
	 * @throws Exception
	 */
	public void insertImage(final XSSFSheet sheet, final String regiao, final double scaleX, final double scaleY, final int dx1, final int dx2, final int dy1, final int dy2,
			final AnchorType anchorType, final String imagePath) throws Exception {

		validarRegioes(new String[] {regiao});
		final int[][] linhasColunas = getRowsCols(regiao);
		FileInputStream fis = null;
		ByteArrayOutputStream baos = null;
		try {
			fis = new FileInputStream(new File(imagePath));
			baos = new ByteArrayOutputStream();
			byte[] bufferImg = new byte[4096];
			int tamanhoLido = 0;
			while ((tamanhoLido = fis.read(bufferImg)) != -1) {
				baos.write(bufferImg, 0, tamanhoLido);
			}
			final CreationHelper helper = this.workbook.getCreationHelper();
			final XSSFDrawing drawing = sheet.createDrawingPatriarch();
			final ClientAnchor anchor = helper.createClientAnchor();
			anchor.setAnchorType(anchorType);
			anchor.setCol1(linhasColunas[0][1]);
			anchor.setCol2(linhasColunas[1][1]);
			anchor.setRow1(linhasColunas[0][0]);
			anchor.setRow2(linhasColunas[1][0]);
			anchor.setDx1(dx1 * Units.EMU_PER_POINT);
			anchor.setDx2(dx2 * Units.EMU_PER_POINT);
			anchor.setDy1(dy1 * Units.EMU_PER_POINT);
			anchor.setDy2(dy2 * Units.EMU_PER_POINT);
			final int pictureIndex = this.workbook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
			final Picture picture = drawing.createPicture(anchor, pictureIndex);
			picture.resize(scaleX, scaleY);
		} finally {
			if (null != fis) {
				fis.close();
			}
			if (null != baos) {
				baos.close();
			}
		}
	}
	
	/**
	 * Insere imagem nos pontos XY definidos.
	 * 
	 * @param sheet      (XSSFSheet) - folha a ser alterada
	 * @param regiao (String) - linha e células que imagem ocupará. Ex: "A1:H5"
	 * @param scaleX     (double) - Escala (tamanho) em X da imagem de acordo com
	 *                   row1, row2, cell1, cell2
	 * @param scaleY     (double) - Escala (tamanho) em Y da imagem de acordo com
	 *                   row1, row2, cell1, cell2
	 * @param dx1        (int) - Posição X do canto superior esquerdo em relação ao
	 *                   col1
	 * @param dx2        (int) - Posição X do canto inferior direito em relação ao
	 *                   col2
	 * @param dy1        (int) - Posição Y do canto superior esquerdo em relação ao
	 *                   row1
	 * @param dy2        (int) - Posição Y do canto inferior direito em relação ao
	 *                   row2
	 * @param anchorType (AnchorType) - Comportamento da imagem
	 * @param fillColorRgb (int[]) - Cor de fundo da imagem em RGB.
	 * @param image  (byte[]) - Byte array da imagem
	 * @throws Exception 
	 */
	public void insertImage(final XSSFSheet sheet, final String regiao,	final double scaleX, final double scaleY, final int dx1, final int dx2, final int dy1, final int dy2,
			final AnchorType anchorType, final int[] fillColorRgb,final byte[] image) throws Exception {
		
		validarRegioes(new String[] {regiao});
		final int[][] linhasColunas = getRowsCols(regiao);
		
			final CreationHelper helper = this.workbook.getCreationHelper();
			final XSSFDrawing drawing = sheet.createDrawingPatriarch();
			final ClientAnchor anchor = helper.createClientAnchor();
			anchor.setAnchorType(anchorType);
			anchor.setCol1(linhasColunas[0][1]);
			anchor.setCol2(linhasColunas[1][1]);
			anchor.setRow1(linhasColunas[0][0]);
			anchor.setRow2(linhasColunas[1][0]);
			anchor.setDx1(dx1 * Units.EMU_PER_POINT);
			anchor.setDx2(dx2 * Units.EMU_PER_POINT);
			anchor.setDy1(dy1 * Units.EMU_PER_POINT);
			anchor.setDy2(dy2 * Units.EMU_PER_POINT);
			final int pictureIndex = this.workbook.addPicture(image, Workbook.PICTURE_TYPE_PNG);
			final Picture picture = drawing.createPicture(anchor, pictureIndex);
			picture.setFillColor(fillColorRgb[0], fillColorRgb[1], fillColorRgb[2]);
			picture.resize(scaleX, scaleY);
	}

	// COLUMNS

	/**
	 * Aplica o autoSize nas colunas informadas.
	 * 
	 * @param sheet   (XSSFSheet) - folha a ser alterada
	 * @param columns (int[]) - colunas a aplicar o autoSize. Base 0
	 * @return PoiUtil
	 */
	public POIUtil2 autoSizeColumns(final XSSFSheet sheet, final int[] columns) {
		for (final int column : columns) {
			sheet.autoSizeColumn(column);
		}
		return this;
	}
	
	//OUTROS
	private void close(final ByteArrayOutputStream baos, final ServletOutputStream sos) {
		if (null != baos) {
			try {
				baos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if (null != sos) {
			try {
				sos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	/**
	 * Valida células e lança erro caso alguma esteja errada
	 * @param celulas (String[]) array contendo todas as células (separadas por vírgula). Ex: "A1","B3"...
	 * @throws Exception
	 */
	private void validarCelulas(final String[] celulas) throws Exception {
		for(int index = 0; index < celulas.length; index++) {
			if(!celulas[index].matches(REGEX_CELULAS)) {
				throw new Exception("Célula inválida -> " + celulas[index]);
			}
		}
	}
	
	/**
	 * Valida regiões e lança erro caso alguma esteja errada
	 * @param regioes (String[]) array contendo todas as regiões (separadas por vírgula). Ex: "A1:F1","D4:F7"
	 * @throws Exception
	 */
	private void validarRegioes(final String[] regioes) throws Exception {
		for(int index = 0; index < regioes.length; index++) {
			if(!regioes[index].matches(REGEX_REGIOES)) {
				throw new Exception("Região inválida -> " + regioes[index]);
			}
		}
	}
	
	/**
	 * Valida células e regiões e lança erro caso alguma esteja errada
	 * @param celulasRegioes (String[]) array contendo todas as células/regiões (separadas por vírgula). Ex: "A1:F1","A3"...
	 * @throws Exception
	 */
	private void validarCelulasRegioes(final String[] celulasRegioes) throws Exception {
		for(int index = 0; index < celulasRegioes.length; index++) {
			if(!celulasRegioes[index].matches(REGEX_CELULAS_REGIOES)) {
				throw new Exception("Célula/Região inválida -> " + celulasRegioes[index]);
			}
		}
	}
	
	/**
	 * Recupera o numero da linha e coluna de determinada célula
	 * @param celula (String) Ex: "A1"
	 * @return array com linha(0) e coluna(1) (int[]) 
	 */
	private int[] getRowCol(final String celula) {
		final String[] colunaArray = celula.replaceAll(REGEX_APENAS_LETRAS, "").split("");
		final int linha = Integer.parseInt(celula.replaceAll(REGEX_APENAS_NUMEROS, ""));
		int coluna = 0;
		if(colunaArray.length == 1) {//SE TAMANHO == 1. EX: "F1" -> A COLUNA É O indexOf OF DA LETRA "F"
			coluna = this.colunas.indexOf(colunaArray[0]);
			return new int[] {linha - 1,coluna};
		}else {
			/*SOMA (indexOf DA ULTIMA LETRA) + (26) * (INTERVAL DA COLUNA A à A). EX: "ACA2" ->
			 * indexOf "A" = 0;
			 * 26 (LETRAS DO ALFABETO. SÃO NECESSÁRIAS 26 COLUNAS PARA IR DE AA ATE BA, ETC)
			 * indexOf "C" + 1 = 3 
			 * 0 + 26 * 3 = 78 (COLUNA CA)
			 * (indexOf de "C" é 2, PORÉM O EXCEL COMEÇA NA COLUNA "A", NÃO "AA", PORTANTO É SOMADO +1, LOGO PARA CHEGAR À COLUNA "CA" MULTIPLICA-SE 26 * 3)*/ 
			coluna = this.colunas.indexOf(colunaArray[colunaArray.length - 1]) + INTERVALO_A_A * (this.colunas.indexOf(colunaArray[colunaArray.length - 2]) + 1); 
			/*ITERA DO FIM PARA O INÍCIO PULANDO AS DUAS ÚLTIMAS LETRAS. A CADA LOOP O INTERVALO É EXPONENCIALMENTE INCREMENTADO. EX:
			 * "XFD1" -> ÚLTIMA COLUNA DE UM DOCUMENTO EXCEL
			 * "FD" = 159  
			 * "X" = 26 * 26 * indexOf "X" = 16.224
			 * 16.224 + 159 = 16383 -> ÚLTIMA COLUNA "XFD"
			 */
			int index = colunaArray.length - 3;
			int aux = 2;
			while(index > -1) {
				coluna += Math.pow(INTERVALO_A_A, aux++) * (this.colunas.indexOf(colunaArray[index])+1);
				index--;
			}
			System.out.println(coluna);
			return new int[] {linha - 1,coluna};
		}
	}
	
	/**
	 * Recupera os números das linhas (iniciais e finais) e colunas (iniciais e finais) de determinada regiao.
	 * @param regiao (String) Ex: "A1:D1"
	 * @return matriz com dois arrays de linha e coluna
	 */
	private int[][] getRowsCols(final String regiao){
		final String[] regiaoArray = regiao.split(":");
		final int[][] retorno = new int[2][1];
		retorno[0] = getRowCol(regiaoArray[0]);
		retorno[1] = getRowCol(regiaoArray[1]);
		return retorno;
	}
	
	//EXEMPLO
	public static void main(String[] args) throws Exception {
		final POIUtil2 poi = new POIUtil2();
		final DateFormat df = DateFormat.getDateTimeInstance(DateFormat.LONG, DateFormat.SHORT, new Locale("pt","BR"));
		final XSSFSheet sheet = poi.createSheet("Planilha 1").getSheetAt(0);
		final String estiloTitulo = "estiloTitulo";
		final String calibriBold11 = "calibriBold11";
		final byte[] arquivo = 
		poi
		.createFont(calibriBold11, POIUtil2.CALIBRI, true, 11d)
		.createCellStyle(estiloTitulo, calibriBold11, true, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, FillPatternType.SOLID_FOREGROUND, IndexedColors.GREY_25_PERCENT,
				BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN)
		.createMergedRegions(sheet, "B1:N3","B4:N4")
		.createRowCellStyle(sheet,estiloTitulo, "B1:N3")
		.createRowCell(sheet, "B4:N4")
		.setCellValue(sheet, "Documento Exemplo POIUtil2", "B1")
		.setCellValue(sheet, "Arquivo gerado em " + df.format(new Date()), "B4")
		.write();
		final FileOutputStream fos = new FileOutputStream(new File(path));
		fos.write(arquivo, 0, arquivo.length);
		fos.flush();
		fos.close();
	}
}
