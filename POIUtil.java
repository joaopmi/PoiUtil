import java.awt.Color;
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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Data: 02/07/2020
 * Exemplo de uso no main
 * 
 * https://github.com/joaopmi/PoiUtil
 *
 */
public class POIUtil implements Serializable {

	/** Longs */
	private static final long serialVersionUID = 1191392087071562651L;
	/**Boolean*/
	private boolean auditoriaRegiaoAtualizada;
	/**String*/
	@SuppressWarnings("unused")
	private String _ultimaRegiaoAtualizada;
	private static final String REGEX_CELULAS = "^[A-Z]+\\d+$";
	private static final String REGEX_REGIOES = "^[A-Z]+\\d+:[A-Z]+\\d+?$";
	private static final String REGEX_CELULAS_REGIOES = "^[A-Z]+\\d+(:[A-Z]+\\d+)?$";
	private static final String REGEX_APENAS_LETRAS = "[^A-Z]";
	private static final String REGEX_APENAS_NUMEROS = "[^\\d]";
	private static final String CALIBRI = "Calibri";
	/** XSSF */
	private transient XSSFWorkbook _workbook;
	/** Maps */
	private transient Map<String, XSSFCellStyle> _cellStyles = new HashMap<>();
	private transient Map<String, XSSFFont> _fonts = new HashMap<>();

	// CONSTRUTORES

	/** Inicializa um novo workbook vazio. */
	public POIUtil2() {
		this._workbook = new XSSFWorkbook();
	}
	
	/** Inicializa o workbook a partir de um array de bytes de um excel
	 * @param workbook (byte[]) - array de byte do arquivo 
	 * @throws IOException */
	public POIUtil2(final byte[] workbook) throws IOException {
		this._workbook = new XSSFWorkbook(new ByteArrayInputStream(workbook));
	}

	/**
	 * Inicializa um workbook através do filePath
	 * 
	 * @param filePath (String) - caminho do arquivo
	 * @throws IOException
	 */
	public POIUtil2(final String filePath) throws IOException {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(new File(filePath));
			this._workbook = new XSSFWorkbook(fis);
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
		return this._workbook;
	}
	
	/**
	 * Escreve e fecha workbook no filePath.
	 * 
	 * @param filePath (String) - caminho destino para escrever excel
	 * @param autoSizeColumns (boolean) - Se True -> tamanho de todas as c células criadas será ajeitado de acordo com o tamanho de seus conteúdos. Afeta performance dependendo do tamanho das folhas do excel
	 * @param evaluateAllFormulas (boolean) - Se True -> atualiza os valores de todas as fórmulas das folhas. Afeta performance dependendo do tamanho das folhas do excel
	 * @throws IOException
	 */
	public void write(final String filePath, final boolean autoSizeColumns, final boolean evaluateAllFormulas) throws IOException {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(filePath));
			if(autoSizeColumns) {
				autoSizeColumns();
			}
			if(evaluateAllFormulas) {
				evaluateAllFormulas();
			}
			this._workbook.write(fos);
		} finally {
			if (null != fos) {
				fos.close();
			}
			close();
		}
	}

	/**
	 * Escreve workbook no buffer
	 * 
	 * @param ByteArrayOutputStream baos - buffer para escrever os bytes
	 * @param autoSizeColumns (boolean) - Se True -> tamanho de todas as c células criadas será ajeitado de acordo com o tamanho de seus conteúdos. Afeta performance dependendo do tamanho das folhas do excel
	 * @param evaluateAllFormulas (boolean) - Se True -> atualiza os valores de todas as fórmulas das folhas. Afeta performance dependendo do tamanho das folhas do excel
	 * @throws IOException
	 */
	public void write(final ByteArrayOutputStream baos, final boolean autoSizeColumns, final boolean evaluateAllFormulas) throws IOException {
		try {
			if(autoSizeColumns) {
				autoSizeColumns();
			}
			if(evaluateAllFormulas) {
				evaluateAllFormulas();
			}
			this._workbook.write(baos);
		} finally {
			close();
		}
	}

	/**
	 * Escreve workbook no buffer e retornar array de byte
	 * 
	 * @return byte[]
	 * @throws IOException
	 */
	public byte[] write() throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			autoSizeColumns().evaluateAllFormulas();
			this._workbook.write(baos);
		} finally {
			baos.close();
			close();
		}
		return baos.toByteArray();
	}
	
	/**
	 * Escreve no buffer ByteArrayOutputStream, fecha o buffer e workbook no finally
	 * @param autoSizeColumns (boolean) - Define se irá ajustar o tamanho de todas as colunas usadas em todas as folhas. Afeta performance gravemente dependendo do tamanho do arquivo
	 * @param evaluateFormulas (boolean) - Define se irá atualizar todas as fórmulas criadas e fará os cálculos
	 * @return byte[]
	 * @throws IOException
	 */
	public byte[] write(final boolean autoSizeColumns, final boolean evaluateFormulas) throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			if(autoSizeColumns) {
				autoSizeColumns();
			}
			if(evaluateFormulas) {
				evaluateAllFormulas();
			}
			this._workbook.write(baos);
		} finally {
			baos.close();
			close();
		}
		return baos.toByteArray();
	}
	
	/**
	 * Download do workbook como .xlsx
	 * @param fileName (String) - nome do arquivo. Caso não possua o .xlsx será inserido.
	 * @param autoSizeColumns (boolean) - Define se irá ajustar o tamanho de todas as colunas usadas em todas as folhas. Afeta performance gravemente dependendo do tamanho do arquivo
	 * @param evaluateFormulas (boolean) - Define se irá atualizar todas as fórmulas criadas e fará os cálculos
	 * @throws IOException 
	 */
	public void download(String fileName, final boolean autoSizeColumns, final boolean evaluateFormulas) throws IOException {
		final String xlsx = ".xlsx";
		if(!fileName.contains(xlsx)) {
			fileName = fileName + xlsx;
		}
		final ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ServletOutputStream sos = null;
		try {
			this.write(baos,autoSizeColumns,evaluateFormulas);
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
		this._workbook.close();
	}

	// SHEET

	/**
	 * Cria uma nova folha (XSSFSheet).
	 * 
	 * @param name (String) - nome da folha
	 */
	public POIUtil2 createSheet(final String name) {
		this._workbook.createSheet(name);
		return this;
	}
	
	/**
	 * Reordena folha
	 * @param sheet (XSSFSheet) - folha a ser reordenada
	 * @param pos (int) - posição (base 0) da folha
	 * @return POIUtil2
	 */
	public POIUtil2 moveSheet(final XSSFSheet sheet, final int pos) {
		this._workbook.setSheetOrder(sheet.getSheetName(), pos);
		return this;
	}
	
	/**
	 * Seta folha como a principal
	 * @param index (int) - posição (base 0) da folha
	 * @return POIUtil2
	 */
	public POIUtil2 activeSheet(final int index) {
		this._workbook.setActiveSheet(index);
		return this;
	}

	/**
	 * Retorna folha da posição informada.
	 * 
	 * @param index (int) - Posição da folha. Base 0
	 * @return XSSFSheet
	 */
	public XSSFSheet getSheetAt(final int index) {
		return this._workbook.getSheetAt(index);
	}
	
	/**
	 * Configura folha para padrão
	 * @param index (int) - Índice (base 0) da folha
	 * @return PoiUtil
	 */
	public POIUtil2 setActiveSheet(final int index) {
		this._workbook.setActiveSheet(index);
		return this;
	}

	/**
	 * Retorna folha pelo nome informado
	 * 
	 * @param name (String) - nome da folha
	 * @return XSSFSheet
	 */
	public XSSFSheet getSheet(final String name) {
		return this._workbook.getSheet(name);
	}
	
	/**
	 * Remove folha no index informado. Base 0
	 * @param index (int) index base 0
	 * @return POIUtil
	 */
	public POIUtil2 removeSheetAt(final int index) {
		this._workbook.removeSheetAt(index);
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
		setUltimaRegiaoAutalizada(sheet,regioes);
		return this;
	}
	

	/**
	 * Cria uma nova área mesclada
	 * 
	 * @param sheet    (XSSFSheet) - folha a ser alterada
	 * @param firstRow (int) - primeira linha para mesclar. Base 0
	 * @param lastRow  (int) - última linha para mesclar. Base 0
	 * @param firstCol (int) - primeira coluna para mesclar. Base 0
	 * @param lastCol  (int) - última coluna para mesclar. Base 0
	 * @return PoiUtil
	 */
	public POIUtil2 createMergedRegion(final XSSFSheet sheet, final int firstRow, final int lastRow, final int firstCol,
			final int lastCol) {

		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
		return this;
	}

	/**
	 * Cria várias áreas mescladas de vários pares de linhas Ex: firstLastRows =
	 * [[2,3],[4,5]] columns = [[2,3,8,6],[10,11,20,60]] Nas linhas 2 e 3 serão
	 * criadas as áreas mescladas (2 a 3) e (8 a 6) Nas linhas 4 e 5 as áreas
	 * mescladas (10,11) e (20,60)
	 * 
	 * @param sheet         (XSSFSheet) - folha a ser alterada
	 * @param firstLastRows (int[][]) - Matriz de linhas a serem alteradas. Cada
	 *                      array deve ser um par de ints indicando a linha inicio e
	 *                      fim: [[1,2],[3,4]]. Base 0
	 * @param columns       (int[][]) - Matriz de colunas a serem mescladas. A
	 *                      quantidade de colunas deve ser um número par. Base 0
	 * @throws IndexOutOfBoundsException - caso firstLastRows e columns sejam de
	 *                                   tamanhos diferentes
	 * @return PoiUtil
	 */
	public POIUtil2 createMergedRegions(final XSSFSheet sheet, final int[][] firstLastRows, final int[][] columns) {
		for (int index = 0; index < firstLastRows.length; index++) {
			final int[] rows = firstLastRows[index];
			final int[] cols = columns[index];
			for (int j = 0; j < cols.length; j++) {
				sheet.addMergedRegion(new CellRangeAddress(rows[0], rows[1], cols[j], cols[++j]));
			}

		}
		return this;
	}
	
	/**
	 * Cria painel estático
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row (int) - linha final (base 1) a ser congelada
	 * @param column (int) - coluna inicial a ser congelada
	 * @return
	 */
	public POIUtil2 createFreezPanel(final XSSFSheet sheet, final int row, final int column) {
		sheet.createFreezePane(column, row);
		return this;
	}

	// CELL

	/**
	 * Cria células a partir do array int[]
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param value (String) - valor em String
	 * @param row   (int) - número da linha onde serão criadas as células. Base 0
	 * @param cells (int[]) - números das células a serem criadas. Base 0
	 * @throws NullPointerException - caso linha não exista
	 * @return PoiUtil2
	 */
	public POIUtil2 createCell(final XSSFSheet sheet, final String value, final int row, final int cellNum) {
		final XSSFCell cell = this.getRow(sheet, row).createCell(cellNum);
		cell.setCellValue(value);
		setUltimaRegiaoAutalizadaLinhaColuna(sheet,row,cellNum);
		return this;
	}
	
	/**
	 * Cria célula, seta valor String e configura estilo 
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param estilo (String) - nome do estilo existente
	 * @param value (String) - valor em String
	 * @param row   (int) - número da linha onde serão criadas as células. Base 0
	 * @param cells (int[]) - números das células a serem criadas. Base 0
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @return PoiUtil2
	 */
	public POIUtil2 createCell(final XSSFSheet sheet, final String estilo, final String value, final int row, final int cellNum) {
		final XSSFCell cell = this.getRow(sheet, row).createCell(cellNum);
		cell.setCellStyle(getCellStyle(estilo));
		cell.setCellValue(value);
		setUltimaRegiaoAutalizadaLinhaColuna(sheet,row,cellNum);
		return this;
	}
	
	/**
	 * Cria célula, seta valor Date e configura estilo 
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param estilo (String) - nome do estilo existente
	 * @param value (Date) - valor em Date
	 * @param row   (int) - número da linha onde serão criadas as células. Base 0
	 * @param cells (int[]) - números das células a serem criadas. Base 0
	 * @throws NullPointerException - caso linha ou estilo não existam
	 * @return PoiUtil2
	 */
	public POIUtil2 createCell(final XSSFSheet sheet, final String estilo, final Date value, final int row, final int cellNum) {
		final XSSFCell cell = this.getRow(sheet, row).createCell(cellNum);
		cell.setCellStyle(getCellStyle(estilo));
		if(null == value) {
			cell.setCellValue("");
		}else {
			cell.setCellValue(value);
		}
		setUltimaRegiaoAutalizadaLinhaColuna(sheet,row,cellNum);
		return this;
	}
	
	/**
	 * Cria célula e seta valor int
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param value (int) - valor em int
	 * @param row   (int) - número da linha onde serão criadas as células. Base 0
	 * @param cells (int[]) - números das células a serem criadas. Base 0
	 * @throws NullPointerException - caso linha não exista
	 * @return PoiUtil2
	 */
	public POIUtil2 createCell(final XSSFSheet sheet, final int value, final int row, final int cellNum) {
		final XSSFCell cell = this.getRow(sheet, row).createCell(cellNum);
		cell.setCellValue(value);
		setUltimaRegiaoAutalizadaLinhaColuna(sheet,row,cellNum);
		return this;
	}
	
	/**
	 * Cria célula, seta valor int e configura estilo
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param styleName (String) - nome do estilo criado
	 * @param value (int) - valor em int
	 * @param row   (int) - número da linha onde serão criadas as células. Base 0
	 * @param cells (int[]) - números das células a serem criadas. Base 0
	 * @throws NullPointerException - caso linha não exista
	 * @return PoiUtil2
	 */
	public POIUtil2 createCell(final XSSFSheet sheet, final String styleName, final int value, final int row, final int cellNum) {
		final XSSFCell cell = this.getRow(sheet, row).createCell(cellNum);
		cell.setCellStyle(getCellStyle(styleName));
		cell.setCellValue(value);
		setUltimaRegiaoAutalizadaLinhaColuna(sheet,row,cellNum);
		return this;
	}
	
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
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
		return this;
	}
	
	
	/**
	 * Cria células
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row (int) - número da linha
	 * @param cols (int[]) - números das células a serem criadas
	 * @throws NullPointerException - caso linha não exista
	 * @return PoiUtil
	 */
	public POIUtil2 createCells(final XSSFSheet sheet, final int row, final int[] cols) throws Exception {
		final XSSFRow xssfRow = getRow(sheet, row);
		int index = 0;
		while(index < cols.length) {
			xssfRow.createCell(cols[index++]);
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
	public POIUtil2 createCellsComStyle(final XSSFSheet sheet, final String cellStyleName, final String... celulasRegioes) throws Exception {
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
						this.getRow(sheet, linhaInicial).createCell(coluna++).setCellStyle(this._cellStyles.get(cellStyleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]).setCellStyle(this._cellStyles.get(cellStyleName));
			}
		}
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
	public POIUtil2 createCellsComStyleValue(final XSSFSheet sheet, final String cellStyleName, final String cellValue, final String... celulasRegioes) throws Exception {
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
						cell.setCellStyle(this._cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this._cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
	public POIUtil2 createCellsComStyleValue(final XSSFSheet sheet, final String cellStyleName, final long cellValue, final String... celulasRegioes) throws Exception {
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
						cell.setCellStyle(this._cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this._cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
	public POIUtil2 createCellsComStyleValue(final XSSFSheet sheet, final String cellStyleName, final int cellValue, final String... celulasRegioes) throws Exception {
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
						cell.setCellStyle(this._cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this._cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
	public POIUtil2 createCellsComStyleValue(final XSSFSheet sheet, final String cellStyleName, final double cellValue, final String... celulasRegioes) throws Exception {
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
						cell.setCellStyle(this._cellStyles.get(cellStyleName));
						cell.setCellValue(cellValue);
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				final XSSFCell cell = this.getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]);
				cell.setCellStyle(this._cellStyles.get(cellStyleName));
				cell.setCellValue(cellValue);
			}
		}
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
		setUltimaRegiaoAutalizada(sheet,celula);
		return this;
	}
	
	/**
	 * Configura em string o valor da célula
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param value (String) - valor a ser configurado na célula
	 * @param row (int) - linha
	 * @param column (int) - coluna
	 * @throws NullPointerException - caso linha ou célula não existam ou valor seja null
	 * @throws Exception - caso célula informada seja inválida 
	 * @return PoiUtil
	 */
	public POIUtil2 setCellValue(final XSSFSheet sheet, final String value, final int row, final int column) throws Exception {
		sheet.getRow(row).getCell(column).setCellValue(value);
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
		setUltimaRegiaoAutalizada(sheet,celula);
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
		setUltimaRegiaoAutalizada(sheet,celula);
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
		setUltimaRegiaoAutalizada(sheet,celula);
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
	

	/**
	 * Retorna célula. Retorna nulo caso célula não exista
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param linha (int) - linha da célula
	 * @param coluna (int) - coluna da célula
	 * @throws NullPointerException - caso linha não exista
	 * @return XSSFCell
	 */
	public XSSFCell getCell(final XSSFSheet sheet, final int linha, final int coluna){
		return sheet.getRow(linha).getCell(coluna);
	}
	
	
	/**
	 * Recupera a localização de uma célula existente em String.
	 * @param sheet (XSSFSheet) - folha
	 * @param linha (int) - número da linha
	 * @param coluna (int) - número da coluna
	 * @return localização em String. Ex: "ABF31"
	 */
	public String getCellCreatedString(final XSSFSheet sheet, final int linha, final int coluna){
		return sheet.getRow(linha).getCell(coluna).getAddress().formatAsString();
	}
	
	public POIUtil2 editCellDataFormat(final XSSFCell cell, final String format) {
		cell.getCellStyle().setDataFormat(getDataFormat(format));
		return this;
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
		setUltimaRegiaoAutalizada(sheet,row);
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
		setUltimaRegiaoAutalizada(sheet,rows);
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
		this._cellStyles.put(cellStyleName, this._workbook.createCellStyle());
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
	 * Cria e configura XSSFCellStyle no HashMap<String,XSSFCellStyle>
	 * @param name (String) - key no hashmap
	 * @param font (String) - key da fonte no HashMap<String,XSSFFont>
	 * @param hAlign (HorizontalAlignment) - alinhameto horizontal texto
	 * @param vAlign (VerticalAlignment) - posialinhametoção vertical texto
	 * @param fill (FillPatternType) - padrão de preenchimento do background
	 * @param color (Color) - índice da cor de preenchimento
	 * @param borders - (BorderStyle...) - array de bordas -> Top, Right, Bottom, Left 
	 * @return
	 */
	public POIUtil2 createCellStyle(final String name, final String font,final boolean wrapText, final HorizontalAlignment hAlign, final VerticalAlignment vAlign,
			final FillPatternType fill, final Color color,final BorderStyle... borders) {
		
		createCellStyle(name).editCellStyleFont(name, font).editCellStyleWrapText(name, wrapText).editCellStyleAlignment(name, hAlign, vAlign)
		.editCellStyleFillPattern(name, fill).editCellStyleForegroundColor(name, color);
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
	 * Apaga o estilo criado. Não afeta as células criadas com o estilo passado.
	 * @param name (String) - nome do estilo
	 * @return POIUtil2
	 */
	public POIUtil2 removeCellStyle(final String name) {
		this._cellStyles.remove(name);
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
		return this._cellStyles.get(nome);
	}
	
	/**
	 * Altera o formato do estilo. Ex: "dd/MM/yyyy" -> célula será do tipo Data.
	 * @param cellStyleName (String) - nome chave do estilo no HashMap
	 * @param format (String) - formato desejado
	 * @return
	 */
	public POIUtil2 editCellStyleDataFormat(final String cellStyleName, final String format) {
		final XSSFCellStyle style = getCellStyle(cellStyleName);
		style.setDataFormat(getDataFormat(format));
		return this;
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
	 * Edita borda da célula
	 * 
	 * @param cell (XSSF) - célula com estilo
	 * @param borderTop     (BorderStyle) - borda topo. Passe nulo para não setar
	 * @param borderRight   (BorderStyle) - borda direita. Passe nulo para não setar
	 * @param borderBottom  (BorderStyle) - borda inferior. Passe nulo para não
	 *                      setar
	 * @param borderLeft    (BorderStyle) - borda esquerda. Passe nulo para não
	 *                      setar
	 * @throws NullPointerException - caso estilo não exista
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleBorder(final XSSFCell cell, final BorderStyle borderTop,
			final BorderStyle borderRight, final BorderStyle borderBottom, final BorderStyle borderLeft) {

		final XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle().clone();
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
		cell.setCellStyle(cellStyle);
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
	 * Edita o alinhamento horizontal e vertical da célula
	 * 
	 * @param cell       (XSSFCell) - célula com estilo
	 * @param horizontalAlignment (HorizontalAlignment) - alinhamento horizontal.
	 *                            Passe null para não setar
	 * @param verticalAlignment   (VerticalAlignment) - alinhamento vertical. Passe
	 *                            null para não setar
	 * @throws NullPointerException - caso estilo não exista
	 * @return
	 */
	public POIUtil2 editCellStyleAlignment(final XSSFCell cell, final HorizontalAlignment horizontalAlignment,
			final VerticalAlignment verticalAlignment) {

		final XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle().clone();
		if (null != horizontalAlignment) {
			style.setAlignment(horizontalAlignment);
		}
		if (null != verticalAlignment) {
			style.setVerticalAlignment(verticalAlignment);
		}
		cell.setCellStyle(style);
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
	 * Configura padrão de preenchimento do estilo da célula
	 * 
	 * @param cell (XSSFCell) - célula a ser alterada
	 * @param fillPattern   (FillPatternType) - padrão de preenchimento
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleFillPattern(final XSSFCell cell, final FillPatternType fillPattern) {
		final XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle().clone();
		style.setFillPattern(fillPattern);
		cell.setCellStyle(style);
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
	 * Configura cor de fundo do estilo
	 * @param cellStyleName (String) - nome chave da fonte no HashMap
	 * @param color (Color) - cor de fundo
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleForegroundColor(final String cellStyleName, final Color color) {
		if(null != color) {
			this.getCellStyle(cellStyleName).setFillForegroundColor(new XSSFColor(color));			
		}
		return this;
	}
	
	/**
	 * Configura cor de fundo do etilo da célula
	 * @param cell (XSSFCell) - célula a ser alterada
	 * @param color (Color) - cor de fundo
	 * @return PoiUtil
	 */
	public POIUtil2 editCellStyleForegroundColor(final XSSFCell cell, final Color color) {
		final XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle().clone();
		style.setFillForegroundColor(new XSSFColor(color));
		cell.setCellStyle(style);
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
		setUltimaRegiaoAutalizada(sheet,celulasRegioes);
		return this;
	}
	
	/**
	 * Cria linha e célula
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row (int) - número da linha a ser criada. (base 0)
	 * @param col (int) - número da coluna a ser criada. (base 0)
	 * @return PoiUtil
	 */
	public POIUtil2 createRowCell(final XSSFSheet sheet, final int row, final int col) throws Exception{
		sheet.createRow(row).createCell(col);
		return this;
	}
	
	
	/**
	 * Cria linha e células
	 * 
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row (int) - número da linha a ser criada. (base 0)
	 * @param cols (int[]) - array de colunas a serem criadas. (base 0)
	 * @return PoiUtil
	 */
	public POIUtil2 createRowCells(final XSSFSheet sheet, final int row, final int[] cols) throws Exception{
		final XSSFRow xssfRow = sheet.createRow(row);
		int index = 0;
		while(index < cols.length) {
			xssfRow.createCell(cols[index++]);
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
						row.createCell(coluna++).setCellStyle(this._cellStyles.get(styleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.createRow(sheet, linhaColuna[0]).getRow(sheet, linhaColuna[0]).createCell(linhaColuna[1]).setCellStyle(this._cellStyles.get(styleName));
			}
			setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
						this.getRow(sheet, linhaInicial).getCell(coluna).setCellStyle(this._cellStyles.get(cellStyleName));
					}
					linhaInicial++;
				}
			}else {
				final int[] linhaColuna = getRowCol(celulasRegioes[i]);
				this.getRow(sheet, linhaColuna[0]).getCell(linhaColuna[1]).setCellStyle(this._cellStyles.get(cellStyleName));
			}
			setUltimaRegiaoAutalizada(sheet,celulasRegioes);
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
		this._fonts.put(name, this._workbook.createFont());
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
	 * Configura negrito da fonte da célula
	 * 
	 * @param cell (XSSFCell) - célula com estilo a ser configurado
	 * @param bold (boolean) - se é negrito ou não
	 * @return PoiUtil
	 */
	public POIUtil2 editFontBold(final XSSFCell cell, final boolean bold) {
		final XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle().clone();
		final XSSFFont font = copyFont(style.getFont());
		font.setBold(bold);
		style.setFont(font);
		cell.setCellStyle(style);
		return this;
	}
	
	/**
	 * Configura tamanho da fonte da célula
	 * 
	 * @param cell (XSSFCell) - célula com estilo a ser configurado
	 * @param size (double) - altura da fonte em double
	 * @return PoiUtil
	 */
	public POIUtil2 editFontSize(final XSSFCell cell, final double size) {
		final XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle().clone();
		final XSSFFont font = copyFont(style.getFont());
		font.setFontHeight(size);
		style.setFont(font);
		cell.setCellStyle(style);
		return this;
	}
	
	/**
	 * Configura tamanho da fonte da célula
	 * @param sheet (XSSFSheet) - folha a ser alterada
	 * @param row (int) - número da linha
	 * @param col (int) - número da coluna
	 * @param size (double) - altura da fonte em double
	 * @return
	 */
	public POIUtil2 editFontSize(final XSSFSheet sheet, final int row, final int col, final double size) {
		final XSSFCell cell = sheet.getRow(row).getCell(col);
		return editFontSize(cell, size);
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
		return this._fonts.get(name);
	}
	
	public XSSFFont copyFont(final XSSFFont font) {
		final XSSFFont newFont = this._workbook.createFont();
		newFont.setFamily(font.getFamily());
		newFont.setFontName(font.getFontName());
		newFont.setFontHeight(font.getFontHeight());
		newFont.setColor(font.getColor());
		newFont.setBold(font.getBold());
		return newFont;
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
		setUltimaRegiaoAutalizada(sheet,regiao);
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
			final CreationHelper helper = this._workbook.getCreationHelper();
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
			final int pictureIndex = this._workbook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
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
		setUltimaRegiaoAutalizada(sheet,regiao);
		final int[][] linhasColunas = getRowsCols(regiao);
			final CreationHelper helper = this._workbook.getCreationHelper();
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
			final int pictureIndex = this._workbook.addPicture(image, Workbook.PICTURE_TYPE_PNG);
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
	
	
	/**
	 * Configura em todas as folhas a largura de todas as colunas para "auto". Afeta performance gravemente dependendo do tamanho do arquivo.
	 * @return PoiUtil
	 */
	public POIUtil2 autoSizeColumns() {
		for(int i = 0; i < this._workbook.getNumberOfSheets(); i++) {
			autoSizeColumns(this.getSheetAt(i));
		}
		return this;
	}
	
	/**
	 * Ajusta o tamanho de todas as colunas preenchidas na folha. Afeta performance gravemente dependendo do tamanho da folha.
	 * @param sheet (XSSFSheet) - Folha para ser ajustada
	 * @return
	 */
	public POIUtil2 autoSizeColumns(final XSSFSheet sheet) {
		int lastCellNum = 0;
		int i = 0;
		XSSFRow row = null;
		while(i < sheet.getLastRowNum()) {
			row = sheet.getRow(i++);
			if(null == row) {
				break;
			}
			final int lastCellRow = row.getLastCellNum();
			lastCellNum = lastCellRow > lastCellNum ? lastCellRow : lastCellNum;
		}
		i = 0;
		while(i < lastCellNum) {
			sheet.autoSizeColumn(i++);
		}
		return this;
	}
	
	//FÓRMULAS

	/**
	 * Atualiza todas as fórmulas criadas nas células. É chamado automaticamente nos métodos write().
	 * @return PoiUtil
	 */
	public POIUtil2 evaluateAllFormulas() {
		XSSFFormulaEvaluator.evaluateAllFormulaCells(this._workbook);
		return this;
	}

	/**
	 * Cria fórmula de soma para a célula. ATENÇÃO: o tipo da célula será alterado para CELL_TYPE_FORMULA
	 * @param cell (XSSFCell) - célula a ser alterada
	 * @param formula (String) - fórmula String (em inglês). Ex: SUM(A1:B1).
	 * @return POIUtil2
	 */
	public POIUtil2 createCellFormula(final XSSFCell cell, final String formula) {
		cell.setCellType(CellType.FORMULA);
		cell.setCellFormula(formula);
		return this;
	}
	
	/**
	 * Cria fórmula de SOMA para a célula. ATENÇÃO: o tipo da célula será alterado para CELL_TYPE_FORMULA
	 * @param cell (XSSFCell) - célula a ser alterada
	 * @param formula (String) - região a ser somada. Ex: A1:B1
	 * @return POIUtil2
	 */
	public POIUtil2 sum(final XSSFCell cell, final String regiao) {
		return createCellFormula(cell, "SUM("+regiao+")");
	}
	
	
	/**
	 * Cria fórmula de CONT.SES para a célula. ATENÇÃO: o tipo da célula será alterado para CELL_TYPE_FORMULA
	 * @param cell (XSSFCell) - célula a ser alterada
	 * @param regiaoCriterio (String...) - array String contendo a região e o critério da função CONT.SES. Deve ser passado em pares. Ex: "A1:B1","VERDADEIRO","A3","FALSO"...
	 * @return POIUtil2
	 */
	public POIUtil2 countIfs(final XSSFCell cell,final String... regiaoCriterio) {
		final StringBuilder contSe = new StringBuilder(32);
		final String virgula = ",";
		contSe.append("COUNTIFS(");
		for(int i = 0; i < regiaoCriterio.length;) {
			contSe.append(regiaoCriterio[i] + virgula + regiaoCriterio[i + 1]);
			i += 2;
			if(i < regiaoCriterio.length) {
				contSe.append(virgula);
			}
		}
		contSe.append(")");
		return createCellFormula(cell,contSe.toString());
	}
	
	/**
	 * Retorna fórmula DATA(ano,mes,dia)
	 * @param cellRow (String) - localização da célula em String
	 * @param month (int) - mês desejado (base 1)
	 * @param day (int) - dia desejado (base 1)
	 * @return
	 */
	public String date(final String cellRow, final int month, final int day) {
		return "DATE("+cellRow+","+month+","+day+")";
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
	 * Cria formato e retorna seu valor em short
	 * @param format (String) - formato desejado. Ex: "dd/MM/yyyy"
	 * @return formato (short)
	 */
	public short getDataFormat(final String format) {
		return this._workbook.getCreationHelper().createDataFormat().getFormat(format);
	}
	
	/**
	 * Concatena duas colunas para formar uma região. Ex: concat("A1","B1") = "A1:B1"
	 * @param col1 (String) - coluna início
	 * @param col2 (String) - coluna final
	 * @throws NullPointerException - caso coluna seja null
	 * @return (String) - coluna início : coluna final
	 */
	public String concat(final String col1, final String col2) {
		return col1 + ":" + col2;
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
	 * @param cell (String) Ex: "A1"
	 * @return array com linha(0) e coluna(1) (int[]) 
	 */
	private int[] getRowCol(final String cell) {
		final CellReference cellRef = new CellReference(cell);
		return new int[] {cellRef.getRow(),cellRef.getCol()};
	}
	
	
	/**
	 * Recupera os números das linhas (iniciais e finais) e colunas (iniciais e finais) de determinada regiao.
	 * @param region (String) Ex: "A1:D1"
	 * @return matriz com dois arrays de linha e coluna
	 */
	private int[][] getRowsCols(final String region){
		final String[] regiaoArray = region.split(":");
		final int[][] retorno = new int[2][1];
		retorno[0] = getRowCol(regiaoArray[0]);
		retorno[1] = getRowCol(regiaoArray[1]);
		return retorno;
	}
	
	/**
	 * Recupera a localização da célula em String.
	 * @param row (int) - número da linha
	 * @param col (int) - número da coluna
	 * @return localização em String. Ex: "ABF31"
	 */
	public String getCellString(final int row, final int col) {
		final CellReference cellReference = new CellReference(row, col);
		return cellReference.formatAsString();
	}
	
	/**
	 * Recupera a localização da célula em String.
	 * @param linha (int) - número da linha
	 * @param coluna (int) - número da coluna
	 * @return localização em String. Ex: "ABF31"
	 */
	public String getLockedCellString(final int linha, final int coluna, final boolean lockRow, final boolean lockCell) {
		final String cellString = getCellString(linha, coluna);
		final String cell = cellString.replaceAll(REGEX_APENAS_LETRAS, "");
		final String row = cellString.replaceAll(REGEX_APENAS_NUMEROS, "");
		final String dollarSign = "$";
		return (lockCell ? dollarSign + cell : cell) + (lockRow ? dollarSign + row : row);
	}
	
	/**
	 * Seta última célula/região que foi criada/editada
	 * @param celulasRegioes
	 */
	private void setUltimaRegiaoAutalizada(final XSSFSheet sheet, final String... celulasRegioes) {
		if(this.auditoriaRegiaoAtualizada) {
			this._ultimaRegiaoAtualizada = null;
			if(celulasRegioes.length > 0 ) {
				this._ultimaRegiaoAtualizada = sheet.getSheetName() + "!"+celulasRegioes[celulasRegioes.length -1 ];
			}
		}
	}
	
	/**
	 * Seta última linha criada
	 * @param linhas
	 */
	private void setUltimaRegiaoAutalizada(final XSSFSheet sheet,final int... linhas) {
		if(this.auditoriaRegiaoAtualizada) {
			this._ultimaRegiaoAtualizada = null;
			if(linhas.length > 0) {
				this._ultimaRegiaoAtualizada = sheet.getSheetName() + "!Linha " + linhas[linhas.length -1 ] + " criada";			
			}
		}
	}
	
	/**
	 * Seta última célula/região que foi criada/editada
	 * @param linhas
	 */
	public void setUltimaRegiaoAutalizadaLinhaColuna(final XSSFSheet sheet,final int linha, final int coluna) {
		if(this.auditoriaRegiaoAtualizada) {
			this._ultimaRegiaoAtualizada = sheet.getSheetName() + "!" + getCellCreatedString(sheet, linha, coluna);			
		}
	}

	/**
	 * Configura se a variável _ultimaRegiaoAtualizada deve ser atualizada. Deve ser usada somente em desenvolvimento, pois afeta performance. 
	 * @param auditoriaRegiaoAtualizada (boolean)
	 */
	public void setAuditoriaRegiaoAtualizada(boolean auditoriaRegiaoAtualizada) {
		this.auditoriaRegiaoAtualizada = auditoriaRegiaoAtualizada;
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
