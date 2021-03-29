import java.awt.Container;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.imageio.ImageIO;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableModel;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OdeonConverter extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public OdeonConverter() {
        JTextField txtTituloPlanilha = new JTextField("", 40) {{ setPreferredSize(new Dimension(10, 30)); }};
        JButton btnSelecionarArquivo = new JButton("Escolher Arquivo", obterIcone("open.png")) {{ setIconTextGap(10); setHorizontalAlignment(SwingConstants.RIGHT); }};
		JFileChooser filSelecionarArquivo = new JFileChooser() {
			@Override
			public void updateUI() {
				putClientProperty("FileChooser.useShellFolder", Boolean.FALSE);
				super.updateUI();
			}
		};
        JTextField txtNomeArquivo = new JTextField("", 40) {{ setPreferredSize(new Dimension(10, 30)); }};
        JButton btnAdicionarArquivo = new JButton("Adicionar Arquivo", obterIcone("add-file.png")) {{ setIconTextGap(10); setHorizontalAlignment(SwingConstants.LEFT); }};
        JButton btnLimparLista = new JButton("Limpar Lista", obterIcone("delete-file.png")) {{ setIconTextGap(10); }};
        JButton btnProcessar = new JButton("Processar", obterIcone("engineering.png")) {{ setIconTextGap(10); }};
        JPanel pnlBotoes = new JPanel() {{ add(btnLimparLista); add(btnProcessar); }};
        JTable tblArquivos = obterTabela(new String[] { "Título da Planilha", "Nome do Arquivo" }, new Integer[] { 300, 300 });
        JScrollPane scrTabela = new JScrollPane(tblArquivos);
        scrTabela.setSize(new Dimension(800, 400));
        tblArquivos.setPreferredSize(new Dimension(600, 400));
        tblArquivos.setPreferredScrollableViewportSize(tblArquivos.getPreferredSize());
        tblArquivos.setFillsViewportHeight(true);

        JPanel painel = new JPanel();
        painel.setLayout(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(5, 5, 5, 5);

        c.gridy = 0;
        c.gridx = 0;
        c.anchor = GridBagConstraints.EAST;
        painel.add(new JLabel("Título da Planilha"), c);

        c.fill = GridBagConstraints.BOTH;
        c.gridx = 1;
        c.gridwidth = 2;
        c.anchor = GridBagConstraints.CENTER;
        painel.add(txtTituloPlanilha, c);
        
        c.gridy = 1;
        c.gridx = 0;
        c.gridwidth = 1;
        painel.add(btnSelecionarArquivo, c);

        c.gridx = 1;
        c.gridwidth = 2;
        painel.add(txtNomeArquivo, c);

        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridy = 2;
        c.gridx = 0;
        c.gridwidth = 1;
        painel.add(btnAdicionarArquivo, c);

        c.gridx = 2;
        c.fill = GridBagConstraints.NONE;
        c.anchor = GridBagConstraints.EAST;
        painel.add(pnlBotoes, c);

        c.gridy = 3;
        c.gridx = 0;
        c.gridwidth = 3;
        c.fill = GridBagConstraints.BOTH;
        painel.add(scrTabela, c);

		Container painelConteudo = getContentPane();
		painelConteudo.add(painel);
 
        setTitle("Odeon Converter");
        pack();
        setResizable(false);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		btnLimparLista.addActionListener(new ActionListener() { 
			public void actionPerformed(ActionEvent e) {
				DefaultTableModel modelo = (DefaultTableModel) tblArquivos.getModel();
				modelo.setRowCount(0);
			} 
		}); 

		btnProcessar.addActionListener(new ActionListener() { 
			public void actionPerformed(ActionEvent e) {
				try {
					lerArquivo(tblArquivos);
					JOptionPane.showMessageDialog(null, "Arquivo gerado com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(null, e1.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
					e1.printStackTrace();
				}
			} 
		}); 

		btnAdicionarArquivo.addActionListener(new ActionListener() { 
			public void actionPerformed(ActionEvent e) {
				try {
					adicionarArquivo(tblArquivos, txtTituloPlanilha, txtNomeArquivo);
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(null, e1.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
					e1.printStackTrace();
				}
			}

			private void adicionarArquivo(JTable tabela, JTextField txtTituloPlanilha, JTextField txtNomeArquivo) throws Exception {
				if (txtTituloPlanilha.getText().trim().equals("")) {
					throw new Exception("O título da planilha precisa estar preenchido.");
				}

				if (txtNomeArquivo.getText().trim().equals("")) {
					throw new Exception("O arquivo a ser processado deve ser selecionado.");
				}

			    DefaultTableModel modelo = (DefaultTableModel) tabela.getModel();
		        modelo.addRow(new Object[] { txtTituloPlanilha.getText(), txtNomeArquivo.getText() });
		        txtTituloPlanilha.setText("");
		        txtNomeArquivo.setText("");
		    }
		}); 

		btnSelecionarArquivo.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				filSelecionarArquivo.setAcceptAllFileFilterUsed(false);
				int retorno = filSelecionarArquivo.showOpenDialog(OdeonConverter.this);
				if (retorno == JFileChooser.APPROVE_OPTION) {
					if (filSelecionarArquivo.getSelectedFile().exists()) {
						txtNomeArquivo.setText(filSelecionarArquivo.getSelectedFile().getAbsolutePath());
					}
				}
			}
		});
	}

	private JTable obterTabela(String[] nomesColunas, Integer[] larguras) {
	    DefaultTableModel modelo = new DefaultTableModel(nomesColunas, 0) {
            @SuppressWarnings({ "unchecked", "rawtypes" })
			public Class getColumnClass(int column)
            {
                return getValueAt(0, column).getClass();
            }
	    };
		JTable retorno = new JTable(modelo);
	    retorno.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
		for (int i = 0; i < larguras.length; i++) {
			retorno.getColumnModel().getColumn(i).setPreferredWidth(larguras[i]);
		}
		
	    return retorno;
	}

	private void lerArquivo(JTable tabelaArquivos) throws Exception {
		DefaultTableModel modelo = (DefaultTableModel) tabelaArquivos.getModel();
		if (modelo.getRowCount() == 0) {
			throw new Exception("Informe o arquivo a ser processado.");
		}

		// criar planilha
		Workbook wb = new XSSFWorkbook();
		Sheet planilha = wb.createSheet("Planilha1");
		planilha.setDefaultColumnWidth(10);

		Map<String, CellStyle> mapaEstilos = new LinkedHashMap<String, CellStyle>();
		mapaEstilos.put("titulo", estiloTitulo(wb));
		mapaEstilos.put("subtitulo", estiloSubTitulo(wb));
		mapaEstilos.put("tituloresumo", estiloTituloResumo(wb));
		mapaEstilos.put("cabecalho", estiloCabecalho(wb));
		mapaEstilos.put("dado0dec", estiloDado(wb, 0));
		mapaEstilos.put("dado1dec", estiloDado(wb, 1));
		mapaEstilos.put("dado2dec", estiloDado(wb, 2));
		mapaEstilos.put("dado3dec", estiloDado(wb, 3));

		int colBase = 0;

		for (int l = 0; l < modelo.getRowCount(); l++) {
			String titulo = modelo.getValueAt(l, 0).toString();
			String nomeArquivo = modelo.getValueAt(l, 1).toString();

			if (!(new File(nomeArquivo).exists())) {
				throw new Exception("O arquivo informado não foi encontrado.");
			}

	        String linha;
	        BufferedReader in = new BufferedReader(new FileReader(nomeArquivo));
	        List<String> linhas = new ArrayList<String>();

	        do {
	        	linha = in.readLine();

	        	if (linha != null) {
	        		linhas.add(retornaLinhaConvertida(linha));
	        	}
	        } while (linha != null);

	        in.close();

	        // exporta o arquivo em Excel
	        exportarArquivo(planilha, titulo, linhas, colBase, mapaEstilos);
	        colBase += 10;
		}
		
		FileOutputStream fos = new FileOutputStream(System.getProperty("user.home") + File.separator + "Documents" + File.separator + "OdeonConverter " + formatarData(new Date(), "yyyy-MM-dd HH_mm_ss") + ".xlsx");
		wb.write(fos);
		wb.close();
		fos.flush();
		fos.close();
	}

	private String retornaLinhaConvertida(String linha) {
		String retorno = linha;

		// título da tabela
		if (linha.toLowerCase().startsWith("receiver")) {
			retorno = linha.substring(0, linha.indexOf("(")).trim();
			retorno += " - ";
			retorno += linha.substring(linha.indexOf("(")).trim();
		}

		// cabeçalho
		if (linha.length() > 15 && (
				linha.toLowerCase().startsWith("band (hz)")
				|| linha.toLowerCase().startsWith("edt       (s)")
				|| linha.toLowerCase().startsWith("t30       (s)")
				|| linha.toLowerCase().startsWith("spl       (db)")
				|| linha.toLowerCase().startsWith("c80       (db)")
				|| linha.toLowerCase().startsWith("d50      ")
				|| linha.toLowerCase().startsWith("ts        (ms)")
				|| linha.toLowerCase().startsWith("lf80     ")
				|| linha.toLowerCase().startsWith("minimum")
				|| linha.toLowerCase().startsWith("maximum")
				|| linha.toLowerCase().startsWith("average")
				)) {
			
			int t = linha.indexOf(")") + 1;
			if (t == 0) t = 7;

			retorno = linha.substring(0, t).trim().replaceAll("( )+", " ");
			retorno += ";";
			retorno += linha.substring(t).trim().replaceAll("( )+", ";");
		}

		// moldura
		if (linha.startsWith("_____")) {
			retorno = "";
		}

		return retorno;
	}

	public static String formatarData(Date data, String formato) {
		SimpleDateFormat f = new SimpleDateFormat(formato, new Locale("pt", "BR"));
		return f.format(data);
	}

	private void exportarArquivo(Sheet planilha, String titulo, List<String> linhas, int colBase, Map<String, CellStyle> mapaEstilos) throws Exception {
		int l = -1;
		
		if (++l > planilha.getLastRowNum() || !planilha.rowIterator().hasNext()) {
			planilha.createRow(l);
		}

		CellRangeAddress areaMescladaTitulo = new CellRangeAddress(l, l, colBase, colBase + 8);
		planilha.getRow(l).createCell(colBase).setCellValue(titulo);
		planilha.getRow(l).getCell(colBase).setCellStyle(mapaEstilos.get("titulo"));
		planilha.addMergedRegion(areaMescladaTitulo);
		adicionarBordaAreaMesclada(areaMescladaTitulo, planilha);
		String tipoDado = null;

		for (String linha : linhas) {
			if (++l > planilha.getLastRowNum()) {
				planilha.createRow(l);
			}

			int c = colBase;
			String[] celulas = linha.split(";");
			boolean isCabecalho = false;
			boolean isDado = false;

			for (String celula : celulas) {
				if (!isCabecalho && !isDado) {
					if (celulas.length > 1) {
						if (celulas[0].startsWith("Band")) {
							isCabecalho = true;
						} else if (celulas[0].startsWith("EDT")
								|| celulas[0].startsWith("T30")
								|| celulas[0].startsWith("SPL")
								|| celulas[0].startsWith("C80")
								|| celulas[0].startsWith("D50")
								|| celulas[0].startsWith("Ts ")
								|| celulas[0].startsWith("LF80")
								|| celulas[0].startsWith("Min")
								|| celulas[0].startsWith("Max")
								|| celulas[0].startsWith("Ave")
								) 
						{
							isDado = true;
						}
					}
				}

				if (NumberUtils.isParsable(celula)) {
					planilha.getRow(l).createCell(c++).setCellValue(Double.parseDouble(celula));
				} else {
					planilha.getRow(l).createCell(c++).setCellValue(celula);
				}

				if (isCabecalho) {
					planilha.getRow(l).getCell(c - 1).setCellStyle(mapaEstilos.get("cabecalho"));
				}

				if (isDado) {
					planilha.getRow(l).getCell(c - 1).setCellStyle(obterEstiloDado(tipoDado == null ? celulas[0] : tipoDado, mapaEstilos));
				}
			}

			if (celulas.length == 1 && !celulas[0].trim().equals("")) {
				CellRangeAddress areaMesclada = new CellRangeAddress(l, l, colBase, colBase + 8);
				planilha.addMergedRegion(areaMesclada);

				if (celulas[0].startsWith("Receiver")) {
					planilha.getRow(l).getCell(colBase).setCellStyle(mapaEstilos.get("subtitulo"));
					tipoDado = null;
				} else if (celulas[0].length() <= 10
						&& (celulas[0].startsWith("EDT")
						|| celulas[0].startsWith("T30")
						|| celulas[0].startsWith("SPL")
						|| celulas[0].startsWith("C80")
						|| celulas[0].startsWith("D50")
						|| celulas[0].startsWith("Ts ")
						|| celulas[0].startsWith("LF80")
						)) {
					tipoDado = celulas[0];
					planilha.getRow(l).getCell(colBase).setCellStyle(mapaEstilos.get("tituloresumo"));
					adicionarBordaAreaMesclada(areaMesclada, planilha);
				}
			}
		}
	}
	
	private void adicionarBordaAreaMesclada(CellRangeAddress areaMesclada, Sheet planilha) {
		RegionUtil.setBorderBottom(BorderStyle.THIN, areaMesclada, planilha);
		RegionUtil.setBorderRight(BorderStyle.THIN, areaMesclada, planilha);
		RegionUtil.setBorderLeft(BorderStyle.THIN, areaMesclada, planilha);
		RegionUtil.setBorderTop(BorderStyle.THIN, areaMesclada, planilha);
	}

	private CellStyle estiloTitulo(Workbook wb) {
		XSSFCellStyle estilo = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor corFundo = new XSSFColor(new java.awt.Color(255, 212, 40));
		Font fonteNegrito = wb.createFont();
		fonteNegrito.setBold(true);
		estilo.setFillForegroundColor(corFundo);
		estilo.setFont(fonteNegrito);
		estilo.setAlignment(HorizontalAlignment.CENTER);
		estilo.setWrapText(true);
		estilo.setVerticalAlignment(VerticalAlignment.CENTER);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estilo.setBorderBottom(BorderStyle.THIN);
		estilo.setBorderLeft(BorderStyle.THIN);
		estilo.setBorderRight(BorderStyle.THIN);
		estilo.setBorderTop(BorderStyle.THIN);

		return estilo;
	}

	private CellStyle estiloSubTitulo(Workbook wb) {
		XSSFCellStyle estilo = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor corFundo = new XSSFColor(new java.awt.Color(180, 199, 220));
		Font fonteNegrito = wb.createFont();
		fonteNegrito.setBold(true);
		estilo.setFillForegroundColor(corFundo);
		estilo.setFont(fonteNegrito);
		estilo.setAlignment(HorizontalAlignment.LEFT);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		return estilo;
	}

	private CellStyle estiloTituloResumo(Workbook wb) {
		XSSFCellStyle estilo = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor corFundo = new XSSFColor(new java.awt.Color(63, 175, 70));
		Font fonteNegrito = wb.createFont();
		fonteNegrito.setBold(true);
		estilo.setFillForegroundColor(corFundo);
		estilo.setFont(fonteNegrito);
		estilo.setAlignment(HorizontalAlignment.LEFT);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estilo.setBorderBottom(BorderStyle.THIN);
		estilo.setBorderLeft(BorderStyle.THIN);
		estilo.setBorderRight(BorderStyle.THIN);
		estilo.setBorderTop(BorderStyle.THIN);

		return estilo;
	}

	private CellStyle estiloCabecalho(Workbook wb) {
		XSSFCellStyle estilo = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor corFundo = new XSSFColor(new java.awt.Color(221, 221, 221));
		estilo.setFillForegroundColor(corFundo);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estilo.setBorderBottom(BorderStyle.THIN);
		estilo.setBorderLeft(BorderStyle.THIN);
		estilo.setBorderRight(BorderStyle.THIN);
		estilo.setBorderTop(BorderStyle.THIN);

		return estilo;
	}

	private CellStyle estiloDado(Workbook wb, int decimais) {
		XSSFCellStyle estilo = (XSSFCellStyle) wb.createCellStyle();
		DataFormat formato = wb.createDataFormat();
		estilo.setDataFormat(formato.getFormat("#,##0" + (decimais > 0 ? "." + String.join("", Collections.nCopies(decimais, "0")) : "")));
		estilo.setBorderBottom(BorderStyle.THIN);
		estilo.setBorderLeft(BorderStyle.THIN);
		estilo.setBorderRight(BorderStyle.THIN);
		estilo.setBorderTop(BorderStyle.THIN);

		return estilo;
	}

	private CellStyle obterEstiloDado(String dado, Map<String, CellStyle> mapaEstilos) {
		if (dado.startsWith("Ts")) {
			return mapaEstilos.get("dado0dec");
		} else if (dado.startsWith("SPL") || dado.startsWith("C80")) {
			return mapaEstilos.get("dado1dec");
		} else if (dado.startsWith("EDT") || dado.startsWith("T30") || dado.startsWith("D50")) {
			return mapaEstilos.get("dado2dec");
		} else if (dado.startsWith("LF80")) {
			return mapaEstilos.get("dado3dec");
		}

		return null;
	}
	
	private Icon obterIcone(String arquivo) {
		return obterIcone(arquivo, 20);
	}

	private Icon obterIcone(String arquivo, int tamanho) {
        Image img = null;
		try {
			img = ImageIO.read(getClass().getResource(arquivo));
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		Image newimg = img.getScaledInstance(tamanho, tamanho, Image.SCALE_SMOOTH ) ;  
		return new ImageIcon(newimg);
	}
	
	public static void main(String[] args) {
	    try { 
	        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); 
	    } catch(Exception ignored) {
	    	ignored.printStackTrace();
	    }
		OdeonConverter app = new OdeonConverter();
		app.setDefaultCloseOperation(DISPOSE_ON_CLOSE);
		app.setVisible(true);
	}
}
