import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.WindowConstants;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OdeonConverter extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public OdeonConverter() {
		JFileChooser filSelecionarArquivo = new JFileChooser();
        JButton btnSelecionarArquivo = new JButton("Escolher Arquivo");
        JTextField txtNomeArquivo = new JTextField("", 40);
        JButton btnProcessar = new JButton("Processar");
 
        GroupLayout layout = new GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        setResizable(false);
        layout.setAutoCreateGaps(true);
        layout.setAutoCreateContainerGaps(true);
 
        layout.setHorizontalGroup(layout.createSequentialGroup()
            .addComponent(btnSelecionarArquivo)
            .addGroup(layout.createParallelGroup(Alignment.LEADING)
                .addComponent(txtNomeArquivo))
            .addGroup(layout.createParallelGroup(Alignment.LEADING)
                .addComponent(btnProcessar))
        );
        
        layout.linkSize(SwingConstants.HORIZONTAL, btnProcessar);
 
        layout.setVerticalGroup(layout.createSequentialGroup()
            .addGroup(layout.createParallelGroup(Alignment.BASELINE)
                .addComponent(btnSelecionarArquivo)
                .addComponent(txtNomeArquivo)
                .addComponent(btnProcessar)));
 
        setTitle("Odeon Converter");
        pack();
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

		btnProcessar.addActionListener(new ActionListener() { 
			public void actionPerformed(ActionEvent e) {
				try {
					lerArquivo(txtNomeArquivo.getText());
					JOptionPane.showMessageDialog(null, "Arquivo gerado com sucesso!", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(null, e1.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
					e1.printStackTrace();
				}
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

	private void lerArquivo(String nomeArquivo) throws Exception {
		if (nomeArquivo == null || nomeArquivo.trim().equals("")) {
			throw new Exception("Informe o arquivo a ser processado.");
		}
		
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
        exportarArquivo(nomeArquivo, linhas);
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

	private void exportarArquivo(String arquivo, List<String> linhas) throws Exception {
		Workbook wb = new XSSFWorkbook();
		wb.createSheet("Planilha1");
		Sheet planilha = wb.getSheetAt(0);
		
		int l = -1;
		
		for (String linha : linhas) {
			planilha.createRow(++l);
			int c = 0;

			for (String celula : linha.split(";")) {
				if (NumberUtils.isParsable(celula)) {
					planilha.getRow(l).createCell(c++).setCellValue(Float.parseFloat(celula));
				} else {
					planilha.getRow(l).createCell(c++).setCellValue(celula);
				}
			}
		}

		File fileOutput = null;
		String arquivoSaida = null;
		int seqArq = 0;

		do {
			arquivoSaida = arquivo.substring(0, arquivo.lastIndexOf('.')) + (seqArq == 0 ? "" : " (" + seqArq + ")") + ".xlsx";
			seqArq ++;
			fileOutput = new File(arquivoSaida);
		} while (fileOutput.exists());
		
		FileOutputStream fos = new FileOutputStream(fileOutput);
		wb.write(fos);
		wb.close();
		fos.flush();
		fos.close();
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
