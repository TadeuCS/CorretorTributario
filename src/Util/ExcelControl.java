/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Util;

import Model.Produto;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.MergedCellsRecord;

public class ExcelControl {

    private Label titulo;
    private Label colunaA;
    private Label colunaB;
    private Label colunaC;
    private Label colunaD;
    private Label colunaE;
    private Label colunaF;
    private Label colunaG;
    private Label colunaH;
    private Label colunaI;
    private Label colunaJ;
    private Label colunaK;
    private Label colunaL;
    private Label colunaM;
    private Label colunaN;
    private Label colunaO;
    private Label colunaP;
    private Label colunaQ;

    public List<Produto> carregaProdutos(String caminhoPlanilha) {
        List<Produto> produtos = new ArrayList<>();
        try {
            Workbook workbook = Workbook.getWorkbook(new File(caminhoPlanilha));
            Sheet sheet = workbook.getSheet(0);
            for (int i = 1; i < sheet.getRows(); i++) {
                Produto produto = new Produto();
                Cell colunaA = sheet.getCell(0, i);
                Cell colunaB = sheet.getCell(1, i);
                Cell colunaC = sheet.getCell(2, i);
                Cell colunaD = sheet.getCell(3, i);
                Cell colunaE = sheet.getCell(4, i);
                Cell colunaF = sheet.getCell(5, i);
                Cell colunaG = sheet.getCell(6, i);
                Cell colunaH = sheet.getCell(7, i);
                Cell colunaI = sheet.getCell(8, i);
                Cell colunaJ = sheet.getCell(9, i);
                Cell colunaK = sheet.getCell(10, i);
                Cell colunaL = sheet.getCell(11, i);
                Cell colunaM = sheet.getCell(12, i);
                Cell colunaN = sheet.getCell(13, i);
                Cell colunaO = sheet.getCell(14, i);
                Cell colunaP = sheet.getCell(15, i);
                Cell colunaQ = sheet.getCell(16, i);
                produto.setCodprod(colunaA.getContents());
                produto.setReferencia(colunaB.getContents());
                produto.setDescricao(colunaC.getContents());
                produto.setCodigoncm(colunaD.getContents());
                produto.setCodnatreceita(colunaE.getContents());
                produto.setCodtribut00(colunaF.getContents());
                produto.setBaseicmsreg00(colunaG.getContents());
                produto.setAliqicmsreg00(colunaH.getContents());
                produto.setAliqicmspdv(colunaI.getContents());
                produto.setPis_cst(colunaJ.getContents());
                produto.setCofins_cst(colunaK.getContents());
                produto.setAliqpis(colunaL.getContents());
                produto.setAliqcofins(colunaM.getContents());
                produto.setPisent_cst(colunaN.getContents());
                produto.setCofinsent_cst(colunaO.getContents());
                produto.setAliqpisent(colunaP.getContents());
                produto.setAliqcofinsent(colunaQ.getContents());
                produtos.add(produto);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Erro ao carregar a planilha!" + e);
        }
        return produtos;
    }

    public void exportaDados(String caminhoPlanilha, List<Produto> produtos, JProgressBar barra) {
        try {
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("pt_br"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(caminhoPlanilha), ws);
            WritableSheet s = workbook.createSheet("Cotação", 0);
            WritableCellFormat cf2 = new WritableCellFormat();
            WritableFont bold = new WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            WritableCellFormat arial10font = new WritableCellFormat(bold);
            bold.setColour(Colour.WHITE);
            arial10font.setBackground(Colour.GREY_40_PERCENT);
            arial10font.setFont(bold);
            
            barra.setMaximum(produtos.size());
            for (int i = 0; i < produtos.size(); i++) {
                if (i == 0) {
                    colunaA = new Label(0, i, "CÓDIGO", arial10font);
                    colunaB = new Label(1, i, "REFERENCIA", arial10font);
                    colunaC = new Label(2, i, "DESCRICAO", arial10font);
                    colunaD = new Label(3, i, "NCM", arial10font);
                    colunaE = new Label(4, i, "NAT-RECEITA", arial10font);
                    colunaF = new Label(5, i, "CST", arial10font);
                    colunaG = new Label(6, i, "BASE", arial10font);
                    colunaH = new Label(7, i, "ICMS-ENT", arial10font);
                    colunaI = new Label(8, i, "ICMS-SAI", arial10font);
                    colunaJ = new Label(9, i, "PIS_SAI", arial10font);
                    colunaK = new Label(10, i, "COFINS_SAI", arial10font);
                    colunaL = new Label(11, i, "ALIQPIS_SAI", arial10font);
                    colunaM = new Label(12, i, "ALIQCOFINS_SAI", arial10font);
                    colunaN = new Label(13, i, "PIS_ENT", arial10font);
                    colunaO = new Label(14, i, "COFINS_ENT", arial10font);
                    colunaP = new Label(15, i, "ALIQPIS_ENT", arial10font);
                    colunaQ = new Label(16, i, "ALIQCOFINS_ENT", arial10font);
                } else {
                    colunaA = new Label(0, i, produtos.get(i).getCodprod(), cf2);
                    colunaB = new Label(1, i, produtos.get(i).getReferencia(), cf2);
                    colunaC = new Label(2, i, produtos.get(i).getDescricao(), cf2);
                    colunaD = new Label(3, i, produtos.get(i).getCodigoncm(), cf2);
                    colunaE = new Label(4, i, produtos.get(i).getCodnatreceita(), cf2);
                    colunaF = new Label(5, i, produtos.get(i).getCodtribut00(), cf2);
                    colunaG = new Label(6, i, produtos.get(i).getBaseicmsreg00(), cf2);
                    colunaH = new Label(7, i, produtos.get(i).getAliqicmsreg00(), cf2);
                    colunaI = new Label(8, i, produtos.get(i).getAliqicmspdv(), cf2);
                    colunaJ = new Label(9, i, produtos.get(i).getPis_cst(), cf2);
                    colunaK = new Label(10, i, produtos.get(i).getCofins_cst(), cf2);
                    colunaL = new Label(11, i, produtos.get(i).getAliqpis(), cf2);
                    colunaM = new Label(12, i, produtos.get(i).getAliqcofins(), cf2);
                    colunaN = new Label(13, i, produtos.get(i).getPisent_cst(), cf2);
                    colunaO = new Label(14, i, produtos.get(i).getCofinsent_cst(), cf2);
                    colunaP = new Label(15, i, produtos.get(i).getAliqpisent(), cf2);
                    colunaQ = new Label(16, i, produtos.get(i).getAliqcofinsent(), cf2);
                }
                s.addCell(colunaA);
                s.addCell(colunaB);
                s.addCell(colunaC);
                s.addCell(colunaD);
                s.addCell(colunaE);
                s.addCell(colunaF);
                s.addCell(colunaG);
                s.addCell(colunaH);
                s.addCell(colunaI);
                s.addCell(colunaJ);
                s.addCell(colunaK);
                s.addCell(colunaL);
                s.addCell(colunaM);
                s.addCell(colunaN);
                s.addCell(colunaO);
                s.addCell(colunaP);
                s.addCell(colunaQ);
                barra.setValue(barra.getValue() + 1);
            }
            workbook.write();
            workbook.close();
            JOptionPane.showMessageDialog(null, "Exportação Realizada com sucesso!");
            barra.setValue(0);
        } catch (Exception e) {
            if(e.toString().contains("outro processo")==true){
                JOptionPane.showMessageDialog(null, "O arquivo selecionado está aberto!\n");
            }else{
                JOptionPane.showMessageDialog(null, "Erro ao exportar os produtos!\n"+e);
            }
        }
    }
}
