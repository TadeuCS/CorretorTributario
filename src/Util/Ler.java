/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Util;

import Model.Produto;
import java.io.File;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javax.swing.JFileChooser;

import jxl.Cell;

import jxl.Sheet;

import jxl.Workbook;

import jxl.read.biff.BiffException;

public class Ler {

    public static void main(String[] args)
            throws BiffException, IOException {

        /**
         *
         * Carrega a planilha
         *
         */
        String filename = null;
        JFileChooser chooser = new JFileChooser();

        int retorno = chooser.showOpenDialog(null);
        if (retorno == JFileChooser.APPROVE_OPTION) {
            filename = chooser.getSelectedFile().getAbsolutePath();

//            Workbook workbook = Workbook.getWorkbook(new File("C:/Users/Tadeu/Desktop/teste.xls"));
            Workbook workbook = Workbook.getWorkbook(new File(filename));

            /**
             *
             * Aqui Ã© feito o controle de qual aba do xls * serÃ¡ realiza a
             * leitura dos dados
             *
             */
            Sheet sheet = workbook.getSheet(0);

            /**
             *
             * Numero de linhas com dados do xls
             *
             */
            List<Produto> produtos = new ArrayList<>();
            
            int linhas = sheet.getRows();

            for (int i = 1; i < linhas; i++) {
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
            System.out.println(produtos.size());
        }

    }
}
