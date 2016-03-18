/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package View;

import Model.Produto;
import Util.Conexao;
import Util.ExcelControl;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

public class Frm_Principal extends javax.swing.JFrame {

    Conexao con;
    Statement st;
    ResultSet rs;
    List<Produto> produtos;
    Produto produto;
    ExcelControl xls;

    public Frm_Principal() {
        initComponents();
        grupo.add(rbt_codigo);
        grupo.add(rbt_referencia);
        start();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        grupo = new javax.swing.ButtonGroup();
        jPanel1 = new javax.swing.JPanel();
        btn_importar = new javax.swing.JButton();
        btn_exportar = new javax.swing.JButton();
        barra = new javax.swing.JProgressBar();
        qtde = new javax.swing.JLabel();
        jPanel4 = new javax.swing.JPanel();
        jLabel4 = new javax.swing.JLabel();
        txt_caminhoBanco = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        txt_servidor = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        txt_usuario = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        txt_senha = new javax.swing.JPasswordField();
        btn_buscarBanco = new javax.swing.JButton();
        btn_testar = new javax.swing.JButton();
        rbt_referencia = new javax.swing.JRadioButton();
        rbt_codigo = new javax.swing.JRadioButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setTitle("Importador e Exportador de dados tributários de produtos");
        setResizable(false);

        btn_importar.setText("Importar");
        btn_importar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_importarActionPerformed(evt);
            }
        });

        btn_exportar.setText("Exportar");
        btn_exportar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_exportarActionPerformed(evt);
            }
        });

        barra.setForeground(new java.awt.Color(32, 122, 18));
        barra.setStringPainted(true);

        qtde.setFont(new java.awt.Font("Courier New", 1, 14)); // NOI18N
        qtde.setForeground(new java.awt.Color(32, 122, 18));
        qtde.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        qtde.setText("0");

        jPanel4.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));

        jLabel4.setText("Caminho *:");

        jLabel6.setText("Servidor *:");

        jLabel2.setText("Usuário *:");

        jLabel3.setText("Senha *:");

        btn_buscarBanco.setText("...");
        btn_buscarBanco.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_buscarBancoActionPerformed(evt);
            }
        });

        btn_testar.setText("Testar");
        btn_testar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_testarActionPerformed(evt);
            }
        });

        rbt_referencia.setText("Referência");

        rbt_codigo.setText("Código");

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel2)
                            .addComponent(jLabel4)
                            .addComponent(jLabel6))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addComponent(txt_usuario, javax.swing.GroupLayout.PREFERRED_SIZE, 145, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jLabel3)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(txt_senha, javax.swing.GroupLayout.PREFERRED_SIZE, 139, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addComponent(txt_servidor, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addComponent(txt_caminhoBanco, javax.swing.GroupLayout.PREFERRED_SIZE, 361, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, Short.MAX_VALUE)
                                .addComponent(btn_buscarBanco))))
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addComponent(rbt_codigo)
                        .addGap(18, 18, 18)
                        .addComponent(rbt_referencia)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(btn_testar)))
                .addContainerGap())
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(txt_servidor, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(6, 6, 6)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(txt_caminhoBanco, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btn_buscarBanco))
                .addGap(6, 6, 6)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(txt_usuario, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3)
                    .addComponent(txt_senha, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 11, Short.MAX_VALUE)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel4Layout.createSequentialGroup()
                        .addComponent(btn_testar)
                        .addGap(6, 6, 6))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel4Layout.createSequentialGroup()
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(rbt_codigo)
                            .addComponent(rbt_referencia))
                        .addContainerGap())))
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(btn_exportar, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(qtde, javax.swing.GroupLayout.PREFERRED_SIZE, 74, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(barra, javax.swing.GroupLayout.PREFERRED_SIZE, 286, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(btn_importar, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(btn_importar)
                        .addComponent(btn_exportar))
                    .addComponent(barra, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(qtde, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void btn_buscarBancoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_buscarBancoActionPerformed
        txt_caminhoBanco.setText(open());
    }//GEN-LAST:event_btn_buscarBancoActionPerformed

    private void btn_exportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_exportarActionPerformed
        if (validaCampos() == true) {
            con = new Conexao();
            st = con.getConexao(txt_servidor.getText(), txt_caminhoBanco.getText(), txt_usuario.getText(), txt_senha.getText());
            exporta(save(), st);
        }
    }//GEN-LAST:event_btn_exportarActionPerformed

    private void btn_testarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_testarActionPerformed
        if (txt_servidor.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "IP ou nome do Servidor inválido!");
            txt_servidor.requestFocus();
        } else {
            if (txt_caminhoBanco.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(null, "Caminho do banco de dados inválido!");
                txt_caminhoBanco.requestFocus();
            } else {
                if (txt_usuario.getText().trim().isEmpty()) {
                    JOptionPane.showMessageDialog(null, "Usuário inválido!");
                    txt_usuario.requestFocus();
                } else {
                    if (txt_senha.getText().trim().isEmpty()) {
                        JOptionPane.showMessageDialog(null, "Senha inválida!");
                        txt_senha.requestFocus();
                    } else {
                        testaConexao();
                    }
                }
            }
        }
    }//GEN-LAST:event_btn_testarActionPerformed

    private void btn_importarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_importarActionPerformed
        if (rbt_codigo.getSelectedObjects() != null || rbt_referencia.getSelectedObjects() != null) {
            if (validaCampos() == true) {
                con = new Conexao();
                st = con.getConexao(txt_servidor.getText(), txt_caminhoBanco.getText(), txt_usuario.getText(), txt_senha.getText());
                importa(open(), st);
            }
        } else {
            JOptionPane.showMessageDialog(null, "Selecione se vai importar pela coluna CODIGO ou REFERÊNCIA");
        }
    }//GEN-LAST:event_btn_importarActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;

                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Frm_Principal.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Frm_Principal.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Frm_Principal.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Frm_Principal.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Frm_Principal().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JProgressBar barra;
    private javax.swing.JButton btn_buscarBanco;
    private javax.swing.JButton btn_exportar;
    private javax.swing.JButton btn_importar;
    private javax.swing.JButton btn_testar;
    private javax.swing.ButtonGroup grupo;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JLabel qtde;
    private javax.swing.JRadioButton rbt_codigo;
    private javax.swing.JRadioButton rbt_referencia;
    private javax.swing.JTextField txt_caminhoBanco;
    private javax.swing.JPasswordField txt_senha;
    private javax.swing.JTextField txt_servidor;
    private javax.swing.JTextField txt_usuario;
    // End of variables declaration//GEN-END:variables

    private String open() {
        String diretorio = null;
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(null);
        if (result == JFileChooser.CANCEL_OPTION) {
        } else {
            diretorio = fileChooser.getSelectedFile().getPath();
        }
        return diretorio;
    }

    private String save() {
        String diretorio = null;
        JFileChooser chooser = new JFileChooser();
        int result = chooser.showSaveDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            if (chooser.getSelectedFile().getAbsolutePath().contains(".xls") == true) {
                diretorio = chooser.getSelectedFile().getAbsolutePath();
            } else {
                diretorio = chooser.getSelectedFile().getAbsolutePath() + ".xls";
            }

        }
        return diretorio;
    }

    private boolean validaCampos() {
        boolean retorno = false;
        if (txt_servidor.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "IP ou nome do Servidor inválido!");
            txt_servidor.requestFocus();
        } else {
            if (txt_caminhoBanco.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(null, "Caminho do banco de dados inválido!");
                txt_caminhoBanco.requestFocus();
            } else {
                if (txt_usuario.getText().trim().isEmpty()) {
                    JOptionPane.showMessageDialog(null, "Usuário inválido!");
                    txt_usuario.requestFocus();
                } else {
                    if (txt_senha.getText().trim().isEmpty()) {
                        JOptionPane.showMessageDialog(null, "Senha inválida!");
                        txt_senha.requestFocus();
                    } else {
                        retorno = true;
                    }
                }
            }
        }
        return retorno;
    }

    private void testaConexao() {
        con = new Conexao();
        if (con.getConexao(txt_servidor.getText(), txt_caminhoBanco.getText(), txt_usuario.getText(), txt_senha.getText()) != null) {
            JOptionPane.showMessageDialog(null, "Conexão bem sucedida!");
        } else {
            JOptionPane.showMessageDialog(null, "Não foi possivel conectar no banco de dados!");
        }
    }

    private String trataCamposByQtde(int tamanho, String grupo) {
        String retorno = grupo;
        while (retorno.length() < tamanho) {
            retorno = "0" + retorno;
        }
        return retorno;
    }

    private void importa(String caminhoPlanilha, Statement st) {
        if (caminhoPlanilha != null && st != null) {
            xls = new ExcelControl();
            produtos = new ArrayList<>();
            produtos = xls.carregaProdutos(caminhoPlanilha);
            barra.setMaximum(produtos.size());
            Thread acao;
            acao = new Thread(new Runnable() {
                @Override
                public void run() {
                    long time1 = System.currentTimeMillis();

                    try {
                        if (rbt_codigo.isSelected()) {
                            for (int i = 0; i < produtos.size(); i++) {
                                produto = produtos.get(i);
                                st.executeUpdate("UPDATE PRODUTO p SET "
                                        + "p.CODCLASFIS=(SELECT FIRST 1 SKIP 0 C.CODCLASFIS FROM CLASFISC C WHERE C.CODIGONCM LIKE '" + produto.getCodigoncm() + "'),"
                                        + "p.codtribut00='" + trataCamposByQtde(3, produto.getCodtribut00()) + "',"
                                        + "p.baseicmsreg00='" + produto.getBaseicmsreg00() + "',"
                                        + "p.aliqicmsreg00=" + Double.parseDouble(produto.getAliqicmsreg00().replace(",", ".")) + ","
                                        + "p.aliqicmspdv=" + Double.parseDouble(produto.getAliqicmspdv().replace(",", ".")) + ","
                                        + "p.codcest='" + produto.getCodcest().replace(".", "") + "'"
                                        + " where p.codprod='" + produto.getCodprod() + "';");
                                st.executeUpdate("UPDATE PRODUTODETALHE d SET "
                                        + "d.pis_cst='" + trataCamposByQtde(2, produto.getPis_cst()) + "',"
                                        + "d.cofins_cst='" + trataCamposByQtde(2, produto.getCofins_cst()) + "',"
                                        + "d.aliqpis=" + Double.parseDouble(produto.getAliqpis().replace(",", ".")) + ","
                                        + "d.aliqcofins=" + Double.parseDouble(produto.getAliqcofins().replace(",", ".")) + ","
                                        + "d.pisent_cst='" + trataCamposByQtde(2, produto.getPisent_cst()) + "',"
                                        + "d.cofinsent_cst='" + trataCamposByQtde(2, produto.getCofinsent_cst()) + "',"
                                        + "d.aliqpisent=" + Double.parseDouble(produto.getAliqpisent().replace(",", ".")) + ","
                                        + "d.aliqcofinsent=" + Double.parseDouble(produto.getAliqcofinsent().replace(",", "."))
                                        + " where d.codprod='" + produto.getCodprod() + "';");
                                st.executeUpdate("UPDATE CLASFISC C SET C.CODNATRECEITA='" + produto.getCodnatreceita() + "' WHERE C.CODIGONCM LIKE '" + produto.getCodigoncm() + "';");
                                barra.setValue(barra.getValue() + 1);
                                qtde.setText(Integer.parseInt(qtde.getText()) + 1 + "");
                            }
                        } else {
                            for (int i = 0; i < produtos.size(); i++) {
                                produto = produtos.get(i);
                                if (!produto.getReferencia().trim().isEmpty()) {
                                    st.executeUpdate("UPDATE PRODUTO p SET "
                                            + "p.CODCLASFIS=(SELECT FIRST 1 SKIP 0 C.CODCLASFIS FROM CLASFISC C WHERE C.CODIGONCM LIKE '" + produto.getCodigoncm() + "'),"
                                            + "p.codtribut00='" + trataCamposByQtde(3, produto.getCodtribut00()) + "',"
                                            + "p.baseicmsreg00='" + produto.getBaseicmsreg00() + "',"
                                            + "p.aliqicmsreg00=" + Double.parseDouble(produto.getAliqicmsreg00().replace(",", ".")) + ","
                                            + "p.aliqicmspdv=" + Double.parseDouble(produto.getAliqicmspdv().replace(",", ".")) + ","
                                            + "p.codcest='" + produto.getCodcest().replace(".", "") + "'"
                                            + " where p.referencia='" + produto.getReferencia() + "';");
                                    st.executeUpdate("UPDATE PRODUTODETALHE d SET "
                                            + "d.pis_cst='" + trataCamposByQtde(2, produto.getPis_cst()) + "',"
                                            + "d.cofins_cst='" + trataCamposByQtde(2, produto.getCofins_cst()) + "',"
                                            + "d.aliqpis=" + Double.parseDouble(produto.getAliqpis().replace(",", ".")) + ","
                                            + "d.aliqcofins=" + Double.parseDouble(produto.getAliqcofins().replace(",", ".")) + ","
                                            + "d.pisent_cst='" + trataCamposByQtde(2, produto.getPisent_cst()) + "',"
                                            + "d.cofinsent_cst='" + trataCamposByQtde(2, produto.getCofinsent_cst()) + "',"
                                            + "d.aliqpisent=" + Double.parseDouble(produto.getAliqpisent().replace(",", ".")) + ","
                                            + "d.aliqcofinsent=" + Double.parseDouble(produto.getAliqcofinsent().replace(",", "."))
                                            + " where d.codprod=(select p.codprod from produto p where p.referencia like '" + produto.getReferencia() + "');");
                                    st.executeUpdate("UPDATE CLASFISC C SET C.CODNATRECEITA='" + produto.getCodnatreceita() + "' WHERE C.CODIGONCM LIKE '" + produto.getCodigoncm() + "';");
                                }
                                barra.setValue(barra.getValue() + 1);
                                qtde.setText(Integer.parseInt(qtde.getText()) + 1 + "");
                            }
                        }
                        JOptionPane.showMessageDialog(null, "Importação Realizada com sucesso!");
                        barra.setValue(0);
                        qtde.setText("0");
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(null, "Erro ao importar os dados do produto: " + produto.getCodprod() + "\n" + e);
                    } finally {
                        long time2 = System.currentTimeMillis();
                        JOptionPane.showMessageDialog(null, "A importação demorou: "+new SimpleDateFormat("mm:ss").format(new Date(time2 - time1)));
                    }
                }

            }
            );
            acao.start();
        }
    }

    private void exporta(String caminhoPlanilha, Statement st) {
        try {
            produtos = new ArrayList<>();
            rs = st.executeQuery("select\n"
                    + "p.CODPROD,p.REFERENCIA,p.DESCRICAO,p.CODCEST,c.CODIGONCM,c.CODNATRECEITA,p.CODTRIBUT00,p.BASEICMSREG00,p.ALIQICMSREG00,p.ALIQICMSPDV,\n"
                    + "d.PIS_CST,d.COFINS_CST,d.ALIQPIS,d.ALIQCOFINS,d.PISent_CST,d.COFINSent_CST,d.ALIQPISent,d.ALIQCOFINSent\n"
                    + "from produto p\n"
                    + "inner join produtodetalhe d on p.CODPROD=d.CODPROD\n"
                    + "inner join clasfisc c on p.CODCLASFIS=c.CODCLASFIS "
                    + "where p.ativo='S' order by p.descricao");
            while (rs.next()) {
                produto = new Produto();
                produto.setCodprod(rs.getString("codprod"));
                produto.setReferencia(rs.getString("referencia"));
                produto.setDescricao(rs.getString("descricao"));
                produto.setCodigoncm(rs.getString("codigoncm"));
                produto.setCodnatreceita(rs.getString("codnatreceita"));
                produto.setCodtribut00(rs.getString("codtribut00"));
                produto.setBaseicmsreg00(rs.getString("baseicmsreg00"));
                produto.setAliqicmsreg00(rs.getString("aliqicmsreg00"));
                produto.setAliqicmspdv(rs.getString("aliqicmspdv"));
                produto.setPis_cst(rs.getString("pis_cst"));
                produto.setCofins_cst(rs.getString("cofins_cst"));
                produto.setAliqpis(rs.getString("aliqpis"));
                produto.setAliqcofins(rs.getString("aliqcofins"));
                produto.setPisent_cst(rs.getString("pisent_cst"));
                produto.setCofinsent_cst(rs.getString("cofinsent_cst"));
                produto.setAliqpisent(rs.getString("aliqpisent"));
                produto.setAliqcofinsent(rs.getString("aliqcofinsent"));
                produto.setCodcest(rs.getString("CODCEST"));
                produtos.add(produto);
            }
            xls = new ExcelControl();
            Thread acao;
            acao = new Thread(new Runnable() {
                @Override
                public void run() {
                    xls.exportaDados(caminhoPlanilha, produtos, barra);
                }
            }
            );
            acao.start();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Erro ao exportar o produto: " + produto.getCodprod() + "\n" + e);
        }
    }

    private void start() {
        txt_servidor.setText("localhost");
        txt_usuario.setText("SYSDBA");
        txt_senha.setText("masterkey");
        rbt_codigo.setSelected(true);
    }
}
