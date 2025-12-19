namespace Escala
{
    /// <summary>
    /// Diálogo para seleção de planilha/aba do Excel
    /// </summary>
    public class SeletorPlanilha : Form
    {
        public ComboBox CbPlanilhas;
        private Button BtnOk;

        public SeletorPlanilha(List<string> planilhas)
        {
            this.Text = "Selecione a Aba";
            this.Size = new Size(300, 150);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            Label lbl = new Label()
            {
                Text = "Escolha a planilha:",
                Left = 10,
                Top = 10,
                Width = 200
            };

            CbPlanilhas = new ComboBox()
            {
                Left = 10,
                Top = 35,
                Width = 260,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            CbPlanilhas.DataSource = planilhas;

            BtnOk = new Button()
            {
                Text = "Carregar",
                Left = 190,
                Top = 70,
                Width = 80,
                DialogResult = DialogResult.OK
            };

            this.Controls.Add(lbl);
            this.Controls.Add(CbPlanilhas);
            this.Controls.Add(BtnOk);
            this.AcceptButton = BtnOk;
        }
    }
}
