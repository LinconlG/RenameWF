using System;
using System.IO;
using System.Windows.Forms;

namespace RenomearWF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string diretorio;
        string diretorioExcel;
        int linhas;
        int colunas = 2;
        string extensao;
        string modrev;//estilo de "rev." para ser alterado



        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                diretorio = $@"{textBox1.Text}";
                diretorioExcel = $@"{textBox2.Text}";
                linhas = Convert.ToInt32(textBox3.Text);
                modrev = textBox4.Text;

                string[] nomesArquivos = new string[linhas];
                string[] revisoes = new string[linhas];


                DirectoryInfo diretorioPasta = new DirectoryInfo($@"{diretorio}");

                var planilha = new Microsoft.Office.Interop.Excel.Application();
                var wb = planilha.Workbooks.Open($@"{diretorioExcel}", ReadOnly: true);
                var ws = wb.Worksheets[1];
                var r = ws.Range["A1"].Resize[linhas, colunas];
                var array = r.Value;

                for (int i = 1; i <= linhas; i++)
                {
                    for (int j = 1; j <= colunas; j++)
                    {
                        string text = Convert.ToString(array[i, j]);

                        if (j == 1)
                        {
                            nomesArquivos[i - 1] = text;
                        }
                        else
                        {
                            revisoes[i - 1] = text;
                        }
                    }
                }

                DirectoryInfo[] listaDePastas = diretorioPasta.GetDirectories();
                FileInfo[] listaArquivos = diretorioPasta.GetFiles();

                RenomearArquivo(listaArquivos, nomesArquivos, revisoes);




                foreach (DirectoryInfo item in listaDePastas)
                {

                    //copiar o codigo de cima

                    int ind = Path.GetDirectoryName(item.FullName).ToLower().IndexOf($"{modrev}");
                    string caminhoCompleto = item.FullName;
                    bool flag = false;

                    for (int i = 0; i < linhas; i++)
                    {

                        if (ind != -1)
                        {
                            if (flag == false)
                            {
                                int inddot = Path.GetFileName(item.FullName).LastIndexOf(".");
                                string rev = Path.GetFileName(item.FullName).Substring(ind, inddot - ind);
                                caminhoCompleto = caminhoCompleto.Replace($"{rev}", $"");
                                File.Move(item.FullName, caminhoCompleto);

                                File.Move(caminhoCompleto, Path.Combine(caminhoCompleto, $"{item.FullName} Rev.{revisoes[i]}"));

                                flag = true;

                                if (checkBox1.Enabled)
                                {

                                    //não executando
                                    DirectoryInfo diretorio2 = new DirectoryInfo($@"{item.FullName} Rev.{revisoes[i]}");
                                    FileInfo[] listaArquivosPasta = diretorio2.GetFiles();

                                    RenomearArqPasta(diretorio2);
                                }
                                break;
                            }
                        }
                        else
                        {
                            Directory.Move(caminhoCompleto, Path.Combine(caminhoCompleto, $"{caminhoCompleto} Rev.{revisoes[i]}"));

                            if (checkBox1.Enabled)
                            {
                                DirectoryInfo diretorio2 = new DirectoryInfo($@"{item.FullName} Rev.{revisoes[i]}");
                                FileInfo[] listaArquivosPasta = diretorio2.GetFiles();

                                RenomearArqPasta(diretorio2);
                            }
                            break;

                        }


                    }

                }

                wb.Close();
                planilha.Quit();

                MessageBox.Show("Concluído!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                string dirPasta = textBox5.Text;

                DirectoryInfo dir = new DirectoryInfo($@"{dirPasta}");
                FileInfo[] listaArquivos = dir.GetFiles();

                string nomePasta = Path.GetFileName(dir.FullName);
                int inddot = nomePasta.LastIndexOf(".");

                if (inddot != -1)
                {
                    string ext;
                    bool flag = false;
                    string revp = nomePasta.Substring(inddot - 3);
                    string reva;
                    string temp;

                    foreach (FileInfo arquivo in listaArquivos)
                    {
                        flag = Path.GetFileName(arquivo.FullName).ToLower().Contains("rev.");
                        ext = Path.GetExtension(arquivo.FullName);

                        if (flag)
                        {
                            reva = Path.GetFileNameWithoutExtension(arquivo.FullName).Substring(Path.GetFileName(arquivo.FullName).LastIndexOf(".") - 5);
                            File.Move(arquivo.FullName, arquivo.FullName.Replace($" {reva}{ext}", $"{ext}"));
                            temp = arquivo.FullName.Replace($" {reva}{ext}", $"{ext}");
                            File.Move(temp, temp.Replace($"{ext}", $" {revp}{ext}"));
                        }
                        else
                        {
                            File.Move(arquivo.FullName, arquivo.FullName.Replace($"{ext}", $" {revp}{ext}"));
                        }


                    }

                    MessageBox.Show("Concluído!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }



        }

        public void RenomearArquivo(FileInfo[] listaArquivos, string[] nomeArq, string[] revs)
        {
            foreach (FileInfo arquivo in listaArquivos)
            {

                extensao = Path.GetExtension(arquivo.FullName);
                int ind = Path.GetFileName(arquivo.FullName).ToLower().IndexOf($"{modrev}");
                string caminhoCompleto = arquivo.FullName;
                string nomeComExtensao = Path.GetFileName(caminhoCompleto);
                bool flag = false;

                for (int i = 0; i < linhas; i++)
                {
                    if (ind != -1)
                    {
                        if (flag == false)
                        {
                            int inddot = Path.GetFileName(arquivo.FullName).LastIndexOf(".");
                            string rev = Path.GetFileName(arquivo.FullName).Substring(ind, inddot - ind);
                            caminhoCompleto = caminhoCompleto.Replace($"{rev}", $"");
                            nomeComExtensao = Path.GetFileName(caminhoCompleto);
                            flag = true;
                        }
                        if (nomeComExtensao == $"{nomeArq[i]}{extensao}")
                        {
                            File.Move(arquivo.FullName, caminhoCompleto);

                            File.Move(caminhoCompleto, caminhoCompleto.Replace($"{extensao}", $" Rev.{revs[i]}{extensao}"));
                            break;
                        }

                    }
                    else
                    {
                        if (nomeComExtensao == $"{nomeArq[i]}{extensao}")
                        {
                            File.Move(arquivo.FullName, arquivo.FullName.Replace($"{extensao}", $" Rev.{revs[i]}{extensao}"));
                            break;
                        }
                    }
                }
            }
        }

        public void RenomearArqPasta(DirectoryInfo dir)
        {
            
            FileInfo[] listaArquivos = dir.GetFiles();

            string nomePasta = Path.GetFileName(dir.FullName);
            int inddot = nomePasta.LastIndexOf(".");

            if (inddot != -1)
            {
                string ext;
                bool flag = false;
                string revp = nomePasta.Substring(inddot - 3);
                string reva;
                string temp;

                foreach (FileInfo arquivo in listaArquivos)
                {
                    flag = Path.GetFileName(arquivo.FullName).ToLower().Contains("rev.");
                    ext = Path.GetExtension(arquivo.FullName);

                    if (flag)
                    {
                        reva = Path.GetFileNameWithoutExtension(arquivo.FullName).Substring(Path.GetFileName(arquivo.FullName).LastIndexOf(".") - 5);
                        File.Move(arquivo.FullName, arquivo.FullName.Replace($" {reva}{ext}", $"{ext}"));
                        temp = arquivo.FullName.Replace($" {reva}{ext}", $"{ext}");
                        File.Move(temp, temp.Replace($"{ext}", $" {revp}{ext}"));
                    }
                    else
                    {
                        File.Move(arquivo.FullName, arquivo.FullName.Replace($"{ext}", $" {revp}{ext}"));
                    }
                }
            }

        }
    }
}
