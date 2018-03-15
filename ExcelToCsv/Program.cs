using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelToCsv
{
    internal class Program
    {
        public static string destinoArquivo = "";

        private static void Main(string[] args)
        {
            //diretório que pega os arquivos (pega todos os arquivos desse diretório)
            var arquivos = Directory.GetFiles(@"C:\Temp\");

            //diretório que grava os novos arquivos
            string diretorioImp = @"C:\Temp\";

            //listando os arquivos
            foreach (var arq in arquivos)
            {
                FileInfo fi = new FileInfo(arq);

                //verifica se o arquivo tem essa extensão
                if (fi.Extension == ".xlsx")
                {
                    //csvOutputFile = fi.FullName.Replace(".xlsx", ".csv").Replace(" ", "");

                    //destino do arquivo final com o mesmo nome.
                    destinoArquivo = diretorioImp + fi.Name.Replace(".xlsx",".csv").Replace(" ","");

                    //método que converte para csv
                    ConvertExcelToCsv(arq, 1);
                }
            }
        }

        /// <summary>
        /// Método que converte excel para CSV
        /// </summary>
        /// <param name="excelFilePath">string</param>
        /// <param name="worksheetNumber">string</param>
        private static void ConvertExcelToCsv(string excelFilePath, int worksheetNumber = 1)
        {
            //Excel - nome utilizado no using (início da classe)
            //Abre o excel
            _Application oApp = new Application();

            try
            {
                FileInfo fi = new FileInfo(excelFilePath);
                if (fi.Exists)
                {
                    oApp.Visible = false; //Não mostra para o usuário que abriu o arquivo

                    //Abre
                    Workbook oWorkbook = oApp.Workbooks.Open(excelFilePath);

                    foreach (Worksheet _workSheet in oWorkbook.Worksheets)
                    {
                        //Navega nas linhas e colunas
                        NavegarColunasELinhas(_workSheet);
                    }

                    //fecha o workbook
                    oWorkbook.Close();
                    oWorkbook = null;
                }
            }
            catch (Exception e)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                throw e;
            }
            finally
            {
                //seta para null para retirar da memória                
                oApp.Quit();
                oApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Método que navega nas colunas e linhas
        /// </summary>
        /// <param name="oWorksheet">Worksheet</param>
        private static void NavegarColunasELinhas(Worksheet oWorksheet)
        {
            using (var wtr = new StreamWriter(destinoArquivo, true, Encoding.GetEncoding(1252)))
            {
                //Pega a quantidade de coluna e linha
                int numeroColuna = oWorksheet.UsedRange.Columns.Count;
                int numeroLinha = oWorksheet.UsedRange.Rows.Count;

                //Lê o arquivo linha por linha através do array
                object[,] array = oWorksheet.UsedRange.Value;

                for (int i = 1; i <= numeroLinha; i++)
                {
                    bool firstLine = true;
                    //int count = 1;

                    for (int j = 1; j <= numeroColuna; j++)
                    {
                        if (array[i, j] != null)
                        {
                            //verifica se é a primeira linha
                            if (!firstLine)
                            {
                                wtr.Write(",");
                            }
                            else
                            {
                                firstLine = false;
                            }

                            //escreve todos os dados no arquivo
                            wtr.Write(String.Format("\"{0}\"", array[i, j].ToString()));
                        }
                    }
                    wtr.WriteLine();
                }
            }
        }
    }
}