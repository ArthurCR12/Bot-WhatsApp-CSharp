using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using OfficeOpenXml;


namespace WppAuto
{
    public class AutomationWebChrome
    {
        ChromeDriver driverC = new ChromeDriver();
        
        public void WppChrome()
        {
            try
            {
                // Definindo o contexto de licença antes de usar o EPPlus
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // Pegando o Caminho da Pasta aonde está rodando
                string pastaApp = AppDomain.CurrentDomain.BaseDirectory;

                // Passando aonde a planilha está
                string caminhoPlan = Path.Combine(pastaApp, "contatos.xlsx");

                driverC.Navigate().GoToUrl("https://www.google.com.br/");
                driverC.Navigate().GoToUrl("https://web.whatsapp.com/");

                WebDriverWait waitSide = new WebDriverWait(driverC, TimeSpan.FromSeconds(100));
                waitSide.Until(ExpectedConditions.ElementIsVisible(By.Id("side")));

                // Verifica se a planilha existe
                if (File.Exists(caminhoPlan))
                {

                    // Carregar a planilha
                    using (var package = new ExcelPackage(new FileInfo(caminhoPlan)))
                    {
                        // Escolher a primeira planilha no arquivo
                        var worksheet = package.Workbook.Worksheets["Sheet1"];

                        for (int linha = 2; linha <= worksheet.Dimension.Rows; linha++)
                        {
                            var nome = worksheet.Cells[linha, 1].Text;
                            var telefone = worksheet.Cells[linha, 2].Text;
                            var mensagem = worksheet.Cells[linha, 3].Text;

                            if (telefone == "")
                            {
                                driverC.Quit();
                                MessageBox.Show("Não foi possivel identificar o telefone!\nO programa será encerrado!", "Fim da Planilha", MessageBoxButtons.OK, MessageBoxIcon.Information);                                
                                break;
                            }

                            string texto = Uri.EscapeDataString($"Olá {nome} ! \n{mensagem}");

                            string link = $"https://web.whatsapp.com/send?phone={telefone}&text={texto}";

                            driverC.Navigate().GoToUrl(link);

                            WebDriverWait waitSendButton = new WebDriverWait(driverC, TimeSpan.FromSeconds(100));
                            // Espera até que o botão de enviar esteja visível
                            waitSendButton.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span")));
                            driverC.FindElement(By.XPath("//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span")).Click();
                            Thread.Sleep(7000);
                        }
                    }
                }
                // Se A planilha não existir
                else
                {
                    OpenFileDialog dialog = new OpenFileDialog();
                    dialog.Filter = "Arquivos do Excel|*.xlsx|Todos os arquivos|*.*";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        using (var package = new ExcelPackage(new FileInfo(dialog.FileName)))
                        {
                            // Escolher a primeira planilha no arquivo
                            var worksheet = package.Workbook.Worksheets["Sheet1"];

                            for (int linha = 2; linha <= worksheet.Dimension.Rows; linha++)
                            {
                                var nome = worksheet.Cells[linha, 1].Text;
                                var telefone = worksheet.Cells[linha, 2].Text;
                                var mensagem = worksheet.Cells[linha, 3].Text;

                                string texto = Uri.EscapeDataString($"Olá {nome} ! \n{mensagem}");

                                string link = $"https://web.whatsapp.com/send?phone={telefone}&text={texto}";

                                driverC.Navigate().GoToUrl(link);

                                WebDriverWait waitSendButton = new WebDriverWait(driverC, TimeSpan.FromSeconds(100));
                                // Espera até que o botão de enviar esteja visível
                                waitSendButton.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span")));
                                driverC.FindElement(By.XPath("//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span")).Click();
                                Thread.Sleep(7000);
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
