using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.IO;
using ClosedXML.Excel;

namespace app
{
    public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddRouting();
        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapGet("/", async context =>
                {
                    string nome = context.Request.Query["nome"];
                    string email = context.Request.Query["email"];
                    string cpf = context.Request.Query["cpf"];
                    string celular = context.Request.Query["celular"];

                    string nomeArquivo = "dados.xlsx";
                    string caminhoArquivo = Path.Combine(env.ContentRootPath, "Arquivos", nomeArquivo);

                    if (nome != null && email != null && cpf != null && celular != null)
                    {
                        SaveToExcel(caminhoArquivo, nome, email, cpf, celular);
                    }

                    await context.Response.WriteAsync(@"
                        <html>
                        <body>
                            <h1>Adicionar Dados</h1>
                            <form action="""" method=""get"">
                                <label for=""nome"">Nome:</label>
                                <input type=""text"" id=""nome"" name=""nome"" required><br>

                                <label for=""email"">Email:</label>
                                <input type=""email"" id=""email"" name=""email"" required><br>

                                <label for=""cpf"">CPF:</label>
                                <input type=""text"" id=""cpf"" name=""cpf"" required><br>

                                <label for=""celular"">Celular:</label>
                                <input type=""text"" id=""celular"" name=""celular"" required><br>

                                <input type=""submit"" value=""Adicionar"">
                            </form>
                        </body>
                        </html>");
                });
            });
        }

        private void SaveToExcel(string caminhoArquivo, string nome, string email, string cpf, string celular)
        {
            if (!File.Exists(caminhoArquivo))
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Dados");
                    worksheet.Cell(1, 1).Value = "Nome";
                    worksheet.Cell(1, 2).Value = "Email";
                    worksheet.Cell(1, 3).Value = "CPF";
                    worksheet.Cell(1, 4).Value = "Celular";
                    worksheet.Cell(2, 1).Value = nome;
                    worksheet.Cell(2, 2).Value = email;
                    worksheet.Cell(2, 3).Value = cpf;
                    worksheet.Cell(2, 4).Value = celular;
                    workbook.SaveAs(caminhoArquivo);
                }
            }
            else
            {
                using (var workbook = new XLWorkbook(caminhoArquivo))
                {
                    var worksheet = workbook.Worksheet(1);
                    int linha = worksheet.LastRowUsed().RowNumber() + 1;
                    worksheet.Cell(linha, 1).Value = nome;
                    worksheet.Cell(linha, 2).Value = email;
                    worksheet.Cell(linha, 3).Value = cpf;
                    worksheet.Cell(linha, 4).Value = celular;
                    workbook.Save();
                }
            }
        }
    }
}

