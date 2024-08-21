using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace LeitorDePlanilhas
{
    public class LeitorDePlanilhas
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Insira o caminho da pasta:");
            string p_folderPath = Console.ReadLine();

            if (Directory.Exists(p_folderPath))
            {
                var l_files = Directory.GetFiles(p_folderPath, "*.xlsx");
                foreach (var l_excel in l_files)
                {
                    var l_compras = ProcessExcel(l_excel);

                    // Imprimir os itens coletados
                    foreach (var l_compra in l_compras)
                    {
                        Console.WriteLine("Data da Compra: " + l_compra.DataDaCompra);
                        Console.WriteLine("Produto: " + l_compra.Produto);
                        Console.WriteLine("Preço Original x Preço Desconto: " + l_compra.PrecoOriginalXPrecoDesconto);
                        Console.WriteLine("Fornecedor: " + l_compra.Fornecedor);
                        Console.WriteLine("Quantidade Comprada: " + l_compra.QtdComprada);
                        Console.WriteLine("Comprador: " + l_compra.Comprador);
                        Console.WriteLine("Nome do Gerente: " + l_compra.NomeGerente);
                        Console.WriteLine("Sobrenome do Gerente: " + l_compra.SobrenomeGerente);
                        Console.WriteLine("Região Destino: " + l_compra.RegiaoDestino);
                        Console.WriteLine("-------------------------------------------");
                    }
                }
            }
            else
            {
                Console.WriteLine("O caminho especificado não existe.");
            }
        }

        public static List<Compras> ProcessExcel(string p_excel)
        {
            var l_workbook = new XLWorkbook(p_excel);
            var l_worksheet = l_workbook.Worksheet(1);
            var l_comprasList = new List<Compras>();

            for (int l_row = 2; l_row <= l_worksheet.LastRowUsed().RowNumber(); l_row++)
            {
                var l_rowData = l_worksheet.Row(l_row);
                var l_compra = new Compras()
                {
                    DataDaCompra = l_rowData.Cell(1).GetValue<string>(),
                    Produto = l_rowData.Cell(2).GetValue<string>(),
                    PrecoOriginalXPrecoDesconto = l_rowData.Cell(3).GetValue<string>(),
                    Fornecedor = l_rowData.Cell(4).GetValue<string>(),
                    QtdComprada = l_rowData.Cell(5).GetValue<string>(),
                    Comprador = l_rowData.Cell(6).GetValue<string>(),
                    NomeGerente = l_rowData.Cell(7).GetValue<string>(),
                    SobrenomeGerente = l_rowData.Cell(8).GetValue<string>(),
                    RegiaoDestino = l_rowData.Cell(9).GetValue<string>(),
                };

                l_comprasList.Add(l_compra);
            }

            return l_comprasList;
        }
    }
}