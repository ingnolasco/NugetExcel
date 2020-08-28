
using ExcelNugget02;
using ExcelNugget02.Interfaces;
using PruebaExcelFormat;
using System;
using System.Collections.Generic;

namespace PrueebasNugguet2
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Prueba> lista = new List<Prueba>();
            Prueba pp;
            for (int a = 1; a <= 50; a++)
            {
                pp = new Prueba()
                {
                    PLANUM = "980150448864",
                    PERIODO = "201307",
                    NIT = "06140208840013",
                    RAZON_SOCIAL = "LUNA PAN S.A DE C.V.",
                    ID_SUCURSAL = "001",
                    EMPLEADOS_DECLARADOS = "3",
                    MONTO_TOTAL = "94.0800",
                    ARCHIVO = "MAX061402088400130012013070001.ZIP",
                    NPE = "",
                    FECHA_ADICION = "3/6/2020 13:55:07",
                    FECHA_PRESENTACIÓN = "27/08/2020",
                    CATEGORIA = "GEM",

                };
                lista.Add(pp);
                pp = new Prueba()
                {
                    PLANUM = "980150448864",
                    PERIODO = "201307",
                    NIT = "06140208840013",
                    RAZON_SOCIAL = "Universidad de el salvador ",
                    ID_SUCURSAL = "001",
                    EMPLEADOS_DECLARADOS = "3",
                    MONTO_TOTAL = "94.0800",
                    ARCHIVO = "MAX061402088400130012013070001.ZIP",
                    NPE = "",
                    FECHA_ADICION = "3/6/2020 13:55:07",
                    FECHA_PRESENTACIÓN = "27/08/2020",
                    CATEGORIA = "GEM",

                };
                lista.Add(pp);
            }

            List<Prueba2> lista2 = new List<Prueba2>();
            Prueba2 p2;
            for (int a = 1; a <= 100; a++)
            {
                p2 = new Prueba2()
                {
                    ID_PLANILLA = "980150448864",
                    ESTADO = "GEN",
                    NIT = "06140208840013",
                    EMPRESA = "",
                    CATEGORIA = "",
                    ASESOR = "",
                    GESTOR_EMPRESARIAL = "",
                    PERIODO = 0,
                    ID_SUCURSAL = "",
                    EMPLEADOS_DECLARADOS =0,
                    MONTO_TOTAL = "",
                    FECHA_ADICION = $"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}",
                    FECHA_ADICCION2 = $"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}",
                    ARCHIVO = "",
                    NPE="",
                    CORREOS= ""
                };
                lista2.Add(p2);
                p2 = new Prueba2()
                {
                    ID_PLANILLA = "980150448864",
                    ESTADO = "GEN",
                    NIT = "06140208840013",
                    EMPRESA = "LUNA PAN S.A DE C.V.",
                    CATEGORIA = "GEM",
                    ASESOR = "",
                    GESTOR_EMPRESARIAL = "KAREN IVETH MENA LOPEZ",
                    PERIODO = 0,
                    ID_SUCURSAL = "001",
                    EMPLEADOS_DECLARADOS = 0,
                    MONTO_TOTAL = "94.0800",
                    FECHA_ADICION = $"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}",
                    FECHA_ADICCION2 = $"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}",
                    ARCHIVO = "MAX061402088400130012013070001.ZIP",
                    NPE = "",
                    CORREOS = "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM" +
                   "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM"+
                   "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM"+
                   "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM"+
                   "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM"+
                   "DAMALYCOREAS@GMAIL.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM - GERSONJOYA.DICARSON@OUTLOOK.COM"
                };
                lista2.Add(p2);

            }

            //deuda


           // int c = 0;
            //IExcel excel = new Excel();
            //for (int s=1;s<=5;s++) {
             
            //    for (int a = 1; a <= 5; a++)
            //    {
            //        c++;
            //        Console.WriteLine(c);
            //        excel.NewContent(lista, $"hoja {c}");
            //        c++;
            //        Console.WriteLine(c);
            //        excel.NewContent(lista2, $"Hoja {c}");
            //        Console.WriteLine(c);
            //    }
            //    var resp = excel.Guardar($"Planilla ").Result;

            //    if (resp.FileName != null)
            //        Console.WriteLine(resp.FileName);
            //}

                    IExcel excel = new Excel("deuda");
  
                    excel.NewContent(lista, $"hoja 1");
             
                    excel.NewContent(lista2, $"Hoja 2");
    
                var resp = excel.Guardar($"SEPP_DNP10{DateTime.Now.ToString("ddMMyyyy")}").Result;

                if (resp.FileName != null)
                    Console.WriteLine(resp.FileName);
           


           

        }


    }
}
