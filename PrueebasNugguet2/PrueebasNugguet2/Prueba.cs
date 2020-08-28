using ExcelNugget02;
using System;
using System.Collections.Generic;
using System.Text;

namespace PruebaExcelFormat
{
    public class Prueba
    {
        [DescripcionExcel(Name = "PLANUM", Ignore = false)]
        public string PLANUM { get; set; }
        [DescripcionExcel(Name = "PERIODO", Ignore = false)]
        public string PERIODO { get; set; }
        [DescripcionExcel(Name = "NIT", Ignore = false)]
        public string NIT { get; set; }
        [DescripcionExcel(Name = "RAZON_SOCIAL", Ignore = false)]
        public string RAZON_SOCIAL { get; set; }
        [DescripcionExcel(Name = "ID_SUCURSAL", Ignore = false)]
        public string ID_SUCURSAL { get; set; }
        [DescripcionExcel(Name = "EMPLEADOS_DECLARADOS", Ignore = false)]
        public string EMPLEADOS_DECLARADOS { get; set; }
        [DescripcionExcel(Name = "MONTO_TOTAL", Ignore = false)]
        public string MONTO_TOTAL { get; set; }
        [DescripcionExcel(Name = "ARCHIVO", Ignore = false)]
        public string ARCHIVO { get; set; }
        [DescripcionExcel(Name = "NPE", Ignore = false)]
        public string NPE { get; set; }
        [DescripcionExcel(Name = "FECHA_ADICION", Ignore = false)]
        public string FECHA_ADICION { get; set; }
        [DescripcionExcel(Name = "FECHA_PRESENTACION", Ignore = false)]
        public string FECHA_PRESENTACIÓN { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CATEGORIA { get; set; }
 
    }
    public class Prueba2
    {
        [DescripcionExcel(Name = "ID PLANILLA", Ignore = false)]
        public string ID_PLANILLA { get; set; }
        [DescripcionExcel(Name = "ESTADO", Ignore = false)]
        public string ESTADO { get; set; }
        [DescripcionExcel(Name = "NIT", Ignore = false)]
        public string NIT { get; set; }
        [DescripcionExcel(Name = "EMPRESA", Ignore = false)]
        public string EMPRESA { get; set; }
        [DescripcionExcel(Name = "CATEGORIA", Ignore = false)]
        public string CATEGORIA { get; set; }
        [DescripcionExcel(Name = "ASESOR", Ignore = false)]
        public string ASESOR { get; set; }
        [DescripcionExcel(Name = "GESTOR EMPRESARIAL", Ignore = false)]
        public string GESTOR_EMPRESARIAL { get; set; }
        [DescripcionExcel(Name = "PERIODO", Ignore = false)]
        public int PERIODO { get; set; }
        [DescripcionExcel(Name = "ID SUCURSAL", Ignore = false)]
        public string ID_SUCURSAL { get; set; }
        [DescripcionExcel(Name = "EMPLEADOS DECLARADOS", Ignore = false)]
        public int EMPLEADOS_DECLARADOS { get; set; }
        [DescripcionExcel(Name = "MONTO TOTAL", Ignore = false)]
        public string MONTO_TOTAL { get; set; }
        [DescripcionExcel(Name = "FECHA ADICION", Ignore = false)]
        public string FECHA_ADICION { get; set; }
        [DescripcionExcel(Name = "FECHA ADICION2", Ignore = false)]
        public string FECHA_ADICCION2 { get; set; }
        [DescripcionExcel(Name = "ARCHIVO", Ignore = false)]
        public string ARCHIVO { get; set; }
        [DescripcionExcel(Name = "NPE", Ignore = false)]
        public string NPE { get; set; }
        [DescripcionExcel(Name = "Correos", Ignore = false)]
        public string CORREOS { get; set; }
    }
}
