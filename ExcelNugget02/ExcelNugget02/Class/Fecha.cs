﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNugget02.Class
{
  public  class Fecha
    {
        public Task<string> FechaNow() {
            return Task.FromResult($"Fecha : {DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}");
        }

    }
}
