using ComplementoPagoVOld.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ComplementoPagoVOld.Controllers
{
    public class FacCpController : Controller
    {
        // GET: FacCp
        public ModelFact modelFact = new ModelFact();
        public DataTable detalleFacturas(string fact)
        {
            return this.modelFact.getDatosFacturas(fact);//this.modelFact.getDatosFacturas(fact);
        }
        public DataTable getDatosCPAGDOC(string identificador)
        {
            return this.modelFact.getDatosCPAGDOC(identificador);
        }
        public DataTable Reporte()
        {
            return this.modelFact.Reporte();
        }
        public DataTable ReporteP()
        {
            return this.modelFact.ReporteP();
        }
        public DataTable ReporteDetalle(string folio)
        {
            return this.modelFact.ReporteDetalle(folio);
        }
        public DataTable getDatosCPAGDOCTR(string identificador)
        {
            return this.modelFact.getDatosCPAGDOCTR(identificador);
        }
        public DataTable getDatosCPAGDOCTRL(string identificador, string foliocpag)
        {
            return this.modelFact.getDatosCPAGDOCTRL(identificador, foliocpag);
        }
        public DataTable getDatosMaster(string identificador)
        {
            return this.modelFact.getDatosMaster(identificador);
        }
        public DataTable facturasEnviadas()
        {
            return this.modelFact.getFacturasEnviadas();
        }
        public DataTable facturas()
        {
            return this.modelFact.getFacturas();
        }
        public void Elist(string identificador)
        {
            this.modelFact.Elist(identificador);
        }
        public DataTable Elist2(string identificador)
        {
            return this.modelFact.Elist2(identificador);
        }
        public DataTable facturasSinRfc()
        {
            return this.modelFact.getFacturasSinRfc();
        }
        // GET: FacCp
       
    }
}