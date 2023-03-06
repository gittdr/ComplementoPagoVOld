using ComplementoPagoVOld.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ComplementoPagoVOld
{
    public partial class ListadoRCP : System.Web.UI.Page
    {
        public FacCpController facLabControler = new FacCpController();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                //cargaFacturas();
                cargaFacturas();
                FoliosP();


            }
        }
        private void cargaFacturas()
        {

            DataTable cargaStops = facLabControler.Reporte();
            //cargaStops.AsDataView().RowFilter("");
            int numCells = 1;
            int rownum = 0;
            //cargaStops = cargaStops.Orde
            foreach (DataRow row in cargaStops.Rows)
            {
                TableRow r = new TableRow();
                for (int i = 0; i < numCells; i++)
                {
                    if (i == 0)
                    {
                        HyperLink hp1 = new HyperLink();
                        hp1.ID = "hpIndex" + rownum.ToString();
                        hp1.Text = "<button type='button' class='btn btn-primary btn-lg'>" + row[i].ToString() + "</button>";
                        hp1.NavigateUrl = "DetalleReporte.aspx?folio=" + row[i].ToString();
                        TableCell c = new TableCell();
                        c.Controls.Add(hp1);
                        r.Cells.Add(c);

                    }
                    else
                    {
                        TableCell c = new TableCell();
                        c.Controls.Add(new LiteralControl("row "
                            + rownum.ToString() + ", cell " + i.ToString()));
                        c.Text = row[i].ToString();
                        r.Cells.Add(c);
                    }
                }


                tablaStops.Rows.Add(r);
                rownum++;

            }

        }

        private void FoliosP()
        {

            DataTable cargaStops = facLabControler.ReporteP();
            //cargaStops.AsDataView().RowFilter("");
            int numCells = 3;
            int rownum = 0;
            //cargaStops = cargaStops.Orde
            foreach (DataRow row in cargaStops.Rows)
            {
                TableRow r = new TableRow();
                for (int i = 0; i < numCells; i++)
                {
                    if (i == 0)
                    {
                        HyperLink hp1 = new HyperLink();
                        hp1.ID = "hpIndex" + rownum.ToString();
                        hp1.Text = "<button type='button' class='btn btn-danger'>" + row[i].ToString() + "</button>";
                        //hp1.NavigateUrl = "DetalleReporte.aspx?folio=" + row[i].ToString();
                        TableCell c = new TableCell();
                        c.Controls.Add(hp1);
                        r.Cells.Add(c);

                    }
                    else
                    {
                        TableCell c = new TableCell();
                        c.Controls.Add(new LiteralControl("row "
                            + rownum.ToString() + ", cell " + i.ToString()));
                        c.Text = row[i].ToString();
                        r.Cells.Add(c);
                    }
                }


                Table1.Rows.Add(r);
                rownum++;

            }

        }
        //private async Task okTralix()
        //{
        //    DataTable cargaStops = facLabControler.Reporte();
        //    int numCells = 1;
        //    int rownum = 0;
        //    foreach (DataRow item in cargaStops.Rows)
        //    {
        //        string folio = item["folio"].ToString();
        //        //string rrfolio = "41136";
        //        string idreceptor = item["idreceptor"].ToString().Trim();
        //        string uf = idreceptor + folio;


        //                TableRow r = new TableRow();
        //                for (int i = 0; i < numCells; i++)
        //                {
        //                    if (i == 0)
        //                    {
        //                        HyperLink hp1 = new HyperLink();
        //                        hp1.ID = "hpIndex" + rownum.ToString();
        //                        hp1.Text = "<button type='button' class='btn btn-primary'>" + item[i].ToString() + "</button>";
        //                        hp1.NavigateUrl = "DetallesComplemento.aspx?factura=" + item[i].ToString();
        //                        TableCell c = new TableCell();
        //                        c.Controls.Add(hp1);
        //                        r.Cells.Add(c);

        //                    }
        //                    else
        //                    {
        //                        TableCell c = new TableCell();
        //                        c.Controls.Add(new LiteralControl("row "
        //                            + rownum.ToString() + ", cell " + i.ToString()));
        //                        c.Text = item[i].ToString();
        //                        r.Cells.Add(c);
        //                    }
        //                }


        //                tablaStops.Rows.Add(r);
        //                rownum++;

        //            //FIN


        //    }
        //}
    }
}