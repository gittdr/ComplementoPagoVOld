using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ComplementoPagoVOld
{
    public partial class Inicial : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            FechaActual();
        }

        public void FechaActual()
        {
            lblFecha.Text = DateTime.Now.ToString("dd/MM/yyyy");
        }
    }
}