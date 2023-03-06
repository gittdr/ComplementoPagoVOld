using ComplementoPagoVOld.Controllers;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ComplementoPagoVOld
{
    public partial class Ccp : System.Web.UI.Page
    {
        public facLabController facLabControler = new facLabController();
        //public FacCpController facLabControler = new FacCpController();

        public string fDesde, fHasta, concepto, tipoCobro, tipocomprobante, lugarexpedicion, metodopago33, formadepago, usocfdi, confirmacion, paisresidencia, numtributacion
        , mailenvio, numidentificacion, claveunidad, tipofactoriva, tipofactorret, coditrans, tipofactor, tasatras, codirete, tasarete, relacion, montosoloiva, montoivarete
        , ivadeiva, ivaderet, retderet, conceptoretencion, consecutivoconcepto, claveproductoservicio, valorunitario, importe, descuento, cantidadletra, uuidrel
        , identificador, version, fechapago, monedacpag, tipodecambiocpag, monto, numerooperacion, rfcemisorcuenta, nombrebanco, numerocuentaord, rfcemisorcuentaben, numcuentaben
        , tipocadenapago, certpago, cadenadelpago, sellodelpago, identpag, identdocpago, seriecpag, foliocpag, monedacpagdoc, tipocambiocpag, metododepago, numerodeparcialidad
        , importeSaldoAnterior, importepago, importesaldoinsoluto, total, subt, ivat, rett, cond, tipoc, seriee, folioe, sfolio, Foliosrelacionados, serier, folior, uuidpagadas, IdentificadorDelDocumentoPagado, ipagado, nparcialidades, folio, MetdodoPago, Dserie, monedascpadgoc, interiorsaldoanterior, isaldoinsoluto, identificaciondpago, folioscpag, k1, k3, norden, tmoneda, idcomprobante, cantidad, descripcion, Tuuid, iddelpago, iipagado, basecalculado, basecalculado2, basecalculado3, impSaldoAnterior, impSaldoInsoluto, fechap;

        public bool error = false;

        public string serie;
        public decimal importePagos = 0;
        public decimal importePagos2 = 0;
        public decimal importePagos3 = 0;
        public decimal importePagos4 = 0;
        public decimal importePagos5 = 0;
        public decimal importePagos7 = 0;
        public decimal importePagos22 = 0;
        public decimal importePagos23 = 0;
        public decimal importePagos24 = 0;
        public decimal importePagos25 = 0;
        public decimal importePagos26 = 0;
        public decimal valorunitarios = 0;


        public double ivaa = 0.16;
        public double isrr = 0.04;
        public decimal totalIva = 0;
        public decimal totalIsr = 0;
        public int serietsrl = 0;
        public string ejecutar = "Si";
        public decimal basecalculo = 0;
        public decimal basecalculo2 = 0;
        public decimal basecalculo3 = 0;

        public string uid = "";



        public int contadorPUE = 0;
        public int contadorPPD = 0;
        public int contadortralix = 0;


        string cpagdoc = "";
        public string escrituraFactura = "", idSucursal = "", idTipoFactura = "", jsonFactura = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            imgFDesde.Visible = false;
            imgFHasta.Visible = false;
            //lblFact.Text = Request.QueryString["factura"];
            lblFact.Text = "40979";


            //foliot = Request.QueryString["factura"];
            if (IsPostBack)
            {
                fDesde = txtFechaDesde.Text;
                fHasta = txtFechaHasta.Text;
                concepto = txtConcepto.Text;
                tipoCobro = txtTipoCobro.Text;
                formadepago = txtFormaPago.Text;
                txtFolio.Text = "";
            }



            iniciaDatos();

        }
        public async Task iniciaDatos()
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }

            DataTable td = facLabControler.detalleFacturas(lblFact.Text);
            Div1.Visible = false;
            //Obtencion de datos------------------------------------------------------------------------------------------------------------------------ -

            foreach (DataRow row in td.Rows)
            {


                //01-------------------------------------------------------------------------------------------------------------------------
                if (txtFolio.Text != row["SFolio"].ToString())
                {
                    txtFechaIniOP.Text = txtFechaIniOP.Text + "\r\n" + row["IdentificadorDelDocumentoPagado"].ToString();
                    FolioUUIDTxt.Text = row["UUIDident"].ToString();
                    iddelpago = row["Folio"].ToString();
                    txtFolio.Text = row["SFolio"].ToString();
                    DateTime dt = DateTime.Parse(row["FechaHoraEmision"].ToString());
                    txtFechaFactura.Text = dt.ToString("yyyy'/'MM'/'dd HH:mm:ss");

                    sfolio = row["SFolio"].ToString();


                    seriee = row["Serie"].ToString();

                    folioe = row["Folio"].ToString();
                    subt = row["Subtotal"].ToString();
                    ivat = row["TotalImpuestosTrasladados"].ToString();
                    rett = row["TotalImpuestosRetenidos"].ToString();
                    total = row["Total"].ToString();
                    cantidadletra = row["Totalconletra"].ToString();
                    //formadepago = row["FormaDePago"].ToString();
                    cond = row["CondicionesdePago"].ToString();
                    metodopago33 = row["MetodoPago"].ToString();
                    txtMoneda.Text = row["Moneda"].ToString();
                    tipoc = row["Tipodecambio"].ToString();
                    tipocomprobante = row["TipodeComprobante"].ToString();
                    lugarexpedicion = row["LugardeExpedición"].ToString();
                    usocfdi = row["UsoCFDI"].ToString();
                    confirmacion = row["Confirmación"].ToString();

                    //02-------------------------------------------------------------------------------------------------------------------------

                    txtIdCliente.Text = row["IdReceptor"].ToString();

                    txtRFC.Text = row["RFC"].ToString();
                    //RFC = row["RFC"].ToString();
                    txtCliente.Text = row["Nombre"].ToString();
                    txtPaís.Text = row["Pais"].ToString();
                    txtCalle.Text = row["Calle"].ToString();
                    txtNoExt.Text = row["NumeroExterior"].ToString();
                    txtNoInt.Text = row["NumeroInterior"].ToString();
                    txtColonia.Text = row["Colonia"].ToString();
                    txtLocalidad.Text = row["Localidad"].ToString();
                    txtReferencia.Text = row["Referencia"].ToString();
                    txtMunicipio.Text = row["MunicipioDelegacion"].ToString();
                    txtEstado.Text = row["Estado"].ToString();
                    txtCP.Text = row["CódigoPostal"].ToString();
                    txtFechaPago.Text = row["Fechapago"].ToString();
                    paisresidencia = row["PaísResidenciaFiscal"].ToString();
                    numtributacion = row["NúmeroDeRegistroIdTributacion"].ToString();
                    mailenvio = row["CorreoEnvio"].ToString();

                    //04-------------------------------------------------------------------------------------------------------------------------

                    consecutivoconcepto = row["ConsecutivoConcepto"].ToString();
                    claveproductoservicio = row["ClaveProductooServicio"].ToString();
                    numidentificacion = row["NumeroIdentificación"].ToString();
                    claveunidad = row["ClaveUnidad"].ToString();
                    txtUnidadMedida.Text = row["ClaveUnidad"].ToString();
                    txtIdConcepto.Text = row["ClaveProductooServicio"].ToString();
                    txtCantidad.Text = row["Cantidad"].ToString();
                    txtMetodoPago.Text = row["MedotoDePago"].ToString();

                    if (concepto == null || concepto.Equals(row["Descripcion"].ToString())) { txtConcepto.Text = row["Descripcion"].ToString(); }
                    else { txtConcepto.Text = concepto; }


                    if (formadepago == null || formadepago.Equals(row["Formadepagocpag"].ToString())) { txtFormaPago.Text = row["Formadepagocpag"].ToString(); }
                    else { txtFormaPago.Text = formadepago; }


                    valorunitario = row["ValorUnitario"].ToString();
                    importe = row["Importe"].ToString();
                    descuento = row["Descuento"].ToString();

                    //CPAG-------------------------------------------------------------------------------------------------------------------------


                    DateTime dtdtt = DateTime.Parse(row["Fechapago"].ToString());
                    fechapago = dtdtt.ToString("yyyy'-'MM'-'dd'T'HH:mm:ss");
                    DataTable ctipocambio = facLabControler.getTipoCambio(fechapago);
                    foreach (DataRow tcambio in ctipocambio.Rows)
                    {
                        tipodecambiocpag = tcambio["XCHGRATE"].ToString();
                    }
                    //fechapago =
                    identificador = row["Identificador"].ToString();
                    version = row["version"].ToString();
                    //txtFormaPago.Text = row["Formadepagocpag"].ToString();
                    monedacpag = row["Monedacpag"].ToString();
                    tipodecambiocpag = row["TipoDeCambiocpag"].ToString();
                    monto = row["Monto"].ToString();
                    numerooperacion = row["NumeroOperacion"].ToString();
                    txtRFCbancoEmisor.Text = row["RFCEmisorCuentaBeneficiario"].ToString();
                    txtBancoEmisor.Text = row["NombreDelBanco"].ToString();
                    txtCuentaPago.Text = row["NumeroCuentaOrdenante"].ToString();
                    rfcemisorcuentaben = row["RFCEmisorCuentaBeneficario"].ToString();
                    numcuentaben = row["NumerCuentaBeneficiario"].ToString();
                    tipocadenapago = row["TipoCadenaPago"].ToString();
                    certpago = row["CertificadoPago"].ToString();
                    cadenadelpago = row["CadenaDePago"].ToString();
                    sellodelpago = row["SelloDePago"].ToString();


                    if (txtRFC.Text != "")
                    {
                        //DataTable detalleIdent2 = facLabControler.getDatosCPAGDOC(row["IdentificadorDelPago"].ToString());
                        //if (detalleIdent2.Rows.Count > 0)
                        //{
                        //    foreach (DataRow rowIdent2 in detalleIdent2.Rows)
                        //    {
                        //        identificaciondpago = rowIdent2["IdentificadorDelPago"].ToString();
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("SM-", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("A", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("B", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("C", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("D", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("ND", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace(".", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("-", "").Trim(), @"[A-Z]", "");
                        //        folioscpag = Regex.Replace(rowIdent2["Foliocpag"].ToString().Replace("NS", "").Trim(), @"[A-Z]", "");
                        //        string folioss = folioscpag;
                        //        int folii = Int32.Parse(folioss);
                        //        string urls = "https://canal1.xsa.com.mx:9050/bf2e1036-ba47-49a0-8cd9-e04b36d5afd4/cfdis?folioEspecifico="+folii;
                        //        string url = urls.Trim();


                        //        var request2819a = (HttpWebRequest)WebRequest.Create("https://canal1.xsa.com.mx:9050/bf2e1036-ba47-49a0-8cd9-e04b36d5afd4/cfdis?folioEspecifico="+folii+"&rfc="+txtRFC.Text);
                        //        var response2819a = (HttpWebResponse)request2819a.GetResponse();
                        //        var responseString2819a = new StreamReader(response2819a.GetResponseStream()).ReadToEndAsync();
                        //        //WebRequest request2819 = (HttpWebRequest)WebRequest.Create(url);
                        //        //WebResponse response2819 = (HttpWebResponse)request2819.GetResponse();
                        //        //Stream dataStream = response2819.GetResponseStream();
                        //        //StreamReader responseString2819 = new StreamReader(dataStream);
                        //        //string rt = responseString2819.ReadToEnd();


                        //    }
                        //}
                    }




                }


            }



        }


        public static bool Tralix(string folio)
        {


            return true;
        }

        public void generaTXT()
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }
            txtConcepto.Text = validaCampo(txtConcepto.Text.Trim());

            string path = System.Web.Configuration.WebConfigurationManager.AppSettings["dir"] + lblFact.Text + ".txt";
            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(path))



            {
                string f1 = txtFechaDesde.Text;
                string factura = txtFolio.Text;


                //----------------------------------------Seccion De Datos Generales del CFDI-----------------------------------------------------------------------------------------

                //01 INFORMACION GENERAL DEL CFDI (1:1)
                if (formadepago.Trim() != "02")
                {
                    escritor.WriteLine(
                    "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                         //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|"
                    );
                    escrituraFactura += "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                        //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|";
                }
                else
                {
                    escritor.WriteLine(
                    "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                      //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|"
                    );
                    escrituraFactura += "01"                                                //1.-Tipo De Registro
                   + "|" + sfolio                                      //2-ID Comprobante
                   + "|" + seriee                                      //3-Serie
                   + "|" + folioe                                      //4-Foliio 
                   + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                   + "|" + subt                                        //6-Subtotal
                   + "|" + ivat                                        //7-Total Impuestos Trasladados
                   + "|" + rett                                        //8-Total Impuestos Retenidos
                   + "|"                                               //9-Descuentos
                   + "|" + "0"                                       //10-Total
                   + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                   + "|"                          //12-Forma De Pago
                   + "|" + cond                                        //13-Condiciones De Pago
                   + "|"                                 //14-Metodo de Pago
                   + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                   + "|" + tipoc                                       //16-Tipo De Cambio
                   + "|" + tipocomprobante                             //17-Tipo De Comprobante
                   + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                   + "|" + usocfdi                                     //19-Uso CFDI
                   + "|" + confirmacion                                //20-Confirmacion
                   + "|";
                }
                //----------------------------------------Seccion de los datos del receptor del CFDI -------------------------------------------------------------------------------------

                //02 INFORMACION DEL RECEPTOR (1:1)
                if (monedascpadgoc.Trim() == "USD")
                {
                    escritor.WriteLine(
                       "02"                                                   //1-Tipo De Registro
                       + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                       + "|" + txtRFC.Text.Trim()                                //3-RFC
                       + "|" + txtCliente.Text.Trim()                         //4-Nombre
                       + "|" + txtPaís.Text.Trim()                            //5-Pais
                       + "|" + txtCalle.Text.Trim()                           //6-Calle
                       + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                       + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                       + "|" + txtColonia.Text.Trim()                         //9-Colonia
                       + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                       + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                       + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                       + "|" + txtEstado.Text.Trim()                          //13-EStado
                       + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                       + "|"                                                // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                       + "|"                                   //16-Numero de Registro de ID Tributacion 
                       + "|" + mailenvio                                      //17-Correo de envio                                                    
                       + "|"                                                  //Fin Del Registro 
                       );

                    escrituraFactura += "\\n02"                                                   //1-Tipo De Registro
                       + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                       + "|" + txtRFC.Text.Trim()                                //3-RFC
                       + "|" + txtCliente.Text.Trim()                         //4-Nombre
                       + "|" + txtPaís.Text.Trim()                            //5-Pais
                       + "|" + txtCalle.Text.Trim()                           //6-Calle
                       + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                       + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                       + "|" + txtColonia.Text.Trim()                         //9-Colonia
                       + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                       + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                       + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                       + "|" + txtEstado.Text.Trim()                          //13-EStado
                       + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                       + "|"                                          // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                       + "|"                                   //16-Numero de Registro de ID Tributacion 
                       + "|" + mailenvio                                      //17-Correo de envio                                                    
                       + "|";

                }
                else
                {

                    escritor.WriteLine(
                    "02"                                                   //1-Tipo De Registro
                    + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                    + "|" + txtRFC.Text.Trim()                             //3-RFC
                    + "|" + txtCliente.Text.Trim()                         //4-Nombre
                    + "|" + txtPaís.Text.Trim()                            //5-Pais
                    + "|" + txtCalle.Text.Trim()                           //6-Calle
                    + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                    + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                    + "|" + txtColonia.Text.Trim()                         //9-Colonia
                    + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                    + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                    + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                    + "|" + txtEstado.Text.Trim()                          //13-EStado
                    + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                    + "|" + "" // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                    + "|" + numtributacion                                 //16-Numero de Registro de ID Tributacion 
                    + "|" + mailenvio                                      //17-Correo de envio                                                    
                    + "|"                                                  //Fin Del Registro 
                    );

                    escrituraFactura += "\\n02"                                                   //1-Tipo De Registro
                    + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                    + "|" + txtRFC.Text.Trim()                             //3-RFC
                    + "|" + txtCliente.Text.Trim()                         //4-Nombre
                    + "|" + txtPaís.Text.Trim()                            //5-Pais
                    + "|" + txtCalle.Text.Trim()                           //6-Calle
                    + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                    + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                    + "|" + txtColonia.Text.Trim()                         //9-Colonia
                    + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                    + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                    + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                    + "|" + txtEstado.Text.Trim()                          //13-EStado
                    + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                    + "|" + "" // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                    + "|" + numtributacion                                 //16-Numero de Registro de ID Tributacion 
                    + "|" + mailenvio                                      //17-Correo de envio                                                    
                    + "|";

                }

                //----------------------------------------Seccion de detalles del complemento de pago -------------------------------------------------------------------

                //04 INFORMACION DE LOS CONCEPTOS (1:N)
                escritor.WriteLine(
               "04"                                                   //1-Tipo De Registro
               + "|" + consecutivoconcepto.Trim()                            //2-Consecutivo Concepto
               + "|" + claveproductoservicio.Trim()                          //3-Clave Producto o Servicio SAT                                               
               + "|" + numidentificacion.Trim()                              //4-Numero Identificacion TDR
               + "|" + txtCantidad.Text.Trim()                        //5-Cantidad
               + "|" + claveunidad.Trim()                                    //6-Clave Unidad SAT
               + "|"                                                  //7-Unidad de Medida
               + "|" + txtConcepto.Text.Trim()                                //8-Descripcion
               + "|" + "0"                                  //9-Valor Unitario
               + "|" + "0"                                        //10-Importe
               + "|" + descuento.Trim()                                      //11-Descuento                                                  
                                                                             //12 Importe con iva si el rfc es XAXX010101000 y XEXX010101000 OPCIONAL
               + "|"                                                  //Fin Del Registro
                );

                escrituraFactura += "\\n04"                                                   //1-Tipo De Registro
               + "|" + consecutivoconcepto.Trim()                            //2-Consecutivo Concepto
               + "|" + claveproductoservicio.Trim()                          //3-Clave Producto o Servicio SAT                                               
               + "|" + numidentificacion.Trim()                             //4-Numero Identificacion TDR
               + "|" + txtCantidad.Text.Trim()                        //5-Cantidad
               + "|" + claveunidad.Trim()                                   //6-Clave Unidad SAT
               + "|"                                                  //7-Unidad de Medida
               + "|" + txtConcepto.Text.Trim()                                //8-Descripcion
               + "|" + "0"                                  //9-Valor Unitario
               + "|" + "0"                                        //10-Importe
               + "|" + descuento.Trim()                                      //11-Descuento                                                  
                                                                             //12 Importe con iva si el rfc es XAXX010101000 y XEXX010101000 OPCIONAL
               + "|";

                //----------------------------------------Seccion CPAG20 -------------------------------------------------------------------

                //CPAG20 (1:1)
                //escritor.WriteLine(
                //"CPAG20"                         //1-Tipo De Registro
                //+ "|" + "2.0"                    //2-Version
                //+ "|"                               //Fin Del Registro
                //);

                //escrituraFactura += "CPAG20"    //1-Tipo De Registro
                //+ "|"  + "2.0"                   //2-Version  
                //+ "|";		   

                //----------------------------------------Seccion CPAG20TOT -------------------------------------------------------------------

                //CPAG20TOT (1:1)
                //escritor.WriteLine(
                //"CPAG20TOT"                         //1-Tipo De Registro
                //+ "|"                               //2-TotalRetencionesIVA
                //+ "|"                               //3-TotalRetencionesISR                                              
                //+ "|"                               //4-TotalRetencionesIEPS
                //+ "|"                               //5-TotalTrasladosBaseIVA16
                //+ "|"                               //6-TotalTrasladosImpuestoIVA16
                //+ "|"                               //7-TotalTrasladosBaseIVA8
                //+ "|"                               //8-TotalTrasladosImpuestoIVA8
                //+ "|"                               //9-TotalTrasladosBaseIVA0
                //+ "|"                               //10-TotalTrasladosImpuestoIVA0
                //+ "|"                               //11-TotalTrasladosBaseIVAExento
                // + "|" + monto                       //12-MontoTotalPagos                                                                                                 
                // + "|"                               //Fin Del Registro
                //);

                //escrituraFactura += "CPAG20TOT"    //1-Tipo De Registro
                //+ "|"                               //2-TotalRetencionesIVA
                //+ "|"                               //3-TotalRetencionesISR                                              
                //+ "|"                               //4-TotalRetencionesIEPS
                //+ "|"                               //5-TotalTrasladosBaseIVA16
                //+ "|"                               //6-TotalTrasladosImpuestoIVA16
                //+ "|"                               //7-TotalTrasladosBaseIVA8
                //+ "|"                               //8-TotalTrasladosImpuestoIVA8
                //+ "|"                               //9-TotalTrasladosBaseIVA0
                //+ "|"                               //10-TotalTrasladosImpuestoIVA0
                //+ "|"                               //11-TotalTrasladosBaseIVAExento
                //+ "|" + monto                       //12-MontoTotalPagos  
                //+ "|";		   
                //----------------------------------------Seccion CPAG20PAGO-------------------------------------------------------------------------------------------------

                //CPAG20PAGO COMPLEMENTO DE PAGO (1:N)
                //escritor.WriteLine(
                //"CPAG20PAGO"                                           //1-Tipo De Registro
                //+ "|" + identificador                                  //2-Identificador                                             
                //+ "|" + fechapago                                      //3-Fechapago
                //+ "|" + txtFormaPago.Text                              //4-Formadepagocpag
                //+ "|" + monedacpag                                     //5-Monedacpag
                //+ "|" + tipodecambiocpag                               //6-TipoDecambiocpag
                //+ "|" + txtTotal.Text                                  //8-Monto
                //+ "|" + numerooperacion                                //9-NumeroOperacion
                //+ "|" + txtRFCbancoEmisor.Text                         //10-RFCEmisorCuentaBeneficiario
                //+ "|" + txtBancoEmisor.Text                            //11-NombreDelBanco                                                                                            
                //+ "|" + txtCuentaPago.Text                             //12-NumeroCuentaOrdenante
                //+ "|" + rfcemisorcuentaben                             //13-RFCEmisorCuentaBeneficiario
                //+ "|" + numcuentaben                                   //14-NumCuentaBeneficiario
                //+ "|" + tipocadenapago                                 //15-TipoCadenaPago                                               
                //+ "|" + certpago                                       //16-CertificadoPago
                //+ "|" + cadenadelpago                                  //17-CadenaDePago
                //+ "|" + sellodelpago                                   //Fin Del Registro
                //+ "|"
                //);

                //escrituraFactura += "CPAG20PAGO"                      //1-Tipo De Registro
                //+ "|" + identificador                                  //2-Identificador                                             
                //+ "|" + fechapago                                      //3-Fechapago
                //+ "|" + txtFormaPago.Text                              //4-Formadepagocpag
                //+ "|" + monedacpag                                     //5-Monedacpag
                //+ "|" + tipodecambiocpag                               //6-TipoDecambiocpag
                //+ "|" + txtTotal.Text                                  //8-Monto
                //+ "|" + numerooperacion                                //9-NumeroOperacion
                //+ "|" + txtRFCbancoEmisor.Text                         //10-RFCEmisorCuentaBeneficiario
                //+ "|" + txtBancoEmisor.Text                            //11-NombreDelBanco                                                                                            
                //+ "|" + txtCuentaPago.Text                             //12-NumeroCuentaOrdenante
                //+ "|" + rfcemisorcuentaben                             //13-RFCEmisorCuentaBeneficiario
                //+ "|" + numcuentaben                                   //14-NumCuentaBeneficiario
                //+ "|" + tipocadenapago                                 //15-TipoCadenaPago                                               
                //+ "|" + certpago                                       //16-CertificadoPago
                //+ "|" + cadenadelpago                                  //17-CadenaDePago
                //+ "|" + sellodelpago                                   //Fin Del Registro
                //+ "|";


                //----------------------------------------Seccion CPAG-------------------------------------------------------------------------------------------------

                //CPAG COMPLEMENTO DE PAGO (1:N)
                escritor.WriteLine(
               "CPAG"                                                 //1-Tipo De Registro
               + "|" + identificador.Trim()                                  //2-Identificador
               + "|" + version.Trim()                                        //3-Version                                             
               + "|" + fechapago.Trim()                                      //4-Fechapago
               + "|" + formadepago.Trim()                              //5-Formadepagocpag
               + "|" + monedascpadgoc.Trim()                                     //6-Monedacpag
               + "|" + tipodecambiocpag.Trim()                               //7-TipoDecambiocpag AQUI LO VOY A TOMAR DE OTRA CONSULTA
               + "|" + txtTotal.Text.Trim()                                 //8-Monto
               + "|" + numerooperacion.Trim()                               //9-NumeroOperacion
               + "|" + txtRFCbancoEmisor.Text.Trim()                         //10-RFCEmisorCuentaBeneficiario
               + "|" + txtBancoEmisor.Text.Trim()                            //11-NombreDelBanco                                                                                            
               + "|" + txtCuentaPago.Text.Trim()                            //12-NumeroCuentaOrdenante
               + "|" + rfcemisorcuentaben.Trim()                           //13-RFCEmisorCuentaBeneficiario
               + "|" + numcuentaben.Trim()                                //14-NumCuentaBeneficiario
               + "|" + tipocadenapago.Trim()                                //15-TipoCadenaPago                                               
               + "|" + certpago.Trim()                                    //16-CertificadoPago
               + "|" + cadenadelpago.Trim()                                //17-CadenaDePago
               + "|" + sellodelpago.Trim()                                  //Fin Del Registro
               + "|"
                );

                escrituraFactura += "CPAG"                                                 //1-Tipo De Registro
               + "|" + identificador.Trim()                                 //2-Identificador
               + "|" + version.Trim()                                   //3-Version                                             
               + "|" + fechapago.Trim()                                     //4-Fechapago
               + "|" + formadepago.Trim()                              //5-Formadepagocpag
               + "|" + monedascpadgoc.Trim()                                    //6-Monedacpag
               + "|" + tipodecambiocpag.Trim()                              //7-TipoDecambiocpag
               + "|" + txtTotal.Text.Trim()                                 //8-Monto
               + "|" + numerooperacion.Trim()                               //9-NumeroOperacion
               + "|" + txtRFCbancoEmisor.Text.Trim()                        //10-RFCEmisorCuentaBeneficiario
               + "|" + txtBancoEmisor.Text.Trim()                          //11-NombreDelBanco                                                                                            
               + "|" + txtCuentaPago.Text.Trim()                           //12-NumeroCuentaOrdenante
               + "|" + rfcemisorcuentaben.Trim()                            //13-RFCEmisorCuentaBeneficiario
               + "|" + numcuentaben.Trim()                                //14-NumCuentaBeneficiario
               + "|" + tipocadenapago.Trim()                                //15-TipoCadenaPago                                               
               + "|" + certpago.Trim()                                  //16-CertificadoPago
               + "|" + cadenadelpago.Trim()                                 //17-CadenaDePago
               + "|" + sellodelpago.Trim()                                 //Fin Del Registro
               + "|";

                //----------------------------------------Seccion CPAGDOC------------------------------------------------------------------------------------------------

                //CPAG COMPLEMENTO DE PAGO (1:N)
                escritor.WriteLine(cpagdoc
                //"CPAGDOC"                                              //1-Tipo De Registro
                //+ "|" + identpag                                       //2-IdentificadorDelPago
                //+ "|" + txtFechaIniOP.Text                             //3-IdentificadorDelDocumentoPagado                                              
                //+ "|" + seriecpag                                      //4-Seriecpag
                //+ "|" + foliocpag                                      //5-Foliocpag
                //+ "|" + monedacpagdoc                                  //6-Monedacpag
                //+ "|" + tipocambiocpag                                 //7-TipoCambiocpagdpc
                //+ "|" + txtMetodoPago.Text                             //8-MetodoDePago
                //+ "|" + numerodeparcialidad                            //9-NumeroDeParcialidad
                //+ "|" + importeSaldoAnterior                           //10-ImporteSaldoAnterior
                //+ "|" + importepago                                    //11-ImportePagado                                                  
                //+ "|" + importesaldoinsoluto                           //12 ImporteSaldoInsoluto
                //+ "|"                                                  //Fin Del Registro
                );


                escrituraFactura += cpagdoc;
                //escrituraFactura = escrituraFactura.Replace("||02|", "||\\n02|");
                //escrituraFactura = escrituraFactura.Replace("||04|", "||\\n04|");
                escrituraFactura = escrituraFactura.Replace("| \r\n", "|");
                escrituraFactura = escrituraFactura.Replace("|CPAG", "|\\nCPAG");

            }
        }
        public void generadorTXT()
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }
            txtConcepto.Text = validaCampo(txtConcepto.Text.Trim());

            string path = System.Web.Configuration.WebConfigurationManager.AppSettings["dir2"] + lblFact.Text + ".txt";
            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(path))



            {
                string f1 = txtFechaDesde.Text;
                string factura = txtFolio.Text;


                //----------------------------------------Seccion De Datos Generales del CFDI-----------------------------------------------------------------------------------------

                //01 INFORMACION GENERAL DEL CFDI (1:1)
                if (formadepago.Trim() != "02")
                {
                    escritor.WriteLine(
                    "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                         //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|"
                    );
                    escrituraFactura += "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                        //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|";
                }
                else
                {
                    escritor.WriteLine(
                    "01"                                                //1.-Tipo De Registro
                    + "|" + sfolio                                      //2-ID Comprobante
                    + "|" + seriee                                      //3-Serie
                    + "|" + folioe                                      //4-Foliio 
                    + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                    + "|" + subt                                        //6-Subtotal
                    + "|" + ivat                                        //7-Total Impuestos Trasladados
                    + "|" + rett                                        //8-Total Impuestos Retenidos
                    + "|"                                               //9-Descuentos
                    + "|" + "0"                                       //10-Total
                    + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                    + "|"                      //12-Forma De Pago
                    + "|" + cond                                        //13-Condiciones De Pago
                    + "|"                                 //14-Metodo de Pago
                    + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                    + "|" + tipoc                                       //16-Tipo De Cambio
                    + "|" + tipocomprobante                             //17-Tipo De Comprobante
                    + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                    + "|" + usocfdi                                     //19-Uso CFDI
                    + "|" + confirmacion                                //20-Confirmacion
                    + "|"
                    );
                    escrituraFactura += "01"                                                //1.-Tipo De Registro
                   + "|" + sfolio                                      //2-ID Comprobante
                   + "|" + seriee                                      //3-Serie
                   + "|" + folioe                                      //4-Foliio 
                   + "|" + txtFechaFactura.Text.Trim()                 //5-Fecha y Hora De Emision
                   + "|" + subt                                        //6-Subtotal
                   + "|" + ivat                                        //7-Total Impuestos Trasladados
                   + "|" + rett                                        //8-Total Impuestos Retenidos
                   + "|"                                               //9-Descuentos
                   + "|" + "0"                                       //10-Total
                   + "|" + cantidadletra.Trim()                        //11-Total Con Letra
                   + "|"                          //12-Forma De Pago
                   + "|" + cond                                        //13-Condiciones De Pago
                   + "|"                                 //14-Metodo de Pago
                   + "|" + txtMoneda.Text.Trim()                       //15-Moneda
                   + "|" + tipoc                                       //16-Tipo De Cambio
                   + "|" + tipocomprobante                             //17-Tipo De Comprobante
                   + "|" + lugarexpedicion                             //18-Lugar De Expedicion                                        
                   + "|" + usocfdi                                     //19-Uso CFDI
                   + "|" + confirmacion                                //20-Confirmacion
                   + "|";
                }
                //----------------------------------------Seccion de los datos del receptor del CFDI -------------------------------------------------------------------------------------

                //02 INFORMACION DEL RECEPTOR (1:1)
                if (monedascpadgoc.Trim() == "USD")
                {
                    escritor.WriteLine(
                       "02"                                                   //1-Tipo De Registro
                       + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                       + "|" + txtRFC.Text.Trim()                                //3-RFC
                       + "|" + txtCliente.Text.Trim()                         //4-Nombre
                       + "|" + txtPaís.Text.Trim()                            //5-Pais
                       + "|" + txtCalle.Text.Trim()                           //6-Calle
                       + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                       + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                       + "|" + txtColonia.Text.Trim()                         //9-Colonia
                       + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                       + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                       + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                       + "|" + txtEstado.Text.Trim()                          //13-EStado
                       + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                       + "|"                                                // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                       + "|"                                   //16-Numero de Registro de ID Tributacion 
                       + "|" + mailenvio                                      //17-Correo de envio                                                    
                       + "|"                                                  //Fin Del Registro 
                       );

                    escrituraFactura += "\\n02"                                                   //1-Tipo De Registro
                       + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                       + "|" + txtRFC.Text.Trim()                                //3-RFC
                       + "|" + txtCliente.Text.Trim()                         //4-Nombre
                       + "|" + txtPaís.Text.Trim()                            //5-Pais
                       + "|" + txtCalle.Text.Trim()                           //6-Calle
                       + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                       + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                       + "|" + txtColonia.Text.Trim()                         //9-Colonia
                       + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                       + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                       + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                       + "|" + txtEstado.Text.Trim()                          //13-EStado
                       + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                       + "|"                                          // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                       + "|"                                   //16-Numero de Registro de ID Tributacion 
                       + "|" + mailenvio                                      //17-Correo de envio                                                    
                       + "|";

                }
                else
                {

                    escritor.WriteLine(
                    "02"                                                   //1-Tipo De Registro
                    + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                    + "|" + txtRFC.Text.Trim()                             //3-RFC
                    + "|" + txtCliente.Text.Trim()                         //4-Nombre
                    + "|" + txtPaís.Text.Trim()                            //5-Pais
                    + "|" + txtCalle.Text.Trim()                           //6-Calle
                    + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                    + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                    + "|" + txtColonia.Text.Trim()                         //9-Colonia
                    + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                    + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                    + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                    + "|" + txtEstado.Text.Trim()                          //13-EStado
                    + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                    + "|" + "" // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                    + "|" + numtributacion                                 //16-Numero de Registro de ID Tributacion 
                    + "|" + mailenvio                                      //17-Correo de envio                                                    
                    + "|"                                                  //Fin Del Registro 
                    );

                    escrituraFactura += "\\n02"                                                   //1-Tipo De Registro
                    + "|" + txtIdCliente.Text.Trim()                       //2-Id Receptor
                    + "|" + txtRFC.Text.Trim()                             //3-RFC
                    + "|" + txtCliente.Text.Trim()                         //4-Nombre
                    + "|" + txtPaís.Text.Trim()                            //5-Pais
                    + "|" + txtCalle.Text.Trim()                           //6-Calle
                    + "|" + txtNoExt.Text.Trim()                           //7-Numero Exterior
                    + "|" + txtNoInt.Text.Trim()                           //8-Numero Interior
                    + "|" + txtColonia.Text.Trim()                         //9-Colonia
                    + "|" + txtLocalidad.Text.Trim()                       //10-Localidad
                    + "|" + txtReferencia.Text.Trim()                      //11-Referencia
                    + "|" + txtMunicipio.Text.Trim()                       //12-Municio/Delegacion
                    + "|" + txtEstado.Text.Trim()                          //13-EStado
                    + "|" + txtCP.Text.Trim()                              //14-Codigo Postal
                    + "|" + "" // paisresidencia                                 //15-Pais de Residecia Fiscal Cuando La Empresa Sea Extrajera
                    + "|" + numtributacion                                 //16-Numero de Registro de ID Tributacion 
                    + "|" + mailenvio                                      //17-Correo de envio                                                    
                    + "|";

                }

                //----------------------------------------Seccion de detalles del complemento de pago -------------------------------------------------------------------

                //04 INFORMACION DE LOS CONCEPTOS (1:N)
                escritor.WriteLine(
               "04"                                                   //1-Tipo De Registro
               + "|" + consecutivoconcepto.Trim()                            //2-Consecutivo Concepto
               + "|" + claveproductoservicio.Trim()                          //3-Clave Producto o Servicio SAT                                               
               + "|" + numidentificacion.Trim()                              //4-Numero Identificacion TDR
               + "|" + txtCantidad.Text.Trim()                        //5-Cantidad
               + "|" + claveunidad.Trim()                                    //6-Clave Unidad SAT
               + "|"                                                  //7-Unidad de Medida
               + "|" + txtConcepto.Text.Trim()                                //8-Descripcion
               + "|" + "0"                                  //9-Valor Unitario
               + "|" + "0"                                        //10-Importe
               + "|" + descuento.Trim()                                      //11-Descuento                                                  
                                                                             //12 Importe con iva si el rfc es XAXX010101000 y XEXX010101000 OPCIONAL
               + "|"                                                  //Fin Del Registro
                );

                escrituraFactura += "\\n04"                                                   //1-Tipo De Registro
               + "|" + consecutivoconcepto.Trim()                            //2-Consecutivo Concepto
               + "|" + claveproductoservicio.Trim()                          //3-Clave Producto o Servicio SAT                                               
               + "|" + numidentificacion.Trim()                             //4-Numero Identificacion TDR
               + "|" + txtCantidad.Text.Trim()                        //5-Cantidad
               + "|" + claveunidad.Trim()                                   //6-Clave Unidad SAT
               + "|"                                                  //7-Unidad de Medida
               + "|" + txtConcepto.Text.Trim()                                //8-Descripcion
               + "|" + "0"                                  //9-Valor Unitario
               + "|" + "0"                                        //10-Importe
               + "|" + descuento.Trim()                                      //11-Descuento                                                  
                                                                             //12 Importe con iva si el rfc es XAXX010101000 y XEXX010101000 OPCIONAL
               + "|";

                //----------------------------------------Seccion CPAG20 -------------------------------------------------------------------

                //CPAG20 (1:1)
                //escritor.WriteLine(
                //"CPAG20"                         //1-Tipo De Registro
                //+ "|" + "2.0"                    //2-Version
                //+ "|"                               //Fin Del Registro
                //);

                //escrituraFactura += "CPAG20"    //1-Tipo De Registro
                //+ "|"  + "2.0"                   //2-Version  
                //+ "|";		   

                //----------------------------------------Seccion CPAG20TOT -------------------------------------------------------------------

                //CPAG20TOT (1:1)
                //escritor.WriteLine(
                //"CPAG20TOT"                         //1-Tipo De Registro
                //+ "|"                               //2-TotalRetencionesIVA
                //+ "|"                               //3-TotalRetencionesISR                                              
                //+ "|"                               //4-TotalRetencionesIEPS
                //+ "|"                               //5-TotalTrasladosBaseIVA16
                //+ "|"                               //6-TotalTrasladosImpuestoIVA16
                //+ "|"                               //7-TotalTrasladosBaseIVA8
                //+ "|"                               //8-TotalTrasladosImpuestoIVA8
                //+ "|"                               //9-TotalTrasladosBaseIVA0
                //+ "|"                               //10-TotalTrasladosImpuestoIVA0
                //+ "|"                               //11-TotalTrasladosBaseIVAExento
                // + "|" + monto                       //12-MontoTotalPagos                                                                                                 
                // + "|"                               //Fin Del Registro
                //);

                //escrituraFactura += "CPAG20TOT"    //1-Tipo De Registro
                //+ "|"                               //2-TotalRetencionesIVA
                //+ "|"                               //3-TotalRetencionesISR                                              
                //+ "|"                               //4-TotalRetencionesIEPS
                //+ "|"                               //5-TotalTrasladosBaseIVA16
                //+ "|"                               //6-TotalTrasladosImpuestoIVA16
                //+ "|"                               //7-TotalTrasladosBaseIVA8
                //+ "|"                               //8-TotalTrasladosImpuestoIVA8
                //+ "|"                               //9-TotalTrasladosBaseIVA0
                //+ "|"                               //10-TotalTrasladosImpuestoIVA0
                //+ "|"                               //11-TotalTrasladosBaseIVAExento
                //+ "|" + monto                       //12-MontoTotalPagos  
                //+ "|";		   
                //----------------------------------------Seccion CPAG20PAGO-------------------------------------------------------------------------------------------------

                //CPAG20PAGO COMPLEMENTO DE PAGO (1:N)
                //escritor.WriteLine(
                //"CPAG20PAGO"                                           //1-Tipo De Registro
                //+ "|" + identificador                                  //2-Identificador                                             
                //+ "|" + fechapago                                      //3-Fechapago
                //+ "|" + txtFormaPago.Text                              //4-Formadepagocpag
                //+ "|" + monedacpag                                     //5-Monedacpag
                //+ "|" + tipodecambiocpag                               //6-TipoDecambiocpag
                //+ "|" + txtTotal.Text                                  //8-Monto
                //+ "|" + numerooperacion                                //9-NumeroOperacion
                //+ "|" + txtRFCbancoEmisor.Text                         //10-RFCEmisorCuentaBeneficiario
                //+ "|" + txtBancoEmisor.Text                            //11-NombreDelBanco                                                                                            
                //+ "|" + txtCuentaPago.Text                             //12-NumeroCuentaOrdenante
                //+ "|" + rfcemisorcuentaben                             //13-RFCEmisorCuentaBeneficiario
                //+ "|" + numcuentaben                                   //14-NumCuentaBeneficiario
                //+ "|" + tipocadenapago                                 //15-TipoCadenaPago                                               
                //+ "|" + certpago                                       //16-CertificadoPago
                //+ "|" + cadenadelpago                                  //17-CadenaDePago
                //+ "|" + sellodelpago                                   //Fin Del Registro
                //+ "|"
                //);

                //escrituraFactura += "CPAG20PAGO"                      //1-Tipo De Registro
                //+ "|" + identificador                                  //2-Identificador                                             
                //+ "|" + fechapago                                      //3-Fechapago
                //+ "|" + txtFormaPago.Text                              //4-Formadepagocpag
                //+ "|" + monedacpag                                     //5-Monedacpag
                //+ "|" + tipodecambiocpag                               //6-TipoDecambiocpag
                //+ "|" + txtTotal.Text                                  //8-Monto
                //+ "|" + numerooperacion                                //9-NumeroOperacion
                //+ "|" + txtRFCbancoEmisor.Text                         //10-RFCEmisorCuentaBeneficiario
                //+ "|" + txtBancoEmisor.Text                            //11-NombreDelBanco                                                                                            
                //+ "|" + txtCuentaPago.Text                             //12-NumeroCuentaOrdenante
                //+ "|" + rfcemisorcuentaben                             //13-RFCEmisorCuentaBeneficiario
                //+ "|" + numcuentaben                                   //14-NumCuentaBeneficiario
                //+ "|" + tipocadenapago                                 //15-TipoCadenaPago                                               
                //+ "|" + certpago                                       //16-CertificadoPago
                //+ "|" + cadenadelpago                                  //17-CadenaDePago
                //+ "|" + sellodelpago                                   //Fin Del Registro
                //+ "|";


                //----------------------------------------Seccion CPAG-------------------------------------------------------------------------------------------------

                //CPAG COMPLEMENTO DE PAGO (1:N)
                escritor.WriteLine(
               "CPAG"                                                 //1-Tipo De Registro
               + "|" + identificador.Trim()                                  //2-Identificador
               + "|" + version.Trim()                                        //3-Version                                             
               + "|" + fechapago.Trim()                                      //4-Fechapago
               + "|" + formadepago.Trim()                              //5-Formadepagocpag
               + "|" + monedascpadgoc.Trim()                                     //6-Monedacpag
               + "|" + tipodecambiocpag.Trim()                               //7-TipoDecambiocpag AQUI LO VOY A TOMAR DE OTRA CONSULTA
               + "|" + txtTotal.Text.Trim()                                 //8-Monto
               + "|" + numerooperacion.Trim()                               //9-NumeroOperacion
               + "|" + txtRFCbancoEmisor.Text.Trim()                         //10-RFCEmisorCuentaBeneficiario
               + "|" + txtBancoEmisor.Text.Trim()                            //11-NombreDelBanco                                                                                            
               + "|" + txtCuentaPago.Text.Trim()                            //12-NumeroCuentaOrdenante
               + "|" + rfcemisorcuentaben.Trim()                           //13-RFCEmisorCuentaBeneficiario
               + "|" + numcuentaben.Trim()                                //14-NumCuentaBeneficiario
               + "|" + tipocadenapago.Trim()                                //15-TipoCadenaPago                                               
               + "|" + certpago.Trim()                                    //16-CertificadoPago
               + "|" + cadenadelpago.Trim()                                //17-CadenaDePago
               + "|" + sellodelpago.Trim()                                  //Fin Del Registro
               + "|"
                );

                escrituraFactura += "CPAG"                                                 //1-Tipo De Registro
               + "|" + identificador.Trim()                                 //2-Identificador
               + "|" + version.Trim()                                   //3-Version                                             
               + "|" + fechapago.Trim()                                     //4-Fechapago
               + "|" + formadepago.Trim()                              //5-Formadepagocpag
               + "|" + monedascpadgoc.Trim()                                    //6-Monedacpag
               + "|" + tipodecambiocpag.Trim()                              //7-TipoDecambiocpag
               + "|" + txtTotal.Text.Trim()                                 //8-Monto
               + "|" + numerooperacion.Trim()                               //9-NumeroOperacion
               + "|" + txtRFCbancoEmisor.Text.Trim()                        //10-RFCEmisorCuentaBeneficiario
               + "|" + txtBancoEmisor.Text.Trim()                          //11-NombreDelBanco                                                                                            
               + "|" + txtCuentaPago.Text.Trim()                           //12-NumeroCuentaOrdenante
               + "|" + rfcemisorcuentaben.Trim()                            //13-RFCEmisorCuentaBeneficiario
               + "|" + numcuentaben.Trim()                                //14-NumCuentaBeneficiario
               + "|" + tipocadenapago.Trim()                                //15-TipoCadenaPago                                               
               + "|" + certpago.Trim()                                  //16-CertificadoPago
               + "|" + cadenadelpago.Trim()                                 //17-CadenaDePago
               + "|" + sellodelpago.Trim()                                 //Fin Del Registro
               + "|";

                //----------------------------------------Seccion CPAGDOC------------------------------------------------------------------------------------------------

                //CPAG COMPLEMENTO DE PAGO (1:N)
                escritor.WriteLine(cpagdoc
                //"CPAGDOC"                                              //1-Tipo De Registro
                //+ "|" + identpag                                       //2-IdentificadorDelPago
                //+ "|" + txtFechaIniOP.Text                             //3-IdentificadorDelDocumentoPagado                                              
                //+ "|" + seriecpag                                      //4-Seriecpag
                //+ "|" + foliocpag                                      //5-Foliocpag
                //+ "|" + monedacpagdoc                                  //6-Monedacpag
                //+ "|" + tipocambiocpag                                 //7-TipoCambiocpagdpc
                //+ "|" + txtMetodoPago.Text                             //8-MetodoDePago
                //+ "|" + numerodeparcialidad                            //9-NumeroDeParcialidad
                //+ "|" + importeSaldoAnterior                           //10-ImporteSaldoAnterior
                //+ "|" + importepago                                    //11-ImportePagado                                                  
                //+ "|" + importesaldoinsoluto                           //12 ImporteSaldoInsoluto
                //+ "|"                                                  //Fin Del Registro
                );


                escrituraFactura += cpagdoc;
                //escrituraFactura = escrituraFactura.Replace("||02|", "||\\n02|");
                //escrituraFactura = escrituraFactura.Replace("||04|", "||\\n04|");
                escrituraFactura = escrituraFactura.Replace("| \r\n", "|");
                escrituraFactura = escrituraFactura.Replace("|CPAG", "|\\nCPAG");

            }
        }

        //Valida que el campo no contenga caracteres como el char 13,10,9
        public string validaCampo(string campo)
        {
            char salto = (char)10;
            char tab = (char)9;
            char carrier = (char)13;
            char espacio = (char)32;
            campo = campo.Replace(salto, espacio);
            campo = campo.Replace(tab, espacio);
            campo = campo.Replace(carrier, espacio);
            return campo;
        }

        protected void btnEdit_Click(object sender, EventArgs e)
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }
            bool readOnly = false;
            string stilo = "editTextBox";
            string textoBoton = "Editar";
            bool visible = false;
            if (btnEdit.Text.Equals("Editar"))
            {
                readOnly = false;
                visible = true;
                stilo = "editTextBox";
                textoBoton = "Guardar";
            }
            else
            {
                string fecha1 = "", fecha2 = "";

                bool vacio = false;

                System.Globalization.DateTimeFormatInfo dateInfo = new System.Globalization.DateTimeFormatInfo();
                dateInfo.ShortDatePattern = "dd/MM/yyyy";

                if (!txtFechaDesde.Text.Equals("")) { fecha1 = txtFechaDesde.Text; imgFDesde.Visible = false; }
                else { error = true; imgFDesde.Visible = true; imgFDesde.ToolTip = "La fecha no puede estar vacía"; vacio = true; }

                if (!txtFechaHasta.Text.Equals("")) { fecha2 = txtFechaHasta.Text; imgFHasta.Visible = false; }
                else { error = true; imgFHasta.Visible = true; imgFHasta.ToolTip = "La feche no puede estar vacía"; vacio = true; }

                if (fecha1.CompareTo(fecha2) == 1 && vacio == false)
                {
                    error = true;
                    imgFDesde.Visible = true;
                    imgFDesde.ToolTip = "La Fecha inicial es mayor que la final";
                }
                if (!error)
                {
                    readOnly = true;
                    visible = false;
                    stilo = "readOnlyTextBox";
                    textoBoton = "Editar";
                }
                else
                {
                    readOnly = false;
                    visible = true;
                    stilo = "editTextBox";
                    textoBoton = "Guardar";
                }
            }
            //Se habilita el campo
            txtFechaDesde.ReadOnly = readOnly;
            txtFechaHasta.ReadOnly = readOnly;
            txtConcepto.ReadOnly = readOnly;
            txtTipoCobro.ReadOnly = readOnly;
            txtFormaPago.ReadOnly = readOnly;


            //Se elimina el estilo
            txtFechaDesde.CssClass = stilo;
            txtFechaHasta.CssClass = stilo;
            txtFechaDesde.CssClass = stilo;
            txtConcepto.CssClass = stilo;
            txtTipoCobro.CssClass = stilo;

            btnEdit.Text = textoBoton;

        }
        protected void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }
            bool campoIncorrecto = false;
            if (txtCliente.Text.Equals(""))
            {
                imgCliente.Visible = true;
                imgCliente.ToolTip = "El cliente no está capturado";
                campoIncorrecto = true;
            }

            if (txtCalle.Text.Equals(""))
            {
                imgCalle.Visible = true;
                imgCalle.ToolTip = "La calle no está capturada";
                campoIncorrecto = true;
            }

            if (txtNoExt.Equals(""))
            {
                imgNoExt.Visible = true;
                imgNoExt.ToolTip = "El No. Ext. no está capturado";
                campoIncorrecto = true;
            }

            // if (noInt.Equals (""))
            //{
            //    imgDir.Visible = true;
            //    imgDir.ToolTip = "El No. Int. no esta capturado";
            //    campoIncorrecto = true;
            //}


            if (txtColonia.Text.Equals(""))
            {
                imgColonia.Visible = true;
                imgColonia.ToolTip = "La colonia no está capturada";
                campoIncorrecto = true;
            }

            if (txtMunicipio.Text.Equals(""))
            {
                imgMunicipio.Visible = true;
                imgMunicipio.ToolTip = "El municipio no está capturado";
                campoIncorrecto = true;
            }

            if (txtEstado.Text.Equals(""))
            {
                imgEstado.Visible = true;
                imgEstado.ToolTip = "El Estado. no está capturado";
                campoIncorrecto = true;
            }

            if (txtPaís.Text.Equals(""))
            {
                imgPais.Visible = true;
                imgPais.ToolTip = "El País no está capturado";
                campoIncorrecto = true;
            }

            if (txtCP.Text.Equals(""))
            {
                imgCP.Visible = true;
                imgCP.ToolTip = "El CP. no está capturado";
                campoIncorrecto = true;
            }


            if (!campoIncorrecto)
            {
                //Se elimina el estilo
                txtFechaDesde.CssClass = "readOnlyTextBox";
                txtFechaHasta.CssClass = "readOnlyTextBox";
                txtFechaDesde.CssClass = "readOnlyTextBox";
                txtConcepto.CssClass = "readOnlyTextBox";
                txtTipoCobro.CssClass = "readOnlyTextBox";

                generaTXT();


                //Traer variable generada por txt
                //Traer vairable idSucursal y idTipoFact
                //Traer lbltex´+.txt
                //Etiqueta archivo fuente
                //Crear el JSON
                //Insertar código petición 
                var request_ = (HttpWebRequest)WebRequest.Create("https://canal1.xsa.com.mx:9050/bf2e1036-ba47-49a0-8cd9-e04b36d5afd4/tiposCfds");
                var response_ = (HttpWebResponse)request_.GetResponse();
                var responseString_ = new StreamReader(response_.GetResponseStream()).ReadToEnd();

                string[] separadas_ = responseString_.Split('}');



                foreach (string dato in separadas_)
                {
                    if (dato.Contains("TDRC"))
                    {
                        string[] separadasSucursal_ = dato.Split(',');
                        foreach (string datoSuc in separadasSucursal_)
                        {
                            if (datoSuc.Contains("idSucursal"))
                            {
                                idSucursal = datoSuc.Replace(dato.Substring(0, 8), "").Replace("\"", "").Split(':')[1];
                            }
                            if (datoSuc.Contains("id") && !datoSuc.Contains("idSucursal"))
                            {
                                idTipoFactura = datoSuc.Replace(dato.Substring(0, 8), "").Replace("\"", "").Split(':')[1];

                            }
                        }
                    }
                }

                jsonFactura = "{\r\n\r\n  \"idTipoCfd\":" + "\"" + idTipoFactura + "\"";
                jsonFactura += ",\r\n\r\n  \"nombre\":" + "\"" + lblFact.Text + ".txt" + "\"";
                jsonFactura += ",\r\n\r\n  \"idSucursal\":" + "\"" + idSucursal + "\"";
                jsonFactura += ", \r\n\r\n  \"archivoFuente\":" + "\"" + escrituraFactura + "\"" + "\r\n\r\n}";

                string folioFactura = "", serieFactura = "", uuidFactura = "", pdf_xml_descargaFactura = "", pdf_descargaFactura = "", xlm_descargaFactura = "", cancelFactura = "", error = "";
                string salida = "";

                try
                {
                    var client = new RestClient("https://canal1.xsa.com.mx:9050/bf2e1036-ba47-49a0-8cd9-e04b36d5afd4/cfdis");
                    var request = new RestRequest(Method.PUT);

                    request.AddHeader("cache-control", "no-cache");

                    request.AddHeader("content-length", "834");
                    request.AddHeader("accept-encoding", "gzip, deflate");
                    request.AddHeader("Host", "canal1.xsa.com.mx:9050");
                    request.AddHeader("Postman-Token", "b6b7d8eb-29f2-420f-8d70-7775701ec765,a4b60b83-429b-4188-98d4-7983acc6742e");
                    request.AddHeader("Cache-Control", "no-cache");
                    request.AddHeader("Accept", "*/*");
                    request.AddHeader("User-Agent", "PostmanRuntime/7.13.0");

                    request.AddParameter("application/json", jsonFactura, ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);

                    string[] separadaFactura = response.Content.ToString().Split(',');
                    foreach (string factura in separadaFactura)
                    {
                        if (factura.Contains("errors"))
                        {
                            error += factura.Replace(factura.Substring(0, 6), "").Replace("\"", "").Split('[')[1] + "\n";
                            error = error.Replace("\\n", "").Replace("]}", "");
                            salida = "FALLA AL SUBIR";
                        }
                        else
                        {
                            if (factura.Contains("folio"))
                            {
                                folioFactura = factura.Replace(factura.Substring(0, 5), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("serie"))
                            {
                                serieFactura = factura.Replace(factura.Substring(0, 5), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("uuid"))
                            {
                                uuidFactura = factura.Replace(factura.Substring(0, 4), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("pdfAndXmlDownload"))
                            {
                                pdf_xml_descargaFactura = factura.Replace(factura.Substring(0, 17), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("pdfDownload"))
                            {
                                pdf_descargaFactura = factura.Replace(factura.Substring(0, 11), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("xmlDownload") && !factura.Contains("pdfAndXmlDownload"))
                            {
                                xlm_descargaFactura = factura.Replace(factura.Substring(0, 11), "").Replace("\"", "").Split(':')[1];
                            }

                            if (factura.Contains("cancellCfdi"))
                            {
                                cancelFactura = factura.Replace(factura.Substring(0, 11), "").Replace("\"", "").Split(':')[1];
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    string msg = "¡Error, ponte en contacto con TI";
                    ScriptManager.RegisterStartupScript(this, GetType(), "swal", "swal('" + msg + "', 'Error con los folios relacionados ', 'error');setTimeout(function(){window.location.href ='Listado.aspx'}, 10000)", true);
                }


                string path = System.Web.Configuration.WebConfigurationManager.AppSettings["dir"] + lblFact.Text + ".txt";

                //UploadFile file = new UploadFile();
                string ftp = System.Web.Configuration.WebConfigurationManager.AppSettings["ftp"];
                if (ftp.Equals("Si"))
                {
                    //File.prubeftp(lblFact.Text + ".txt", path, serie);

                }
                if (salida != "FALLA AL SUBIR")
                {
                    string activa = System.Web.Configuration.WebConfigurationManager.AppSettings["activa"];
                    if (activa.Equals("Si"))
                    {

                        facLabControler.insertaFactura(txtFolio.Text, txtFechaFactura.Text);
                    }

                    string msg = "¡Se ha generado correctamente el CFDi!";
                    ScriptManager.RegisterStartupScript(this, GetType(), "swal", "swal('" + msg + "', 'Carga exitosa', 'success');setTimeout(function(){window.location.href ='Listado.aspx'}, 10000)", true);
                }
                else
                {
                    string msg = "¡Error al conectar al servicio XSA!";
                    ScriptManager.RegisterStartupScript(this, GetType(), "swal", "swal('" + msg + "', 'Error', 'error');setTimeout(function(){window.location.href ='Listado.aspx'}, 10000)", true);

                }
            }
        }
        protected void Button2_Click(object sender, EventArgs e)
        {



            string msg = "¡Se ha generado correctamente el CFDi!";
            ScriptManager.RegisterStartupScript(this, GetType(), "swal", "swal('" + msg + "', 'Carga exitosa', 'success');", true);





        }

        protected void btnGenerarTxt_Click(object sender, EventArgs e)
        {
            try
            { //try TLS 1.3
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)12288
                                                     | (SecurityProtocolType)3072
                                                     | (SecurityProtocolType)768
                                                     | SecurityProtocolType.Tls;
            }
            catch (NotSupportedException)
            {
                try
                { //try TLS 1.2
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072
                                                         | (SecurityProtocolType)768
                                                         | SecurityProtocolType.Tls;
                }
                catch (NotSupportedException)
                {
                    try
                    { //try TLS 1.1
                        ServicePointManager.SecurityProtocol = (SecurityProtocolType)768
                                                             | SecurityProtocolType.Tls;
                    }
                    catch (NotSupportedException)
                    { //TLS 1.0
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                    }
                }
            }
            bool campoIncorrecto = false;
            btnEdit.Visible = false;
            btnGuardar.Visible = false;
            if (txtCliente.Text.Equals(""))
            {
                imgCliente.Visible = true;
                imgCliente.ToolTip = "El cliente no está capturado";
                campoIncorrecto = true;
            }

            if (txtCalle.Text.Equals(""))
            {
                imgCalle.Visible = true;
                imgCalle.ToolTip = "La calle no está capturada";
                campoIncorrecto = true;
            }

            if (txtNoExt.Equals(""))
            {
                imgNoExt.Visible = true;
                imgNoExt.ToolTip = "El No. Ext. no está capturado";
                campoIncorrecto = true;
            }

            // if (noInt.Equals (""))
            //{
            //    imgDir.Visible = true;
            //    imgDir.ToolTip = "El No. Int. no esta capturado";
            //    campoIncorrecto = true;
            //}


            if (txtColonia.Text.Equals(""))
            {
                imgColonia.Visible = true;
                imgColonia.ToolTip = "La colonia no está capturada";
                campoIncorrecto = true;
            }

            if (txtMunicipio.Text.Equals(""))
            {
                imgMunicipio.Visible = true;
                imgMunicipio.ToolTip = "El municipio no está capturado";
                campoIncorrecto = true;
            }

            if (txtEstado.Text.Equals(""))
            {
                imgEstado.Visible = true;
                imgEstado.ToolTip = "El Estado. no está capturado";
                campoIncorrecto = true;
            }

            if (txtPaís.Text.Equals(""))
            {
                imgPais.Visible = true;
                imgPais.ToolTip = "El País no está capturado";
                campoIncorrecto = true;
            }

            if (txtCP.Text.Equals(""))
            {
                imgCP.Visible = true;
                imgCP.ToolTip = "El CP. no está capturado";
                campoIncorrecto = true;
            }


            if (!campoIncorrecto)
            {
                //Se elimina el estilo
                txtFechaDesde.CssClass = "readOnlyTextBox";
                txtFechaHasta.CssClass = "readOnlyTextBox";
                txtFechaDesde.CssClass = "readOnlyTextBox";
                txtConcepto.CssClass = "readOnlyTextBox";
                txtTipoCobro.CssClass = "readOnlyTextBox";

                generadorTXT();
                string msg = "¡Se genero correctamente el TXT!";
                ScriptManager.RegisterStartupScript(this, GetType(), "swal", "swal('" + msg + "', 'Success', 'success');setTimeout(function(){window.location.href ='DownloadTxt.aspx'}, 10000)", true);


            }
        }
    }
}