using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;

namespace ComplementoPagoVOld.Models
{
    public class ModelosFact

    {
        //SELECT * FROM sae_archivos WHERE folio like '%40982%'
        public string uuid { get; set; }
        public string motivo { get; set; }
        public string status { get; set; }
        public string xmlDownload { get; set; }
        public string folio { get; set; }
        public string fecha { get; set; }
        public string serie { get; set; }
        public string rfc { get; set; }
        public string ord_hdrnumber { get; set; }
        public DateTime FechaActual
        {
            get { return DateTime.Now; }
            set { this.FechaActual = value; }
        }
        public string tcfix { get; set; }
        private const string facturasListadop = "select CONVERT(INT, folio) as Folio, FechaHoraEmision as Fecha, Nombre as Cliente from vista_fe_copago_Enviados order by CONVERT(INT, folio) ASC";
        private const string facturas = "select CONVERT(INT, folio) as Folio, FechaHoraEmision as Fecha, Nombre as Cliente,idreceptor from vista_fe_copago order by CONVERT(INT, folio) ASC";
        private const string facturasEnviadas = "select CONVERT(INT, folio) as Folio, FechaHoraEmision as Fecha, Nombre as Cliente from vista_fe_copago_Enviados order by CONVERT(INT, folio) ASC";
        private const string datosFactura = "select * from vista_fe_copago where folio = @factura and medotodepago = 'PPD' union select * from vista_fe_copago_Enviados where folio = @factura and medotodepago = 'PPD' ";
        private const string datosCPAGDOC = "select * from vista_fe_copago_cpagdoc where identificadordelPago = @identificador";
        private const string datosMaster = "select invoice as folio from [172.24.16.112].[TMWSuite].[dbo].VISTA_fe_generadas where nmaster = @identificador";
        private const string insert = "insert into dbo.sae_archivos (serie,folio,fecha,factura) values('C',@factura,@fecha,@factura)";
        private const string P_fact = "@factura";
        private const string P_fecha = "@fecha";
        private const string P_Ident = "@identificador";


        public ModelosFact()
        {

        }

        public DataTable getFacturas()
        {
            DataTable dataTablee = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select top 5 Folio, FechaPago as Fecha, Nombre as Cliente,idreceptor from vista_fe_copago order by Folio ASC", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 200;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTablee);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTablee;
        }
        public DataTable getTipoCambio(string fecha)
        {


            DataTable dataTable3 = new DataTable();
            //NOS CONECTAMOS CON LA BASE DE DATOS
            string cadena = @"Data source=172.24.16.113; Initial Catalog=DYNAMICS; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection cn = new SqlConnection(cadena))
            {
                try
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("TipoCambioJC", cn))
                    {
                        //Le indico que es del itpo procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 1000;
                        //Esta linea define un parametro
                        cmd.Parameters.AddWithValue("@fecha", fecha);
                        //cmd.Parameters.AddWithValue("@foliocpag", foliocpag);
                        //Ejecutamos el procedimiento
                        cmd.ExecuteNonQuery();
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
                        {
                            try
                            {

                                sqlDataAdapter.Fill(dataTable3);
                                cn.Close();
                            }
                            catch (SqlException ex)
                            {
                                cn.Close();
                                string message = ex.Message;

                            }

                        }

                    }
                }
                catch (SqlException ex)
                {

                    cn.Close();
                    string message = ex.Message;

                }
            }

            return dataTable3;
        }
        public DataTable getDatosFacturasV(string fact, string IdRecep)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select * from vista_fe_copago where folio = @factura and medotodepago = 'PPD' and IdReceptor = @IdRecep union select * from vista_fe_copago_Enviados where folio = @factura and medotodepago = 'PPD' and IdReceptor = @IdRecep", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@factura", (object)fact);
                    selectCommand.Parameters.AddWithValue("@IdRecep", (object)IdRecep);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable Elist2(string identificador)
        {


            DataTable dataTable3 = new DataTable();
            //NOS CONECTAMOS CON LA BASE DE DATOS
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection cn = new SqlConnection(cadena))
            {
                try
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("sp_nomostrarsiexiste_JC", cn))
                    {
                        //Le indico que es del itpo procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 1000;
                        //Esta linea define un parametro
                        cmd.Parameters.AddWithValue("@folio", identificador);
                        //cmd.Parameters.AddWithValue("@foliocpag", foliocpag);
                        //Ejecutamos el procedimiento
                        cmd.ExecuteNonQuery();
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
                        {
                            try
                            {

                                sqlDataAdapter.Fill(dataTable3);
                                cn.Close();
                            }
                            catch (SqlException ex)
                            {
                                cn.Close();
                                string message = ex.Message;

                            }

                        }

                    }
                }
                catch (SqlException ex)
                {

                    cn.Close();
                    string message = ex.Message;

                }
            }

            return dataTable3;
        }
        public void IFpagoDeletePapelera(string idnum)
        {
            string cadena2 = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena2))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_deleteBilltoPapelera_CPOLD_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 100000;
                    selectCommand.Parameters.AddWithValue("@idnum", idnum);


                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }
        public DataTable billtoFp()
        {
            DataTable dataTablee = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select id_num, billto from cpoldebillto order by id_num ASC", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 200866;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTablee);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTablee;
        }
        public void IFpago(string billto)
        {
            string cadena2 = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena2))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_insertaBilltoFpago_CPOLD_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 100000;
                    selectCommand.Parameters.AddWithValue("@billto", billto);


                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }
        public DataTable billtoPapelera()
        {
            DataTable dataTablee = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select id_num, folio from insertarComplementos order by id_num DESC", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 200866;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTablee);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTablee;
        }
        public void Elist(string identificador)
        {
            string cadena2 = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena2))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_insertaComplemento_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 100000;
                    selectCommand.Parameters.AddWithValue("@folio", identificador);


                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }


        public DataTable tipoCambio()
        {
            DataTable dataTablee = new DataTable();
            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select TOP 1 cex_rate from currency_exchange order by cex_date desc", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 200;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTablee);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTablee;
        }
        public DataTable getFacturasClientes()
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select distinct  idreceptor from vista_fe_copago", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;

                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable getFacturasPorProcesar(string billto)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select folio,sfolio,idreceptor from vista_fe_copago where idreceptor = @idreceptor", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@idreceptor", (object)billto);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }

        public DataTable getFacturasEnviadas()
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select CONVERT(INT, folio) as Folio, FechaHoraEmision as Fecha, Nombre as Cliente from vista_fe_copago_Enviados order by CONVERT(INT, folio) ASC", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable getFacturasListadop()
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select CONVERT(INT, folio) as Folio, FechaHoraEmision as Fecha, Nombre as Cliente from vista_fe_copago_Enviados order by CONVERT(INT, folio) ASC", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable getDatosCPAGDOCTRL(string identificador, string foliocpag)
        {


            DataTable dataTable3 = new DataTable();
            //NOS CONECTAMOS CON LA BASE DE DATOS
            string cadena = @"Data source=172.24.16.113; Initial Catalog=DYNAMICS; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection cn = new SqlConnection(cadena))
            {
                try
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("usp_ccpp", cn))
                    {
                        //Le indico que es del itpo procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 1000;
                        //Esta linea define un parametro
                        cmd.Parameters.AddWithValue("@identificador", identificador);
                        cmd.Parameters.AddWithValue("@foliocpag", foliocpag);
                        //Ejecutamos el procedimiento
                        cmd.ExecuteNonQuery();
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
                        {
                            try
                            {

                                sqlDataAdapter.Fill(dataTable3);
                                cn.Close();
                            }
                            catch (SqlException ex)
                            {
                                cn.Close();
                                string message = ex.Message;

                            }

                        }

                    }
                }
                catch (SqlException ex)
                {

                    cn.Close();
                    string message = ex.Message;

                }
            }

            return dataTable3;

        }

        public DataTable getDatosSegmentos(string orden)
        {


            DataTable dataTable3 = new DataTable();
            //NOS CONECTAMOS CON LA BASE DE DATOS
            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection cn = new SqlConnection(cadena))
            {
                try
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("usp_obtener_segmento_JC", cn))
                    {
                        //Le indico que es del itpo procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 1000;
                        //Esta linea define un parametro
                        cmd.Parameters.AddWithValue("@orden", orden);

                        //Ejecutamos el procedimiento
                        cmd.ExecuteNonQuery();
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd))
                        {
                            try
                            {

                                sqlDataAdapter.Fill(dataTable3);
                                cn.Close();
                            }
                            catch (SqlException ex)
                            {
                                cn.Close();
                                string message = ex.Message;

                            }

                        }

                    }
                }
                catch (SqlException ex)
                {

                    cn.Close();
                    string message = ex.Message;

                }
            }

            return dataTable3;

        }
        public DataTable getDatosFacturas(string fact)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select * from vista_fe_copago where folio = @factura and medotodepago = 'PPD' union select * from vista_fe_copago_Enviados where folio = @factura and medotodepago = 'PPD' ", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@factura", (object)fact);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public void IFpagoDelete(string idnum)
        {
            string cadena2 = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena2))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_deleteBilltoFpago_CPOLD_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 100000;
                    selectCommand.Parameters.AddWithValue("@idnum", idnum);


                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }
        public DataTable getBillto(string billto)
        {
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            DataTable dataTable = new DataTable();
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                connection.Open();
                using (SqlCommand selectCommand = new SqlCommand("sp_getBillto_CPOLD_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 100000;
                    selectCommand.Parameters.AddWithValue("@billto", billto);
                    selectCommand.ExecuteNonQuery();
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            //selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            connection.Close();
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable getDatosFacturasF(string fact, string IdRecep)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select * from vista_fe_copago where folio = @factura and medotodepago = 'PPD' and Foliocpag !='' and IdReceptor = @IdRecep union select * from vista_fe_copago_Enviados where folio = @factura and medotodepago = 'PPD' and Foliocpag !='' and IdReceptor = @IdRecep", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000434;
                    selectCommand.Parameters.AddWithValue("@factura", (object)fact);
                    selectCommand.Parameters.AddWithValue("@IdRecep", (object)IdRecep);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }

        public DataTable getDatosCPAGDOC(string identificador, string IdRecep)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select * from vista_fe_copago_cpagdoc_JC where identificadordelPago = @identificador and Billto = @IdRecep", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@identificador", (object)identificador);
                    selectCommand.Parameters.AddWithValue("@IdRecep", (object)IdRecep);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }

        public DataTable getDatosMaster(string identificador)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select invoice as folio from [172.24.16.112].[TMWSuite].[dbo].VISTA_Fe_generadas where nmaster = @identificador", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@identificador", (object)identificador);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }
        public DataTable getDatosInvoice(string identificador)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("select ord_hdrnumber from invoiceheader where ivh_hdrnumber = @identificador", connection))
                {
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@identificador", (object)identificador);
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            sqlDataAdapter.Fill(dataTable);
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
            return dataTable;
        }

        public void insertaFactura(string fact, string fecha)
        {
            DataTable dataTable = new DataTable();
            string cadena = @"Data source=172.24.16.113; Initial Catalog=TDR; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            using (SqlConnection connection = new SqlConnection(cadena))
            {
                using (SqlCommand selectCommand = new SqlCommand("insert into dbo.sae_archivos (serie,folio,fecha,factura) values('C',@factura,@fecha,@factura)", connection))
                {
                    DateTimeFormatInfo dateTimeFormatInfo = new DateTimeFormatInfo();
                    DateTime dateTime = Convert.ToDateTime(fecha);
                    string[] strArray = new string[5]
                    {
            dateTime.Year.ToString(),
            "-",
            null,
            null,
            null
                    };
                    int num = dateTime.Month;
                    strArray[2] = num.ToString();
                    strArray[3] = "-";
                    num = dateTime.Day;
                    strArray[4] = num.ToString();
                    string str = string.Concat(strArray);
                    selectCommand.CommandType = CommandType.Text;
                    selectCommand.CommandTimeout = 5000;
                    selectCommand.Parameters.AddWithValue("@factura", (object)fact);
                    selectCommand.Parameters.AddWithValue("@fecha", (object)str);
                    using (new SqlDataAdapter(selectCommand))
                    {
                        try
                        {
                            selectCommand.Connection.Open();
                            selectCommand.ExecuteNonQuery();
                        }
                        catch (SqlException ex)
                        {
                            string message = ex.Message;
                        }
                    }
                }
            }
        }

    }
}