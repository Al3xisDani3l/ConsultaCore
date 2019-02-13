using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.Common;
using System.Reflection;
using System.Windows;
using System.IO;
using ADOX;
using ADODB;
using System.Data;


namespace ConsultaCore
{
    
    public enum Estados
    {
       
       
       
        Operacion_Exitosa,
        Operacion_Fallida,
        Configurado,
       
        SinCadenaDeConexion,
    

       

        
    
    }

   

    public class CoreDataBaseAccess:IDisposable
    {

        
        
        #region Variables privadas
      

        internal OleDbConnection Conexion;

        private string cadenaDeConexionString;

        private bool isConected = false;

        private Connection ConexionInterna;

        private Estados estado = Estados.SinCadenaDeConexion;

        private string pathDatabase;

        private string provider = "Microsoft.ACE.OLEDB.12.0";

        Catalog catalog = new Catalog();
        #endregion

        #region Propiedades
        /// <summary>
        /// Obtiene el estado actual en el que se encuentra el este objeto.
        /// </summary>
        public Estados Estado { get { return estado; } }
        /// <summary>
        /// Obtiene o establece la cadena de conexion, creando una nueva OlDbConnection al asignar una nueva connectionString.
        /// </summary>
        public string CadenaDeConexion
        {
            get
            {
                if (!string.IsNullOrEmpty(cadenaDeConexionString))
                {
                    return cadenaDeConexionString;
                }
                else
                {
                    return string.Format("Provider={0};Data Source=\"{1}\"", Provider, pathDatabase);
                }
                
            }
            set
            {
               
                    try
                    {
                        Conexion = new OleDbConnection(value);
                        Conexion.Open();
                    cadenaDeConexionString = value;
                    int startProvider = value.IndexOf('=');
                    int endProvider = value.IndexOf(';');
                    int startPath = value.IndexOf('"');
                    int cadenaLength = value.Length;
                    provider = value.Substring(startProvider+1, (endProvider - startProvider)-1);
                    pathDatabase = value.Substring(startPath+1, (cadenaLength - startPath) -2);


                        estado = Estados.Configurado;
                        
                    }
                    catch (Exception error)
                    {
                        Conexion = null;
                        estado = Estados.SinCadenaDeConexion;
                        throw error;
                    }
                    finally
                    {
                        Conexion.Close();
                    }

                

            }
        }

        public string Provider
        {
            get
            {
                return provider;
            }
            set
            {
               
                if (!string.IsNullOrEmpty(value))
                { 
                    if (!string.IsNullOrEmpty(cadenaDeConexionString) && string.IsNullOrEmpty(provider))
                    {
                        CadenaDeConexion = cadenaDeConexionString.Replace(provider, value);
                       
                    }
                    else if (!string.IsNullOrEmpty(pathDatabase))
                    {
                        CadenaDeConexion = string.Format("Provider={0};Data Source =\"{1}\"", value, pathDatabase);
                      
                    }
                    else
                    {
                        provider = value;
                    }
                   
                }
            }
        }
       
        public string PathDataBase
        {
            get
            {
                return pathDatabase;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    if (!string.IsNullOrEmpty(cadenaDeConexionString) && string.IsNullOrEmpty(pathDatabase))
                    {
                        CadenaDeConexion = cadenaDeConexionString.Replace(pathDatabase, value);
                        
                    }
                    else if (!string.IsNullOrEmpty(provider))
                    {
                        CadenaDeConexion = string.Format("Provider={0};Data Source =\"{1}\"", provider, value);
                       
                    }
                    else
                    {
                        pathDatabase = value;
                    }
                }
                
            }
        }

        public bool IsConnected { get => isConected; set => isConected = value; }

        #endregion

        #region Constructores
        /// <summary>
        /// Representa un objeto capaz de realizar consultas concretas bases de datos de Access
        /// Atraves del modelo de propiedades, y todo aquel objeto que implemente IDateable.
        /// </summary>
        /// <param name="CadenaDeConexion">Cadena que contiene la direccion y provedor OleDb.</param>
        public CoreDataBaseAccess(string CadenaDeConexion )
        {
            this.cadenaDeConexionString = CadenaDeConexion;
            try
            {
                Conexion = new OleDbConnection(this.cadenaDeConexionString);
                Conexion.Open();
                isConected = true;
                ConexionInterna = new Connection();
                ConexionInterna.ConnectionString = this.cadenaDeConexionString;
               
              
            }
            catch (Exception Error)
            {
                this.estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                throw Error;
               
            }
            finally
            {
                Conexion.Close();
                isConected = false;
            }

        }
        /// <summary>
        /// Representa un objeto capaz de realizar consultas concretas bases de datos de Access
        /// Atraves del modelo de propiedades, y todo aquel objeto que implemente IDateable.
        /// </summary>
        /// <param name="Conexion"> Representa una conexion previamente configurada para un archivo de bases de datos.</param>
        public CoreDataBaseAccess(OleDbConnection Conexion)
        {
            this.Conexion = Conexion;
            cadenaDeConexionString = Conexion.ConnectionString;
           
        }
        /// <summary>
        /// Representa un objeto capaz de realizar consultas concretas bases de datos de Access
        /// Atraves del modelo de propiedades, y todo aquel objeto que implemente IDateable.
        /// Este Constructo se inicia con parametros nullos, para configurar solo debe agregar una cadena de conexion a la propiedad CadenaDeConexion.
        /// </summary>
        public CoreDataBaseAccess()
        {
            estado = Estados.SinCadenaDeConexion;
        }
        #endregion

        #region Funciones

        public void Close()
        {
            if (Conexion != null)
            {
                Conexion.Close();
                isConected = false;
            }
        }


        public bool Conectar(bool isAutomatico = true)
        {
            if (!isAutomatico)
            {
                if (!string.IsNullOrWhiteSpace(this.Provider) || !string.IsNullOrWhiteSpace(this.pathDatabase))
                {
                    try
                    {
                        Conexion = new OleDbConnection(string.Format("Provider={0};Data Source=\"{1}\"", Provider, pathDatabase));
                        Conexion.Open();
                        isConected = true;
                       
                        this.cadenaDeConexionString = string.Format("Provider={0};Data Source=\"{1}\"", Provider, pathDatabase);
                        return true;
                    }
                    catch (Exception Error)
                    {
                        this.estado = Estados.Operacion_Fallida;
                        new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                        return false;
                        
                    }
                    finally
                    {
                        Conexion.Close();
                        isConected = false;
                    }
                }
                else
                {
                    return false;
                    throw new Exception("No se ha especificado una cadena de conexion o una ruta valida, public void Conectar();");
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(this.cadenaDeConexionString))
                {
                    try
                    {
                        Conexion = new OleDbConnection(CadenaDeConexion);
                        Conexion.Open();
                        this.estado = Estados.Configurado;
                        return true;
                    }
                    catch (Exception t)
                    {
                        this.estado = Estados.Operacion_Fallida;
                        new LogInternal(t.ToString());
                        return false;
                        
                    }
                    finally
                    {
                        Conexion.Close();
                    }
                }
                else
                {
                    return false;
                    throw new Exception("No se ha especificado una cadena de conexion o una ruta valida, public void Conectar();");
                }

            }
           
           
           
           
         
        }

        public async void InsertarAsync(IDateable Objeto)
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {
                try
                {
                    using (OleDbCommand Comando = new OleDbCommand(InsertToCommmand(Objeto), Conexion))
                    {
                        Conexion.Open();
                        isConected = true;
                        await Comando.ExecuteNonQueryAsync();
                        this.estado = Estados.Operacion_Exitosa;
                    }
                }
                catch (OleDbException Error)
                {
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    this.estado = Estados.Operacion_Fallida;
                    


                }

                finally
                {
                    Conexion.Close();
                    isConected = false;
                    
                }
            }
            

           
                

        }

        public async void InsertarAsync(IDateable[] Objetos)
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {

                try
                {
                    for (int i = 0; i < Objetos.Length; i++)
                    {
                        using (OleDbCommand Comando = new OleDbCommand(InsertToCommmand(Objetos[i]), Conexion))
                        {
                            Conexion.Open();
                            isConected = true;
                            await Comando.ExecuteNonQueryAsync();
                            
                        }
                        Conexion.Close();
                        isConected = false;
                    }
                    estado = Estados.Operacion_Exitosa;
                }
                catch (OleDbException Error)
                {

                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    estado = Estados.Operacion_Fallida;

                }

                finally
                {
                    Conexion.Close();
                    isConected = false;
                }

            }
        }

        public async void DeleteAsync(IDateable Objeto)
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {


                string cmd = string.Format("DELETE FROM {0} WHERE Id = {1}", Objeto.Tabla, Objeto.Id);
                try
                {
                    using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                    {
                        Conexion.Open();
                        isConected = true;
                        await comando.ExecuteNonQueryAsync();
                        this.estado = Estados.Operacion_Exitosa;
                    }
                }
                catch (OleDbException Error)
                {
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    this.estado = Estados.Operacion_Fallida;

                }
                finally
                {
                    Conexion.Close();
                    isConected = false;
                }
            }
        }

        public async void DeleteAsync(IDateable[] Objetos)
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {
                try
                {
                    for (int i = 0; i < Objetos.Length; i++)
                    {
                        string cmd = string.Format("DELETE FROM {0} WHERE Id = {1}", Objetos[i].Tabla, Objetos[i].Id);

                        using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                        {
                            Conexion.Open();
                            isConected = true;
                            await comando.ExecuteNonQueryAsync();
                            this.estado = Estados.Operacion_Exitosa;
                        }
                    }

                }
                catch (OleDbException Error)
                {
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    this.estado = Estados.Operacion_Fallida;

                }
                finally
                {
                    Conexion.Close();
                    isConected = false;
                
                }
            }
        }

        public async Task<List<T>> ReadAsync<T>(string filtro) where T : IDateable, new()
        {


            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {



                List<T> retorno = new List<T>();
                T buf = new T();

                string cmd = string.Format("SELECT * FROM {0} WHERE {1}", buf.Tabla, filtro);

                try
                {
                    using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                    {
                        Conexion.Open();
                        isConected = true;
                        DbDataReader lectura = await comando.ExecuteReaderAsync().ConfigureAwait(false);

                        while (lectura.Read())
                        {
                     
                            T buffer = new T();
                            for (int i = 0; i < buffer.Propiedades.Count; i++)
                            {
                                if ("Tabla" != buffer.Propiedades[i].Name && "Propiedades" != buffer.Propiedades[i].Name)
                                {
                                    buffer.Propiedades[i].SetValue(buffer, Convert.ChangeType(lectura[buffer.Propiedades[i].Name], buffer.Propiedades[i].PropertyType));
                                }

                            }

                            retorno.Add(buffer);

                        }
                        this.estado = Estados.Operacion_Exitosa;
                        return retorno;

                    }
                }
                catch (Exception Error)
                {
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    this.estado = Estados.Operacion_Fallida;
                    return new List<T>();

                }
                finally
                {
                    Conexion.Close();
                    isConected = false;

                }
            }
        }

        public async Task<List<T>> ReadAsync<T>() where T : IDateable, new()
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {



                List<T> retorno = new List<T>();
                T temp = new T();

                string cmd = string.Format("SELECT * FROM {0}", temp.Tabla);

                try
                {
                    using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                    {
                        Conexion.Open();
                        isConected = true;
                        DbDataReader lectura = await comando.ExecuteReaderAsync().ConfigureAwait(false);

                        while (lectura.Read())
                        {
                            
                            T buffer = new T();
                            for (int i = 0; i < buffer.Propiedades.Count; i++)
                            {
                                if ("Tabla" != buffer.Propiedades[i].Name && "Propiedades" != buffer.Propiedades[i].Name)
                                {



                                    buffer.Propiedades[i].SetValue(buffer, Convert.ChangeType(lectura[buffer.Propiedades[i].Name], buffer.Propiedades[i].PropertyType));


                                }

                            }
                            
                            retorno.Add(buffer);

                        }
                        this.estado = Estados.Operacion_Exitosa;
                        return retorno;

                    }
                }
                catch (Exception Error)
                {
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                    this.estado = Estados.Operacion_Fallida;
                    return new List<T>();
                    

                }
                finally
                {

                    Conexion.Close();
                    isConected = false;

                }
            }
        }

        public async void UpdateAsyn(IDateable Viejo, IDateable Nuevo)
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {



                try
                {
                    using (OleDbCommand command = new OleDbCommand(UpdateToCommand(Viejo, Nuevo), Conexion))
                    {
                        Conexion.Open();
                        IsConnected = true;
                        await command.ExecuteNonQueryAsync();
                        this.estado = Estados.Operacion_Exitosa;
                    }
                }
                catch (Exception Error)
                {
                    this.estado = Estados.Operacion_Fallida;
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);

                }
                finally
                {
                    
                    Conexion.Close();
                    isConected = false;
                }
            }
        }

        public async void CreateTable<T>() where T : IDateable, new()
        {

            if (Estado == Estados.SinCadenaDeConexion)
            {
                estado = Estados.Operacion_Fallida;
                throw new Exception("la cadena de conexion esta vacia, agregue una cadena de conexion");
            }
            else
            {



                try
                {
                    T buffer = new T();
                    string cmd = string.Format("CREATE TABLE {0} (", buffer.Tabla);

                    for (int i = 0; i < buffer.Propiedades.Count; i++)
                    {
                        if (buffer.Propiedades[i].Name != "Tabla" && buffer.Propiedades[i].Name != "Propiedades")
                        {
                            cmd += string.Format("[{0}] {1},", buffer.Propiedades[i].Name, CreateToCommand(buffer.Propiedades[i]));
                        }
                        
                    }
                    cmd.Remove(cmd.Length - 1);
                    cmd += ");";

                    using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                    {
                        Conexion.Open();
                        isConected = true;
                        await comando.ExecuteNonQueryAsync();
                        estado = Estados.Operacion_Exitosa;

                    }

                }
                catch (Exception Error)
                {
                    estado = Estados.Operacion_Fallida;
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);

                }

                finally
                {
                    Conexion.Close();
                    isConected = false;
                }

            }

        }

        public async void CreateTable(string cadenaDeConexion, IDateable tabla)
        {

            try
            {
                string cmd = string.Format("CREATE TABLE {0} (", tabla.Tabla);

                for (int i = 0; i < tabla.Propiedades.Count; i++)
                {
                    if (tabla.Propiedades[i].Name != "Tabla" && tabla.Propiedades[i].Name != "Propiedades")
                    {
                        if (i == tabla.Propiedades.Count - 1)
                        {
                            cmd += string.Format("[{0}] {1}", tabla.Propiedades[i].Name, CreateToCommand(tabla.Propiedades[i]));
                        }
                        else
                        {
                            cmd += string.Format("[{0}] {1},", tabla.Propiedades[i].Name, CreateToCommand(tabla.Propiedades[i]));
                        }
                    }

                }
              
                cmd += ");";

                using (OleDbConnection conexionTemp = new OleDbConnection(cadenaDeConexion))
                {
                    using (OleDbCommand comandon = new OleDbCommand(cmd, conexionTemp))
                    {
                        try
                        {
                            conexionTemp.Open();
                            isConected = true;
                            await comandon.ExecuteNonQueryAsync();
                            estado = Estados.Operacion_Exitosa;
                        }
                        catch (Exception Error)
                        {
                            estado = Estados.Operacion_Fallida;
                            new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                        }
                        finally
                        {
                            conexionTemp.Close();
                            isConected = false;
                        }
                        
                    }
                }
            }
            catch ( Exception error)
            {

                throw error;
            }
           
        }

        public async void CreateTable(OleDbConnection cadenaDeConexion, IDateable tabla)
        {

            try
            {
                string cmd = string.Format("CREATE TABLE {0} (", tabla.Tabla);

                for (int i = 0; i < tabla.Propiedades.Count; i++)
                {
                    if (tabla.Propiedades[i].Name != "Tabla" && tabla.Propiedades[i].Name != "Propiedades")
                    {
                        if (i == tabla.Propiedades.Count - 1)
                        {
                            cmd += string.Format("[{0}] {1}", tabla.Propiedades[i].Name, CreateToCommand(tabla.Propiedades[i]));
                        }
                        else
                        {
                            cmd += string.Format("[{0}] {1},", tabla.Propiedades[i].Name, CreateToCommand(tabla.Propiedades[i]));
                        }
                    }
                  
                }
              
                cmd += ");";

                try
                {
                    using (OleDbCommand comandon = new OleDbCommand(cmd, cadenaDeConexion))
                    {

                        cadenaDeConexion.Open();
                        isConected = true;
                        await comandon.ExecuteNonQueryAsync();
                        estado = Estados.Operacion_Exitosa;

                    }
                }
                catch (Exception Error)
                {
                    estado = Estados.Operacion_Fallida;
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                }
                finally
                {
                    cadenaDeConexion.Close();
                    isConected = false;
                }
               
                   
                
            }
            catch (Exception error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(error.ToString());
            }
            finally
            {
                cadenaDeConexion.Close();
            }

        }

        public void CreateDataBase(string Ubicacion)
        {
            string database = string.Format("Provider={0};Data Source=\"{1}\";Jet OLEDB:Engine Type=5", this.Provider,Ubicacion);

         
            try
            {
                catalog.Create(database);
                this.CadenaDeConexion = database;
                estado = Estados.Operacion_Exitosa;
            }
            catch (Exception Error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);

            }
        }
      
        public async void CreateDataBase(string Ubicacion, params IDateable[] Tablas)
        {



            await Task.Factory.StartNew(() => {

                try
                {
                    if (!File.Exists(Ubicacion))
                    {



                        CreateDataBase(Ubicacion);
                        string database = string.Format("Provider={0};Data Source=\"{1}\"", Provider, Ubicacion);
                        this.CadenaDeConexion = database;
                        for (int i = 0; i < Tablas.Length; i++)
                        {
                            CreateTable(database, Tablas[i]);
                        }

                    }
                    else
                    {
                        string database = string.Format("Provider={0};Data Source=\"{1}\"", Provider, Ubicacion);
                        this.cadenaDeConexionString = database;
                        using (OleDbConnection con = new OleDbConnection(database))
                        {
                            for (int i = 0; i < Tablas.Length; i++)
                            {
                                CreateTable(con, Tablas[i]);
                            }

                            con.Close();
                            isConected = false;
                        }
                    }
                    estado = Estados.Operacion_Exitosa;
                }
                catch (Exception Error)
                {
                    estado = Estados.Operacion_Fallida;
                    new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                }
                finally
                {
                    Conexion.Close();
                    isConected = false;
                }
             


            });
            
          
           

           

        }

        public async void CreateCampo(string tabla, PropertyInfo campo)
        {
            string cmd = string.Format("ALTER TABLE {0} ADD {1} {2};",tabla,campo.Name, CreateToCommand(campo) );
            try
            {
                using (OleDbCommand comando = new OleDbCommand(cmd, Conexion))
                {

                    Conexion.Open();
                    isConected = true;
                    await comando.ExecuteNonQueryAsync();
                    estado = Estados.Operacion_Exitosa;
                }
            }
            catch (Exception Error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
            }
            finally
            {
                Conexion.Close();
                isConected = false;
            }
           


        }

        public void NormalizeDatabase(params IDateable[] Tablas)
        {
            try
            {

            

                  for (int i = 0; i < Tablas.Length; i++)
                  {
                      NormalizeTable(Tablas[i]);
                  }
                estado = Estados.Operacion_Exitosa;
           
                
                

            }
            catch (Exception Error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
            }
            finally
            {
                this.Conexion.Close();
                isConected = false;
            }
        }

        public void NormalizeTable(IDateable modelo)
        {

            try
            {
                if (TableExists(modelo))
                {
                    foreach (PropertyInfo item in modelo.Propiedades)
                    {
                        if (!(item.Name == "Tabla" || item.Name == "Propiedades"))
                        {
                            if (!CampoExist(modelo, item))
                            {
                                CreateCampo(modelo.Tabla, item);
                            }
                        }

                    }
                  
                }
                else
                {
                    CreateTable(this.Conexion, modelo);
                }
                estado = Estados.Operacion_Exitosa;

            }
            catch (Exception Error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
            }
           


            
            

        }

        public IEnumerable<string> GetProviders()
        {
           
            OleDbEnumerator oleDbEnumerator = new OleDbEnumerator();

            var datos = oleDbEnumerator.GetElements();

            List<string> datosderelleno = new List<string>();

            for (int i = 0; i < datos.Rows.Count; i++)
            {
                yield return datos.Rows[i]["SOURCES_NAME"].ToString();
            }

            yield break;
        }
        #endregion

        #region Funciones internas
        private string InsertToCommmand(IDateable dateable)
        {
            
            string buffer = string.Format("INSERT INTO {0} (",dateable.Tabla);
            for (int i = 0; i < dateable.Propiedades.Count; i++)
            {
              
                
                    if ("Id" != dateable.Propiedades[i].Name && "Tabla" != dateable.Propiedades[i].Name && "Propiedades" != dateable.Propiedades[i].Name)
                    {
                        buffer += string.Format("{0},", dateable.Propiedades[i].Name);
                    }

               


            }

            buffer = buffer.Remove(buffer.Length - 1);
            buffer += ")";
            buffer += " VALUES(";

            for (int i = 0; i < dateable.Propiedades.Count; i++)
            {
                
                    if ("Id" != dateable.Propiedades[i].Name && "Tabla" != dateable.Propiedades[i].Name && "Propiedades" != dateable.Propiedades[i].Name)
                    {

                    if (dateable.Propiedades[i].PropertyType == typeof(DateTime))
                    {

                        DateTime date = Convert.ToDateTime(dateable.Propiedades[i].GetValue(dateable));
                        string bf = date.ToString("yyyy-MM-dd HH:mm:ss");
                        buffer += string.Format("\"{0}\",", bf);

                    }
                    else
                    {
                        if (dateable.Propiedades[i].PropertyType == typeof(string))
                        {
                            buffer += string.Format("\"{0}\",", dateable.Propiedades[i].GetValue(dateable).ToString());
                        }
                        else
                        {
                            buffer += string.Format("{0},", dateable.Propiedades[i].GetValue(dateable).ToString());
                        }
                        
                    }

                        
                    }
               

                
            }

            buffer = buffer.Remove(buffer.Length - 1);
            buffer += ");";

            return buffer;

        }

        private string UpdateToCommand(IDateable Viejo, IDateable Nuevo)
        {
            string buffer = string.Format("UPDATE {0} SET ", Nuevo.Tabla);
            if (Viejo.Tabla == Nuevo.Tabla)
            {
                for (int i = 0; i <Nuevo.Propiedades.Count; i++)
                {
                    if ("Id" != Nuevo.Propiedades[i].Name && "Tabla" != Nuevo.Propiedades[i].Name && "Propiedades" != Nuevo.Propiedades[i].Name)
                    {
                       
                            
                        if (Nuevo.Propiedades[i].PropertyType == typeof(DateTime))
                        {

                            DateTime date = Convert.ToDateTime(Nuevo.Propiedades[i].GetValue(Nuevo));
                            string bf = date.ToString("yyyy-MM-dd HH:mm:ss");
                            buffer += string.Format("{0}=\'{1}\', ", Nuevo.Propiedades[i].Name, bf);

                        }
                        else
                        {
                            if (Nuevo.Propiedades[i].PropertyType == typeof(string))
                            {
                                buffer += string.Format("{0}=\"{1}\", ",Nuevo.Propiedades[i].Name,Nuevo.Propiedades[i].GetValue(Nuevo).ToString());
                            }
                            else
                            {
                                buffer += string.Format("{0}={1}, ",Nuevo.Propiedades[i].Name ,Nuevo.Propiedades[i].GetValue(Nuevo).ToString());
                            }

                        }

                    }
                 
                }

                buffer = buffer.Remove(buffer.Length - 2);



                buffer += string.Format(" WHERE Id={0};", Viejo.Id);
                return buffer;
            }
            else
            {
                return "";
                throw new Exception("Los objetos no pertenecen a la misma tabla.");
            }
            
        }
         
        private string CreateToCommand(PropertyInfo propiedad)
        {

            if (propiedad.Name == "Id")
            {

                return "COUNTER PRIMARY KEY";

            }
            else
            {
                if (propiedad.PropertyType == typeof(string))
                {
                    return "TEXT";
                }
                else if (propiedad.PropertyType == typeof(int) || propiedad.PropertyType == typeof(Int16) || propiedad.PropertyType == typeof(Int64))
                {
                    return "INTEGER";
                }
                else if (propiedad.PropertyType == typeof(double))
                {
                    return "DOUBLE";
                }
                else if (propiedad.PropertyType == typeof(bool))
                {
                    return "BIT";
                }
                else if (propiedad.PropertyType == typeof(DateTime))
                {
                    return "DATETIME";
                }
                else if (propiedad.PropertyType == typeof(byte))
                {
                    return "BYTE";
                }
                else if (propiedad.PropertyType == typeof(float))
                {
                    return "FLOAT";
                }
                else if (propiedad.PropertyType == typeof(decimal))
                {
                    return "DECIMAL";
                }
                else
                {
                    return "BINARY";
                }
            }


        }

        private bool TableExists(IDateable Tabla)
        {
            try
            {
                Conexion.Open();
                isConected = true;
                var exists = Conexion.GetSchema("Tables", new string[4] { null, null, Tabla.Tabla, "TABLE" }).Rows.Count > 0;
                return exists;
            }
            catch (Exception Error)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(Error.ToString() + Environment.NewLine + Error.Message);
                return false;
              
                
            }
            finally
            {
                Conexion.Close();
                isConected = false;
                
            }
         
            
           
        }

        private bool CampoExist(IDateable tabla,PropertyInfo campo )
        {
            try
            {
                bool prueba = false;
                Conexion.Open();
                isConected = true;
                var t = Conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tabla.Tabla,null});
              
                for (int i = 0; i < t.Rows.Count; i++)
                {
                   
                        prueba = prueba || t.Rows[i].ItemArray[3].ToString().Contains(campo.Name);
                    
                 
                }
                Conexion.Close();
                isConected = false;
                return prueba;
            }
            catch (Exception exception)
            {
                estado = Estados.Operacion_Fallida;
                new LogInternal(exception.ToString());
                return false;
             
            }
            finally
            {
                Conexion.Close();
                isConected = false;
            }
          



        }
        #endregion

        #region IDisposable Support

        private bool disposedValue = false; // Para detectar llamadas redundantes

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: elimine el estado administrado (objetos administrados).
                    Conexion.Dispose();
                    if (ConexionInterna != null)
                    {
                        ConexionInterna.Close();
                    }
                   
                }

                ConexionInterna = null;

                catalog = null;

                

                // TODO: libere los recursos no administrados (objetos no administrados) y reemplace el siguiente finalizador.
                // TODO: configure los campos grandes en nulos.

                disposedValue = true;
            }
        }

        // TODO: reemplace un finalizador solo si el anterior Dispose(bool disposing) tiene código para liberar los recursos no administrados.
        // ~CoreDataBaseAccess() {
        //   // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
        //   Dispose(false);
        // }

        // Este código se agrega para implementar correctamente el patrón descartable.
        public void Dispose()
        {
            // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
            Dispose(true);
            // TODO: quite la marca de comentario de la siguiente línea si el finalizador se ha reemplazado antes.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
