using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using Atrybut;
using System.Collections.Generic;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Data;
using cdn_api;
using static Atrybut.FS;

[assembly: CallbackAssemblyDescription(Name: "ONE-042 Atrybut wysyłania FS",
                                           Description: "Dodanie atrybutu Automat - Zatwierdzenie wysyłki FS  do wysyłki FS",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2023.2.2",
                                           Created: "16-05-2024")]


namespace ZZnaCSV
{
    [SubscribeProcedure(Procedures.ListaDokumentow, "Atrybut wysyłania FS")]
    public class ZSnaFS_Class : Callback
    {
        private int XLWersjaAPI = 20232;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.1.202";
        private string XLSQLDatabase = "TOMMA-ERP";//"ERPXL_PoloMarket"
        private string XLUser = "MACIEJ1"; //
        private string XLPass = "NE!w0rk12";//"Dupa2003!";
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private List<ListaFS> FSy = new List<ListaFS>();
        int listID;
        public override void Init()
        {

            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?TabHan"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            RapButton.TextRaw = "Wyślij";
            RapButton.ToolTipRaw = "Wysłanie FS";
            //RapButton.IconRaw = "~EXPORT20.ICO";

            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 28, filtrBtn.Bounds.Y, 25, 15);
            listID = GetWindow().AllChildren["?Handlowe"].Id;      //tu sprawdzic

            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);

            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 28, filtrBtn.Bounds.Y, 25, 15);

            return true;
        }

        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {

            try
            {
                string connectionString = "";
                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 16) == "ERPXL_Tomma_test")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_Tomma_test; User ID = sa; Password=fr$g4i9sSa032V,_";
                    XLSQLDatabase = "TOMMA_TEST";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_TOMMA-ERP")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_TOMMA-ERP; User ID = sa; Password=fr$g4i9sSa032V,_";
                    XLSQLDatabase = "TOMMA";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 18) == "ERPXL_BONUS_2_test")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_BONUS_2_test; User ID = sa; Password=fr$g4i9sSa032V,_";
                    XLSQLDatabase = "BONUS 2-TEST";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 13) == "ERPXL_BONUS_2")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_BONUS_2; User ID = sa; Password=fr$g4i9sSa032V,_";
                    XLSQLDatabase = "BONUS 2";
                }

                ADODB._Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcId, listID, GetCallbackThread());
                
                ZbieranieFS(recordset);

               
                 if (FSy.Count != 0)
                 {
                     string komunikatBufor = "";
                     string komunikatEmail = "";
                    string komunikatWyslano = "";
                    string komunikatWyslanoWczesniej = "";
                    string komunikatNieDoWyslania = "";
                     XLLoginInfo_20232 nowyUzytkownik = new XLLoginInfo_20232();
                     nowyUzytkownik.Wersja = XLWersjaAPI;
                     nowyUzytkownik.Winieta = XLWinieta;
                     nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                     nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                     nowyUzytkownik.SerwerKlucza = XLSQLServer;
                     nowyUzytkownik.Baza = XLSQLDatabase;
                     nowyUzytkownik.OpeIdent = XLUser;
                     nowyUzytkownik.OpeHaslo = XLPass;
                     nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                     nowyUzytkownik.RodzajSerweraKlucza = 0;
                     nowyUzytkownik.SerwerKlucza = "COMARCHBD::5000024502";
                     nowyUzytkownik.ProgramID = "ONE-042 Atrybut wysyłania FS";

                     int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                   // MessageBox.Show("sesja " + sesja);

                     if (blad != 0)
                     {
                         MessageBox.Show("Błąd logowania: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }



                     foreach (var item in FSy)
                     {
                        string stanQuery = "SELECT TRn_Stan from cdn.TraNag WHERE TRN_GIDTyp=" + item.GIDTyp + " and TrN_GIDNumer=" + item.GIDNumer;
                        int stan = Convert.ToInt32(ConnectToSQL(connectionString, stanQuery));

                        string numerQuery = "SELECT cdn.NumerDokumentu(TrN_GIDTyp, TrN_SpiTyp, TrN_TrNTyp, TrN_TrNNumer, TrN_TrNRok, TrN_TrNSeria, TrN_TrNMiesiac) FROM cdn.TraNag " +
                                                " WHERE TRN_GIDTyp=" + item.GIDTyp + " and TrN_GIDNumer=" + item.GIDNumer;
                        string numer = ConnectToSQL(connectionString, numerQuery);

                        string czyEmailQuery = "if exists(SELECT  Knt_Email from cdn.KntKarty join cdn.TraNag on TrN_KntNumer = Knt_GIDNumer and TrN_KntTyp = Knt_GIDTyp " +
                                                " WHERE Knt_Email!= '' and TrN_GIDTyp = " + item.GIDTyp + " and TrN_GIDNumer = " + item.GIDNumer + ") SELECT 'TAK' else SELECT 'NIE'";
                        string czyEmail = ConnectToSQL(connectionString, czyEmailQuery);
                        if (stan<3)
                        {
                            komunikatBufor +="\n" +numer;
                        }
                        else if(czyEmail.Trim()=="NIE")
                        {
                            komunikatEmail += "\n" + numer;
                        }
                        else
                        {
                            string atrybutQuery = "if exists(SELECT * FROM cdn.Atrybuty join cdn.AtrybutyKlasy on Atk_Id = Atr_AtkId " +
                                                " WHERE Atk_Nazwa = 'Automat - Zatwierdzenie Wysyłki FS' and Atr_ObiTyp = " + item.GIDTyp + " and Atr_ObiNumer = " + item.GIDNumer + ") SELECT Atr_Wartosc FROM cdn.Atrybuty" +
                                                " join cdn.AtrybutyKlasy on Atk_Id = Atr_AtkId WHERE Atk_Nazwa = 'Automat - Zatwierdzenie Wysyłki FS' and Atr_ObiTyp = " + item.GIDTyp + " and Atr_ObiNumer = " + item.GIDNumer +
                                                " else SELECT 'NIE MA'";
                            string atrybut = ConnectToSQL(connectionString, atrybutQuery).Trim();
                            
                            if(atrybut=="TAK")
                            {
                                komunikatWyslanoWczesniej += "\n" + numer;
                            }
                            else if (atrybut=="NIE MA")
                            {
                                XLAtrybutInfo_20232 atrybut5 = new XLAtrybutInfo_20232();
                                atrybut5.Wersja = XLWersjaAPI;
                                atrybut5.GIDFirma = item.GIDFirma;
                                atrybut5.GIDLp = item.GIDLp;
                                atrybut5.GIDNumer = item.GIDNumer;
                                atrybut5.GIDSubLp = 0;
                                atrybut5.GIDTyp = item.GIDTyp;
                                atrybut5.Wartosc = "TAK";
                                atrybut5.Klasa = "Automat - Zatwierdzenie Wysyłki FS";
                                int blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut5);

                                if (blad_atr != 0)
                                {
                                    MessageBox.Show("Błąd atrybutu: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    komunikatWyslano += "\n" + numer;
                                }

                                XLAtrybutInfo_20232 atrybut6 = new XLAtrybutInfo_20232();
                                atrybut6.Wersja = XLWersjaAPI;
                                atrybut6.GIDFirma = item.GIDFirma;
                                atrybut6.GIDLp = item.GIDLp;
                                atrybut6.GIDNumer = item.GIDNumer;
                                atrybut6.GIDSubLp = 0;
                                atrybut6.GIDTyp = item.GIDTyp;
                                atrybut6.Wartosc = "";
                                atrybut6.Klasa = "Automat - Data Wysyłki FS";
                                blad_atr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut6);

                                if (blad_atr != 0)
                                {
                                    MessageBox.Show("Błąd atrybutu: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }

                            }
                            else
                            {
                                komunikatNieDoWyslania += "\n" + numer;
                            }
                        }

                     }
                    // komunikat += "wysłano upomnienie";
                     //komunikat2 += "nie wysłano upomnień z powodu braku adresu E-mail";

                     cdn_api.cdn_api.XLLogout(sesja);

                     MessageBox.Show((komunikatWyslano != "" ? "Dla dokumentów:" + komunikatWyslano + "\nwysłano faktury\n\n" : "")
                                    + (komunikatWyslanoWczesniej != "" ? "Dla dokumentów:" + komunikatWyslanoWczesniej + "\nzostały już wysłane faktury\n\n" : "")
                                    + (komunikatNieDoWyslania != "" ? "Dla dokumentów:" + komunikatNieDoWyslania + "\nnie wysłano -> atrybut oznaczony jako nie do wysłania\n\n" : "")
                                    + (komunikatEmail != "" ? "Dla dokumentów:" + komunikatEmail + "\nnie wysłano -> brak maila\n\n" : "") 
                                    + (komunikatBufor!=""?"Dla dokumentów:"+komunikatBufor+"\nnie wysłano -> są w buforze":"")
                                    , "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 FSy.Clear();
                


            }
            catch { }
            
            return true;
        }



        bool ZbieranieFS(ADODB._Recordset recordset)
        {

            try
            {
                StringBuilder dokumenty = new StringBuilder();
                if (recordset != null)
                {
                    int count = recordset.RecordCount;

                    if (count != 0)
                    {
                        recordset.MoveFirst();

                        while (recordset.EOF == false)
                        {
                            ADODB.Fields fields = recordset.Fields;
                            if (Convert.ToInt16(fields[0].Value) ==2033)
                            {
                                FSy.Add(new ListaFS()
                                {
                                    GIDTyp = Convert.ToInt16(fields[0].Value),
                                    GIDFirma = Convert.ToInt32(fields[1].Value),
                                    GIDNumer = Convert.ToInt32(fields[2].Value),
                                    GIDLp = Convert.ToInt16(fields[3].Value),
                                });
                            }
                            recordset.MoveNext();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nie zaznaczono dokumentów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("Nie zaznaczono dokumentów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd pobierania zaznaczonych ZS --> " + ex.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }


        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        public override void Cleanup()
        {

        }

    }
}
