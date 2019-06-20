using ASK;
using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace АРМ_Руководителя_проекта
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    /// 

    public interface IDTRKCNom

    {
        DataTable GetData(string[] lot);
    }

    public class DTRKCNom : IDTRKCNom
    {
        public DataTable GetData(string[] lot)
        {
            DataTable ret = new DataTable { TableName = "MyTable" };
            ret.Columns.Add("Kod");
            ret.Columns.Add("Name");
            String sql = @"SELECT KodRKC as [Kod], (NRKC +' | ' + left(NameRKC,70)) as Name FROM RKC WHERE LotKod=" + lot[0] + " order by LotKod, Npp";
            OleDbConnection connection = null;
            try
            {
                connection = new OleDbConnection(MainWindow.GetConnectionString());
                OleDbCommand Command = new OleDbCommand(sql, connection);
                OleDbDataAdapter adRKC = new OleDbDataAdapter(Command);
                connection.Open();
                adRKC.Fill(ret);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); }

            }
            return ret;
        }
    } //Номер строки РКЦ прописывается в раскрывающийся список

    public class DTKSNom : IDTRKCNom
    {
        public DataTable GetData(string[] lot)
        {
            DataTable ret = new DataTable { TableName = "MyTable" };
            ret.Columns.Add("Kod");
            ret.Columns.Add("Name");
            String sql = @"SELECT КС.КодКС as [Kod], (КС.НомерКС & ' | ' & КС.ДатаДокумента & ' | ' & КС.Филиал) as Name FROM КС WHERE (КС.Лот)=" + lot[0] + " ORDER BY КС.[ДатаДокумента] DESC;";
            OleDbConnection connection = null;
            try
            {
                connection = new OleDbConnection(MainWindow.GetConnectionString());
                OleDbCommand Command = new OleDbCommand(sql, connection);
                OleDbDataAdapter adRKC = new OleDbDataAdapter(Command);
                connection.Open();
                adRKC.Fill(ret);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); }

            }
            return ret;
        }
    } //Номер КСки прописывается в раскрывающийся список

    public class MyDataТип : System.Collections.ObjectModel.ObservableCollection<string>
    {
        public MyDataТип()
        {
            Add("");
            Add("Заголовок");
            Add("Итого");
            Add("Работы (МСГ)");
            Add("МТР");
            Add("ОНМ / ЗИП");
            Add("ПНР");
        }
    } //Типы строк определяются в раскрывающемся списке

    public partial class MainWindow : Window
    {

        private Boolean initFlg = false;
        private Boolean MP_flg = true;
        private Boolean KSP_flg = true;
        private String connectionString;
        private OleDbDataAdapter adRKC;
        private OleDbDataAdapter adPunkt;
        private OleDbDataAdapter adKS;
        private OleDbDataAdapter adProbl;
        private OleDbDataAdapter adMeropr;
        private DataTable DGРКЦTable;
        private DataTable DGПроблемыTable;
        private DataTable DG2Table;
        private DataTable DGПМTable;
        private DataTable DG3Table;
        private double refinans;
        public string TipVText = "";
        private string NomLota = "";
        private string NomLPunkta = "";
        public Object[,] lot_p = new Object[2, 1];

        public MainWindow()
        {
            CultureInfo myCIintl = new CultureInfo("ru-RU", false);
            System.Threading.Thread.CurrentThread.CurrentUICulture = myCIintl;
            System.Threading.Thread.CurrentThread.CurrentCulture = myCIintl;

            InitializeComponent();

            connectionString = GetConnectionString();
            string sql;
            // Declares managed prototypes for unmanaged functions.
            long i, j;
            this.Visibility = Visibility.Hidden;
            this.TaskbarItemInfo.ProgressValue = 0;
            this.TaskbarItemInfo.Description = "Загрузка начальных данных";
            this.TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
            UserName.Text = "Пользователь " + (string)Environment.UserName;
            this.TaskbarItemInfo.ProgressValue = 0.1;
            sql = @"SELECT ПРП.[лот], ""Лот "" & [Лоты]![Номер лота] & "" Контракт № "" & [Номер договора] & "" от "" & [Дата договора] AS Выражение1, ПРП.[all]  FROM Лоты INNER JOIN ПРП ON Лоты.Код = ПРП.[лот] WHERE (((ПРП.User)=""" + (string)Environment.UserName + @"""));";
            Object[,] rez1 = ConnectToData(connectionString, sql);
            if (rez1.Length < 1)
            {
                WinNoDost w = new WinNoDost();
                w.SetText("Пользователь " + (string)Environment.UserName + Environment.NewLine + "не имеет права работать с данным файлом." + Environment.NewLine + "Прошу сообщить руководителю департамента управления" + Environment.NewLine + "проектом об этом для включения в список.");
                w.ShowDialog();
                this.Close();
            }
            else
            {
                int rows = rez1.GetUpperBound(0) + 1;
                int[] temp = new int[2];
                temp[0] = 2;
                temp[1] = rows + 5;
                lot_p = (Object[,])ResizeArray(lot_p, temp);
                Lot.Items.Clear();
                Lot.Items.IsLiveSorting = true;
                for (i = 0; i < rows; i++)
                {
                    j = Lot.Items.Add(rez1[i, 1]);
                    lot_p[0, j] = rez1[i, 0];
                    lot_p[1, j] = rez1[i, 1];
                }
                Form1.TaskbarItemInfo.ProgressValue = 0.2;
                Object[,] rez = ConnectToData(connectionString, @"Select Настройки.Значение FROM Настройки WHERE (((Настройки.Название)=""Ставка рефенансирования""));");
                Form1.TaskbarItemInfo.ProgressValue = 0.4;
                refinans = (double)rez[0, 0];
                Refinans_text.Text = String.Format("Ставка рефинансирования ЦБ {0:p}", refinans);
                Form1.TaskbarItemInfo.ProgressValue = 0.6;

                Form1.TaskbarItemInfo.ProgressValue = 1;
                Form1.TaskbarItemInfo.Description = "";
                TipV.Text = "График / РКЦ";
                Form1.Visibility = Visibility.Visible;
                initFlg = true;

                SizeC(true, true);

                Lot.SelectedIndex = 0;// ' Вызывает Lot_SelectionChanged и соответственно ReadDBtoDG

                Form1.TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;

            }

        }

        public static int LotNom;

        public int LotNom1()
        {
            int i = 0;
            try
            {
                i = (int)lot_p[0, Lot.SelectedIndex];
            }
            catch (Exception)
            {

                //throw;
            }
            return i;
        }

        private void ABoxItem_Click(object sender, RoutedEventArgs e)
        {

        }

        public static string GetConnectionString()
        {
            return @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\ORP.mdb";
        }

        private static Array ResizeArray(Array arr, int[] newSizes)
        {
            if (newSizes.Length != arr.Rank)
                throw new ArgumentException("arr must have the same number of dimensions " +
                                            "as there are elements in newSizes", "newSizes");

            var temp = Array.CreateInstance(arr.GetType().GetElementType(), newSizes);
            int length = arr.Length <= temp.Length ? arr.Length : temp.Length;
            Array.ConstrainedCopy(arr, 0, temp, 0, length);
            return temp;
        }

        private static Object[,] ConnectToData(String connectionString, String SQLStr)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {

                int i, j, k1;
                OleDbCommand Command = new OleDbCommand(SQLStr, connection);
                connection.Open();
                Object[,] nas_dat = new object[1, 1];
                i = 0;
                int[] temp = new int[2];
                OleDbDataReader reader = Command.ExecuteReader();
                while (reader.Read())
                {
                    k1 = reader.VisibleFieldCount;
                    temp[0] = i + 1;
                    if (k1 > temp[1]) { temp[1] = k1; }
                    nas_dat = (Object[,])ResizeArray(nas_dat, temp);
                    for (j = 0; j < k1; j++) { nas_dat[i, j] = reader[j]; }
                    i++;
                }
                if (i == 0)
                {
                    nas_dat = new Object[0, 0];
                }
                reader.Close();
                connection.Close();
                return nas_dat;
            }
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            SizeC(e.WidthChanged, e.HeightChanged);
        }

        private void SizeC(Boolean w, Boolean h)
        {
            if (initFlg)
            {
                if (w)
                {
                    Double l = GridMW.ActualWidth; //'Form1.ActualWidth - 25
                    if (l < 20) { l = 30; }
                    MMenu.Width = l;
                    DP.Width = l - 10;
                    if (TipVText == "График / РКЦ")
                    {
                        DGРКЦ.Width = l - 12;

                        DGПроблемы.Width = 0;
                        Gdop.Width = l - 12;

                        DG2.Width = MP_flg ? (l - 12) / 2 : 0;
                        DG3.Width = KSP_flg ? (l - 12) / 2 : 0;
                        DGПМ.Width = 0;
                    }
                    else if (TipVText == "Проблемные пункты / PID")
                    {
                        DGРКЦ.Width = 0;
                        Gdop.Width = l - 12;
                        DGПроблемы.Width = l - 12;
                        DG2.Width = 0;
                        DG3.Width = 0;
                        DGПМ.Width = MP_flg ? (l - 12) : 0;
                    }
                    SB.Width = l;
                    StatusText.Width = l - 20 - UserName.ActualWidth - Refinans_text.ActualWidth;
                }
                if (h)
                {
                    Double l = GridMW.ActualHeight - MMenu.ActualHeight - SB.ActualHeight; //'Form1.ActualWidth - 25
                    if (l < 20) { l = 30; }
                    DP.Height = l;
                    if (TipVText == "График / РКЦ")
                    {
                        Gdop.Height = MP_flg ? l / 3 : 0;
                        DG2.Height = MP_flg ? l / 3 : 0;
                        DG3.Height = KSP_flg ? l / 3 : 0;
                        DGПроблемы.Height = 0;
                        DGПМ.Height = 0;
                        DGРКЦ.Height = l - Gdop.Height;
                    }
                    else if (TipVText == "Проблемные пункты / PID")
                    {
                        Gdop.Height = MP_flg ? l / 3 : 0;
                        DGПМ.Height = MP_flg ? l / 3 : 0;
                        DG2.Height = 0;
                        DG3.Height = 0;
                        DGПроблемы.Height = l - Gdop.Height;
                        DGРКЦ.Height = 0;

                    }

                }
            }
        }
        private void UpdateDB1()
        {
            OleDbCommandBuilder comandbuilder = new OleDbCommandBuilder(adRKC);
            adRKC.Update(DGРКЦTable);
        }

        private void ReadDBtoDGРКЦ()
        {
            if (initFlg)
            {
                if (TipVText == "График / РКЦ")
                {
                    String sql = @"SELECT * FROM RKC WHERE LotKod=" + lot_p[0, Lot.SelectedIndex] + " order by LotKod, Npp";
                    DGРКЦTable = new DataTable();
                    OleDbConnection connection = null;
                    try
                    {
                        connection = new OleDbConnection(connectionString);
                        OleDbCommand Command = new OleDbCommand(sql, connection);
                        adRKC = new OleDbDataAdapter(Command);
                        Command = new OleDbCommand(@"INSERT INTO RKC (LotKod, Npp, NRKC, NameRKC, EI, Kol, Cena, SumRKC, StartD, EndD, Rasdel, Etap, PSD, Region, ObRKC, Prim, NDS, Tip ) " +
                            "VALUES (" + lot_p[0, Lot.SelectedIndex] + ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", connection);

                        Command.Parameters.Add("Npp", OleDbType.Integer, 40, "Npp");
                        Command.Parameters.Add("NRKC", OleDbType.WChar, 255, "NRKC");
                        Command.Parameters.Add("NameRKC", OleDbType.WChar, 0, "NameRKC");
                        Command.Parameters.Add("EI", OleDbType.WChar, 5, "EI");
                        Command.Parameters.Add("Kol", OleDbType.Double, 15, "Kol");
                        Command.Parameters.Add("Cena", OleDbType.Double, 15, "Cena");
                        Command.Parameters.Add("SumRKC", OleDbType.Double, 15, "SumRKC");
                        Command.Parameters.Add("StartD", OleDbType.Date, 15, "StartD");
                        Command.Parameters.Add("EndD", OleDbType.Date, 15, "EndD");
                        Command.Parameters.Add("Rasdel", OleDbType.WChar, 255, "Rasdel");
                        Command.Parameters.Add("Etap", OleDbType.WChar, 255, "Etap");
                        Command.Parameters.Add("PSD", OleDbType.WChar, 255, "PSD");
                        Command.Parameters.Add("Region", OleDbType.WChar, 255, "Region");
                        Command.Parameters.Add("ObRKC", OleDbType.WChar, 255, "ObRKC");
                        Command.Parameters.Add("Prim", OleDbType.WChar, 255, "Prim");
                        Command.Parameters.Add("NDS", OleDbType.Boolean, 1, "NDS");
                        Command.Parameters.Add("Tip", OleDbType.WChar, 25, "Tip");

                        adRKC.InsertCommand = Command;
                        adRKC.DeleteCommand = new OleDbCommand("DELETE FROM RKC WHERE KodRKC = ?");
                        adRKC.DeleteCommand.Parameters.Add("KodRKC", OleDbType.Guid, 5, "KodRKC").SourceVersion = DataRowVersion.Original;

                        Command = new OleDbCommand(@"SELECT KodRKC FROM RKC WHERE LotKod= " + lot_p[0, Lot.SelectedIndex] + " and Npp=? and NRKC='?' and  NameRKC='?' and  EI='?' and  Kol=? and  Cena=? and  SumRKC=? and "+
                            " StartD=#?# and  EndD=#?# and  Rasdel='?' and  Etap='?' and  PSD='?' and  Region='?' and  ObRKC='?' and  Prim='?' and  NDS=? and  Tip = '?' ;", connection);
                        Command.Parameters.Add("Npp", OleDbType.Integer, 40);
                        Command.Parameters.Add("NRKC", OleDbType.WChar, 255);
                        Command.Parameters.Add("NameRKC", OleDbType.WChar, 0);
                        Command.Parameters.Add("EI", OleDbType.WChar, 5);
                        Command.Parameters.Add("Kol", OleDbType.Double, 15);
                        Command.Parameters.Add("Cena", OleDbType.Double, 15);
                        Command.Parameters.Add("SumRKC", OleDbType.Double, 15);
                        Command.Parameters.Add("StartD", OleDbType.Date, 15);
                        Command.Parameters.Add("EndD", OleDbType.Date, 15);
                        Command.Parameters.Add("Rasdel", OleDbType.WChar, 255);
                        Command.Parameters.Add("Etap", OleDbType.WChar, 255);
                        Command.Parameters.Add("PSD", OleDbType.WChar, 255);
                        Command.Parameters.Add("Region", OleDbType.WChar, 255);
                        Command.Parameters.Add("ObRKC", OleDbType.WChar, 255);
                        Command.Parameters.Add("Prim", OleDbType.WChar, 255);
                        Command.Parameters.Add("NDS", OleDbType.Boolean, 1);
                        Command.Parameters.Add("Tip", OleDbType.WChar, 25);
                      //  adRKC.SelectCommand = Command;

                        connection.Open();
                        adRKC.Fill(DGРКЦTable);
                        DGРКЦ.ItemsSource = DGРКЦTable.DefaultView;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw;
                    }
                    finally
                    {
                        if (connection != null) { connection.Close(); }

                    }


                }
                else if (TipVText == "Проблемные пункты / PID")
                {
                    String sql = @"SELECT * FROM Punkt WHERE [LotNom]=" + lot_p[0, Lot.SelectedIndex] + " order by [NPunkta], NRKC";
                    DGПроблемыTable = new DataTable();
                    OleDbConnection connection = null;
                    try
                    {
                        connection = new OleDbConnection(connectionString);
                        OleDbCommand Command = new OleDbCommand(sql, connection);
                        adProbl = new OleDbDataAdapter(Command);
                        Command = new OleDbCommand(@"INSERT INTO Punkt (NPunkta, NamePunkta, StoimPunkta, SConst, SForm, SMin, SMax, DStart, Prim, LotNom, NRKC )" + "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, " + lot_p[0, Lot.SelectedIndex] + ", ? );", connection);
                        // '             Command.Parameters.Add("КодРКЦ", OleDbType.Guid, 16, "КодРКЦ")
                        Command.Parameters.Add("NPunkta", OleDbType.WChar, 255, "NPunkta");
                        Command.Parameters.Add("NamePunkta", OleDbType.WChar, 255, "NamePunkta");
                        Command.Parameters.Add("StoimPunkta", OleDbType.Currency, 15, "StoimPunkta");
                        Command.Parameters.Add("SConst", OleDbType.Currency, 15, "SConst");
                        Command.Parameters.Add("SForm", OleDbType.WChar, 255, "SForm");
                        Command.Parameters.Add("SMin", OleDbType.Currency, 15, "SMin");
                        Command.Parameters.Add("SMax", OleDbType.Currency, 15, "SMax");
                        Command.Parameters.Add("DStart", OleDbType.Date, 15, "DStart");
                        Command.Parameters.Add("Prim", OleDbType.WChar, 255, "Prim");
                        Command.Parameters.Add("NRKC", OleDbType.Guid, 15, "NRKC");
                        adProbl.InsertCommand = Command;
                        connection.Open();
                        adProbl.Fill(DGПроблемыTable);
                        DGПроблемы.ItemsSource = DGПроблемыTable.DefaultView;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw;
                    }
                    finally
                    {
                        if (connection != null) { connection.Close(); }

                    }


                }
                SizeC(true, true);

            }
        }


        private Boolean GuidR()
        {
            try
            {
                if ((initFlg) && (TipVText == "График / РКЦ"))
                {

                    Object t = DGРКЦ.CurrentCell.Item;
                }
                else if ((initFlg) && (TipVText == "Проблемные пункты / PID"))
                {
                    Object t = DGПроблемы.CurrentCell.Item;
                }
                return true;

            }
            catch (Exception)
            {
                return false;

                throw;
            }
        }

        private string GuidRS()
        {
            try
            {
                if ((initFlg) && (TipVText == "График / РКЦ"))
                {
                    DataRowView dvg = (DataRowView)DGРКЦ.CurrentCell.Item;
                    Object d = (Object)dvg.Row.ItemArray[0];
                    NomLota = d.ToString();
                    return NomLota;
                }
                else if ((initFlg) && (TipVText == "Проблемные пункты / PID"))
                {
                    DataRowView dvg = (DataRowView)DGПроблемы.CurrentCell.Item;
                    Object d = (Object)dvg.Row.ItemArray[0];
                    NomLPunkta = d.ToString();
                    return NomLPunkta;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception)
            {
                if ((initFlg) && (TipVText == "График / РКЦ"))
                {
                    return NomLota;
                }
                else if ((initFlg) && (TipVText == "Проблемные пункты / PID"))
                {
                    return NomLPunkta;
                }
                else
                {
                    return "";
                }

                throw;
            }
        }

        private void Lot_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ReadDBtoDGРКЦ();
            DTRKCNom t = new DTRKCNom();
            КодРКЦСт.ItemsSource = null;
            string[] t1 = { " ", " " };
            t1[0] = (string)"" + lot_p[0, Lot.SelectedIndex];
            КодРКЦСт.ItemsSource = t.GetData(t1).DefaultView;
        }


        private void TipV_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem t = (ComboBoxItem)e.AddedItems[0];
            TipVText = t.Content.ToString();
            LotNom = LotNom1();
            ReadDBtoDGРКЦ();

        }

        private void ReadDBtoDG2()
        {
            if (initFlg)
            {
                if (TipVText == "График / РКЦ")
                {
                    if (GuidR())
                    {
                        string s = GuidRS();
                        String sql = "SELECT РКЦ_пункты.* FROM РКЦ_пункты WHERE(((РКЦ_пункты.РКЦ) ={" + s + "})) ORDER BY РКЦ_пункты.Дата_окон, РКЦ_пункты.Дата_нач; ";
                        DG2Table = new DataTable();
                        OleDbConnection connection = null;
                        try
                        {
                            connection = new OleDbConnection(connectionString);
                            OleDbCommand Command = new OleDbCommand(sql, connection);
                            adPunkt = new OleDbDataAdapter(Command);
                            Command = new OleDbCommand(@"INSERT INTO РКЦ_пункты (РКЦ, Дата_нач, Дата_окон, Объем, Деньги) " + "VALUES ({" + s + "} , ?, ?, ?, ? )", connection);

                            // '             Command.Parameters.Add("КодРКЦ", OleDbType.Guid, 16, "КодРКЦ")
                            Command.Parameters.Add("Дата_нач", OleDbType.Date, 15, "Дата_нач");
                            Command.Parameters.Add("Дата_окон", OleDbType.Date, 15, "Дата_окон");
                            Command.Parameters.Add("Объем", OleDbType.Double, 15, "Объем");
                            Command.Parameters.Add("Деньги", OleDbType.Double, 15, "Деньги");


                            adPunkt.InsertCommand = Command;
                            connection.Open();
                            adPunkt.Fill(DG2Table);
                            DG2.ItemsSource = DG2Table.DefaultView;

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            throw;
                        }
                        finally
                        {
                            if (connection != null) { connection.Close(); }

                        }
                    }
                }
                else if (TipVText == "Проблемные пункты / PID")
                {
                    if (GuidR())
                    {
                        string s = GuidRS();
                        if (s == "") s = "0";
                        String sql = "SELECT Мероприятия.* FROM Мероприятия WHERE(Мероприятия.[Номер_пункта] =" + s + ") ORDER BY Срок, Название; ";
                        DGПМTable = new DataTable();
                        OleDbConnection connection = null;
                        try
                        {
                            connection = new OleDbConnection(connectionString);
                            OleDbCommand Command = new OleDbCommand(sql, connection);
                            adKS = new OleDbDataAdapter(Command);
                            /*Command = new OleDbCommand(@"INSERT INTO РКЦ_пункты (РКЦ, Дата_нач, Дата_окон, Объем, Деньги) " + "VALUES ({" + s + "} , ?, ?, ?, ? )", connection);

                            // '             Command.Parameters.Add("КодРКЦ", OleDbType.Guid, 16, "КодРКЦ")
                            Command.Parameters.Add("Дата_нач", OleDbType.Date, 15, "Дата_нач");
                            Command.Parameters.Add("Дата_окон", OleDbType.Date, 15, "Дата_окон");
                            Command.Parameters.Add("Объем", OleDbType.Double, 15, "Объем");
                            Command.Parameters.Add("Деньги", OleDbType.Double, 15, "Деньги");


                            adKS.InsertCommand = Command;*/
                            connection.Open();
                            adKS.Fill(DGПМTable);
                            DGПМ.ItemsSource = DGПМTable.DefaultView;

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            throw;
                        }
                        finally
                        {
                            if (connection != null) { connection.Close(); }

                        }

                    }
                }

                SizeC(true, true);
            }
        }
        private void ReadDBtoDG3()
        {
            if (initFlg)
            {
                if (TipVText == "График / РКЦ")
                {
                    if (GuidR())
                    {
                        string s = GuidRS();
                        String sql = "SELECT РКЦ_КС.* FROM РКЦ_КС WHERE(((РКЦ_КС.РКЦ) ={" + s + "})) ORDER BY РКЦ_КС.Дата_закрытия; ";
                        DG3Table = new DataTable();
                        OleDbConnection connection = null;
                        try
                        {
                            connection = new OleDbConnection(connectionString);
                            OleDbCommand Command = new OleDbCommand(sql, connection);
                            adMeropr = new OleDbDataAdapter(Command);
                            Command = new OleDbCommand(@"INSERT INTO РКЦ_пункты (РКЦ, Дата_закрытия, Объем, Деньги, НомерКС) " + "VALUES ({" + s + "} , ?, ?, ?, ? )", connection);

                            // '             Command.Parameters.Add("КодРКЦ", OleDbType.Guid, 16, "КодРКЦ")
                            Command.Parameters.Add("Дата_закрытия", OleDbType.Date, 15, "Дата закрытия");
                            Command.Parameters.Add("Объем", OleDbType.Double, 15, "Объем");
                            Command.Parameters.Add("Деньги", OleDbType.Double, 15, "Деньги");
                            Command.Parameters.Add("НомерКС", OleDbType.Guid, 15, "Номер КС");

                            adMeropr.InsertCommand = Command;
                            connection.Open();
                            adMeropr.Fill(DG3Table);
                            DG3.ItemsSource = DG3Table.DefaultView;
                            DTKSNom t = new DTKSNom();
                            КодКССт.ItemsSource = null;
                            string[] t1 = { " ", " " };
                            t1[0] = (string)"" + lot_p[0, Lot.SelectedIndex];
                            КодКССт.ItemsSource = t.GetData(t1).DefaultView;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            throw;
                        }
                        finally
                        {
                            if (connection != null) { connection.Close(); }

                        }


                    }
                    else
                    {
                        DG3.Columns.Clear();
                        DG3.ItemsSource = null;

                    }
                }
                SizeC(true, true);
            }
        }

        private void DGРКЦ_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            ReadDBtoDG2();
            ReadDBtoDG3();
        }

        private void DGРКЦ_CurrentCellChanged(object sender, EventArgs e)
        {
            ReadDBtoDG2();
            ReadDBtoDG3();
        }

        private void Form1_Closed(object sender, EventArgs e)
        {
            if (adRKC!=null && DGРКЦTable!=null)
            {
                adRKC.Update(DGРКЦTable);

            }
            if (adPunkt != null && DGПроблемыTable != null)
            {
                adPunkt.Update(DGПроблемыTable);
            }
        }

        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }

        private void CMDel_Click(object sender, RoutedEventArgs e)
        {
            if (initFlg)
            {
                if (TipVText == "График / РКЦ")
                {
                    if (GuidR())
                    {
                        string s = GuidRS();


                    }
                }
            }
        }
    }
}
            