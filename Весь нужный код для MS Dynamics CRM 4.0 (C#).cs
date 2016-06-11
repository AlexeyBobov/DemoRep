using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Net;
using System.Xml;
using System.Xml.Serialization;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Crm.Sdk;                        //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.Sdk.Metadata;               //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.Sdk.Query;                  //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.Workflow.Activities;        //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.Workflow.Services;          //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.SdkTypeProxy;               //Файл microsoft.crm.sdktypeproxy.dll.
using Microsoft.Crm.SdkTypeProxy.Metadata;      //Файл microsoft.crm.sdktypeproxy.dll.
using Microsoft.Crm.Outlook.Sdk;                //Файл microsoft.crm.outlook.sdk.dll.
using ConsoleApplication1.ServiceReference1;    //Веб-служба "http://crm4/mscrmservices/2007/crmservice.asmx".
using FirebirdSql.Data.FirebirdClient;          //Добавляем пространство имен firebird ado.net provider для взаимодействия с Firebird.

namespace ConsoleApplication1
{
    public class Program
    {
        static void Main(string[] args)
        {
            string prompt = "ALL THINGS I need to work.";
            Console.WriteLine(prompt);
            Console.WriteLine("");

            #region Работа с прочими данными.

            string var1 = "12wAVBGTer45- йцукенгЯЧСМИТЬ=7sdf12wKJHGT5611фываПРОЛД11!@#89000";
            string digvar1 = Regex.Replace(var1, @"[^0-9]", String.Empty);
            Console.WriteLine("{0}, {1}.", var1, digvar1);
            
            string var2 = "12wAVBGTer45- йцукенгЯЧСМИТЬ=7sdf12wKJHGT5611фываПРОЛД11!@#89000";
            string digvar2 = new String(var2.Where(Char.IsDigit).ToArray());
            Console.WriteLine("{0}, {1}.", var2, digvar2);

            //Разбиение на подстроки по символам-разделителям с отбрасыванием пробелов в начале полученных строк:
            string[] var3 = var2.Split('-', '=').Select(x => x.Trim()).ToArray();
            //Просто разбиение на подстроки по символам-разделителям:
            string[] var4 = var2.Split('=');

            //Приведение универсальной даты к местному времени (для правильного отображения с учетом разницы в 4 часа):
            CrmDateTime d = new CrmDateTime();
            string result_d = d.UniversalTime.ToLocalTime().ToString();

            //Дата в формате "Год, Месяц, День"; Год - сегодняшний, Месяц и День - из параметра birthdate Контакта:
            d = new DateTime(DateTime.Today.Year, Convert.ToInt32(r3.birthdate.Substring(5, 2)), Convert.ToInt32(r3.birthdate.Substring(8, 2)));

            //Условный оператор в одну строчку:
            string stringToUpdate = ((r.returned == "true") ? "В, " : "") + "Это добавится в любом случае.";

            #endregion

            #region Работа с файлами и папками.

            //Вывести все аргументы, заданные в командной строке при вызове программы из командной строки:
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine("Arg: {0} ", args[i]);
                Console.WriteLine("");
            }
            //Атрибуты текущей папки:
            DirectoryInfo dir1 = new DirectoryInfo(".");
            Console.WriteLine("FullName: {0}", dir1.FullName);
            Console.WriteLine("Name: {0}", dir1.Name);
            Console.WriteLine("Parent: {0}", dir1.Parent);
            Console.WriteLine("Creation: {0}", dir1.CreationTime);
            Console.WriteLine("Attributes: {0}", dir1.Attributes);
            Console.WriteLine("Root: {0}", dir1.Root);
            Console.WriteLine("");
            //Атрибуты нужного файла:
            DirectoryInfo dir2 = new DirectoryInfo(@"c:\88888888\2-Контакты");
            FileInfo[] fileUnderStudy = dir2.GetFiles("FetchFromCSharp1-5000.xml");
            Console.WriteLine("Файлов: {0}", fileUnderStudy.Length);
            foreach (FileInfo f in fileUnderStudy)
            {
                Console.WriteLine("File Name: {0}", f.Name);
                Console.WriteLine("File Size: {0}", f.Length);
                Console.WriteLine("Creation: {0}", f.CreationTime);
                Console.WriteLine("Attributes: {0}", f.Attributes);
            }
            Console.WriteLine("");
            Console.ReadLine();

            //Копировать файл:
            string filePathX = @"c:\88888888\2-Контакты\Файл для копирования.xlsx";
            string filePathY = @"c:\88888888\2-Контакты\Куда копировать.xlsx";
            File.Copy(filePathX, filePathY, true);  //(string SourceFileName, string DestinFileName, bool overwrite)

            //Работа с текстовым файлом: создание и запись:
            StreamWriter file1 = File.CreateText(@"\\inner.company\dfs\bases\1с-CRM\Нужный файл.xml");
            file1.WriteLine(testString);
            file1.Close();

            #endregion

            #region Работа с Excel.

            //Открытие Excel:
            Application xlApp = new ApplicationClass();
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            string filePath1 = @"c:\88888888\2-Контакты\ПРИМЕР для [Program00020 - На осн.00018 - ВСЕ НУЖНОЕ].xlsx";
            //Открытие файла (если такой существует) или создание (если не существует):
            FileInfo fInfo = new FileInfo(filePath1);
            if (!fInfo.Exists)
            {
                xlWorkBook = xlApp.Workbooks.Add(misValue); //Добавить новый Book в файл.
                Console.WriteLine("Файл создан.");
            }
            else //Открыть существующий файл.
            {
                xlWorkBook = xlApp.Workbooks.Open(filePath1, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Console.WriteLine("Файл открыт.");
            }
            //Открытие первого листа:
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "Возможные дубликаты в одном БП";
            //Запись значения в нужную по счету ячейку:
            xlWorkSheet.Cells[2, 4] = prompt;
            //Получить количество используемых столбцов:
            int countOfColumns = xlWorkSheet.UsedRange.Columns.Count;
            //Получить количество используемых строк:
            int countOfRows = xlWorkSheet.UsedRange.Rows.Count;
            //Проверить значение последней используемой ячейки в последней исползуеймой строке и 
            //последнем используемом столбце. Если значение этой ячейки равно "Привет", изменить его на "Пока".
            if ((xlWorkSheet.UsedRange.Cells[countOfRows, countOfColumns] as Range).Value2 != null)
            {
                if ((xlWorkSheet.UsedRange.Cells[countOfRows, countOfColumns] as Range).Value2.ToString() == "Привет")
                {
                    (xlWorkSheet.UsedRange.Cells[countOfRows, countOfColumns] as Range).Value2 = "Пока";
                }
            }

            int writeIntoRow = 1;
            //Как сделать гиперссылку (образец как было + мой пример):
            //Range linkParameter1asWasInTheExample = xlWorkSheet.get_Range("B6", Type.Missing);
            Range linkParameter1my = xlWorkSheet.get_Range("B" + writeIntoRow.ToString(), Type.Missing);
            //String linkParameter2asWasInTheExample = String.Empty;
            String linkParameter2my = "http://www.microsoft.com"; //"http://crm/Company/sfa/accts/edit.aspx?id=" + r.accountid + "#"; //Действующий html-адрес, по которому будет осуществляться переход.
            //String linkParameter2my1 = "http://crm/Company/sfa/accts/edit.aspx?id=" + r3.parentcustomerid + "#"; //Действующий html-адрес, по которому будет осуществляться переход.
            //String linkParameter2my2 = "http://crm/Company/sfa/conts/edit.aspx?id=" + r3.contactid + "#"; //Действующий html-адрес, по которому будет осуществляться переход.
            //String linkParameter2my3 = "http://crm/Company/cs/contractdetails/edit.aspx?id=" + r4.contractdetailid + "&_CreateFromType=1#"; //Действующий html-адрес, по которому будет осуществляться переход.
            //String linkParameter3asWasInTheExample = "Лист2!A1";
            String linkParameter3my = String.Empty;
            //String linkParameter4asWasInTheExample = "Screen Tip Text";
            String linkParameter4my = "Переход к этому БП в рабочей базе CRM.";
            //String linkParameter5asWasInTheExample = "Hyperlink Text";
            String linkParameter5my = "Текст, который будет выведен в ячейку";
            xlWorkSheet.Hyperlinks.Add(
                linkParameter1my,   //Область (набор ячеек), которую охватить ссылкой.
                linkParameter2my,   //Внешний адрес перехода (в интернете или во внутренней сети).
                linkParameter3my,   //Внутренний адрес перехода (ссылка на другое место этого Excel-файла).
                linkParameter4my,   //Текст всплывающей подсказки.
                linkParameter5my    //Текст в ячейке.
            );

            //Если файл существовал, просто сохранить его по умолчанию. Иначе сохранить в указанную директорию
            if (fInfo.Exists)
            {
                xlWorkBook.Save();
            }
            else
            {
                xlWorkBook.SaveAs(filePath1, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }

            //Объявление переменных:
            string[] accountsSetSource = new string[countOfRows - 1];   //Исходный массив - сюда поместим все из Excel.
            string[] accountsSetMiddle = new string[countOfRows - 1];   //Второй массив - сюда будем выбирать неповторяющиеся подряд из первого.
            string[] accountsSetDestin = new string[countOfRows - 1];   //Третий массив - сюда будем выбирать оригинальные из второго.
            string[] accountsSetSorted = new string[countOfRows - 1];   //Четвертый массив - сюда отсортировать третий массив.
            string[] accountsSetFm = new string[countOfRows - 1];
            string[] accountsSetIO = new string[countOfRows - 1];
            string[] accountsSetID = new string[countOfRows - 1];
            //Получение исходного массива, содержащего все БП из Excel:
            for (int iii = 2; iii <= countOfRows; iii++)
            {
                if ((xlWorkSheet.UsedRange.Cells[iii, 3] as Range).Value2 != null)  //проверка на всякий случай, что он есть.
                {
                    accountsSetSource[iii - 2] = (xlWorkSheet.UsedRange.Cells[iii, 3] as Range).Value2.ToString();
                    xlWorkSheet.Cells[iii, 4] = accountsSetSource[iii - 2];
                    Console.WriteLine(iii);
                }
            }
            accountsSetSorted = accountsSetSource;
            Array.Sort(accountsSetSorted, 0, accountsSetSorted.Length - 1); //Сортировать массив accountsSetSorted с 0-го по последний элемент.
            Console.WriteLine("Длина первого массива: {0}.", accountsSetSource.Length);
            
            //Проверка последней буквы у строкового значения:
            string str = "";
            str = (xlWorkSheet.UsedRange.Cells[1, 7] as Range).Value2.ToString();
            char ch = str[str.Length - 1];
            if (ch == 'Н')
            {
                str = 'Q' + str;
            }

            //Сделать первые буквы слов большими, остальные маленькими.
            str = FunctionsToStringWorking.getStringAsFIO(str);

            //Добавление пустой строки:
            xlWorkSheet1.get_Range("B6", "D6").EntireRow.Insert(1, null);
            //Добавление пустого столбца:
            xlWorkSheet1.get_Range("H8", "H9").EntireColumn.Insert(1, null);

            //Шрифт:
            (xlWorkSheet.Cells[3, 18] as Range).Font.Name = "Comic Sans MS"; //"Times New Roman";
            //Размер шрифта:
            (xlWorkSheet.Cells[3, 18] as Range).Font.Size = 16;
            //Жирность шрифта:
            (xlWorkSheet.Cells[3, 18] as Range).Font.Bold = true;
            //Стиль границы:
            (xlWorkSheet.Cells[3, 18] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            //Толщина границы:
            (xlWorkSheet.Cells[3, 18] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            //Выравнивание по горизонтали:
            (xlWorkSheet.Cells[3, 18] as Range).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xlWorkSheet.get_Range("M2", "Q3").EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignRight;  //Выравнивание: в этих столбцах - справа.
            xlWorkSheet.get_Range(xlWorkSheet.Cells[3, 13], xlWorkSheet.Cells[4, 17]).HorizontalAlignment = XlHAlign.xlHAlignLeft; //Выравнивание: в шапке - слева.
            //Выравнивание по вертикали:
            (xlWorkSheet.Cells[3, 18] as Range).VerticalAlignment = XlVAlign.xlVAlignCenter;
            //Объединение ячеек:
            xlWorkSheet.get_Range(xlWorkSheet.Cells[3, 14], xlWorkSheet.Cells[3, 17]).Merge(misValue);

            (xlWorkSheet.UsedRange.Cells[writeIntoRow, 14] as Range).NumberFormat = "@";    //Текстовый формат.

            xlWorkSheet.Columns.AutoFit();  //Ширина столбцов - автоматически по содержимому.
            xlWorkSheet.Rows.AutoFit();     //Высота строк - автоматически по содержимому.
            xlWorkSheet.get_Range("B1", "B3").EntireColumn.ColumnWidth = 50; //Ширина столбцов - задать для указанных столбцов, определяемых указанной областью.

            //Отсортировать строки в файле Excel: сначала по значению одного столбца, затем другого столбца, затем третьего столбца:
            //Сначала берем из листа ячеек прямоугольник значений: в данном случае с ячейки B2 до ячейки H-сколько-то (номер последней нужной строки);...
            //...затем сортируем по столбцам, задавая номера столбцов начиная с первого столбца выбранного прямоугольника значений. ...
            //...В данном случае, т. к. выбран прямоугольник начиная с ячейки B2, то столбец 1 - это столбец B, столбец 2 - это столбец C, и т. д. ...
            //...Далее, сначала сортируем по столбцу 3;...
            //...затем в рамках одинаковых значений в столбце 3 сортируем по столбцу 2;...
            //...затем в рамках одинаковых значений в столбце 2 сортируем по столбцу 1.
            Range sortingRange = xlWorkSheet.get_Range("B2", "H" + (writeIntoRow - 1).ToString());
            sortingRange.Sort(
                sortingRange.Columns[3, Type.Missing],
                XlSortOrder.xlAscending,
                sortingRange.Columns[2, Type.Missing],
                Type.Missing,
                XlSortOrder.xlAscending,
                sortingRange.Columns[1, Type.Missing],
                XlSortOrder.xlAscending,
                XlYesNoGuess.xlGuess,
                Type.Missing,
                Type.Missing,
                XlSortOrientation.xlSortColumns,
                XlSortMethod.xlPinYin,
                XlSortDataOption.xlSortNormal,
                XlSortDataOption.xlSortNormal,
                XlSortDataOption.xlSortNormal
            );
            var sortedRange = (object[,])sortingRange.Value2;
            sortingRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, sortedRange);

            xlApp.Visible = true;   //Отобразить полученный файл.
            xlWorkBook.Activate();  //Передать фокус в отображенный файл. Предупреждение вроде не мешает работе программы.

            //Закрытие Excel:
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            //Освобождение ресурсов:
            FunctionsToAnyDataWorking.releaseObject(xlWorkSheet);
            FunctionsToAnyDataWorking.releaseObject(xlWorkBook);
            FunctionsToAnyDataWorking.releaseObject(xlApp);
            
            Console.WriteLine("Файл закрыт и объект удален.");
            Console.WriteLine("");
            Console.ReadLine();

            #endregion

            #region Работа с Firebird.

            //Для доступа к Firebird:
            //using FirebirdSql.Data.FirebirdClient;          //Добавляем пространство имен firebird ado.net provider, необходимое для взаимодействия с Firebird.

            //Переменные, нужные в контексте всей программы:
            FbConnection fb;    //Переменная fb ссылается на соединение с базой данных, поэтому она должна быть доступна всем методам нашего класса.

            //Формируем connection string для последующего соединения с базой данных (не сработало):
            //FbConnectionStringBuilder fb_con = new FbConnectionStringBuilder();
            //fb_con.Charset = "WIN1251"; //используемая кодировка
            //fb_con.UserID = tbLogin.Text; //логин
            //fb_con.Password = tbPwd.Text; //пароль
            //fb_con.Database = tbPathToDB.Text; //путь к файлу базы данных
            //fb_con.ServerType = 0; //указываем тип сервера (0 - "полноценный Firebird" (classic или super server), 1 - встроенный (embedded))

            //Создаем подключение, передаем строку подключения объекту класса FbConnection:
            //fb = new FbConnection(fb_con.ToString());
            fb = new FbConnection(@"Database=srv-arm:d:\company-service\armdl\mdelta.gdb;UserID=sysdba;Password=CrazyRabbit);Charset=win1251;ServerType=0");
            fb.Open();
            FbDatabaseInfo fb_inf = new FbDatabaseInfo(fb); //Получаем информацию о БД.
            MessageBox.Show("Info: " + fb_inf.ServerClass + "; " + fb_inf.ServerVersion);   //Выводим тип и версию используемого сервера Firebird.

            if (fb.State == ConnectionState.Closed) fb.Open();  //Проверить состояние соединения, если закрыто - открыть.
            FbTransaction fbt = fb.BeginTransaction();  //Стартуем транзакцию (можно только для открытой базы).
            //FbCommand SelectSQL = new FbCommand("SELECT * FROM HIS_USER_COMPLECTS WHERE UC_ID='56990'", fb);  //Задаем запрос на выборку.
            FbCommand SelectSQL = new FbCommand("SELECT P_UC_ID, UC_ID, UC_NAME, DATA_TO FROM HIS_USER_COMPLECTS WHERE UC_ID='56990'", fb);  //Задаем запрос на выборку.
            SelectSQL.Transaction = fbt;    //Необходимо проинициализировать транзакцию для объекта SelectSQL.
            FbDataReader reader = SelectSQL.ExecuteReader();    //Для запросов, которые возвращают результат в виде набора данных, используется этот метод.
            string select_result = "";  //В эту переменную будем складывать результат запроса Select.
            try
            {
                while (reader.Read())   //Выполняем, пока не прочли все данные.
                {
                    try
                    {
                        immediateObject = reader.GetDateTime(3);    //Если сработало, значит reader.GetDateTime(3) не null.
                        immediateString = reader.GetDateTime(3).ToString();
                    }
                    catch
                    {
                        immediateString = "NULL";
                        //Сохранить найденные Имя и Номер заказа:
                        foundName = reader.GetString(2);
                        foundOrderNumber = reader.GetInt32(0).ToString();
                    }
                    select_result = select_result +
                        reader.GetInt32(0).ToString() + ", " +
                        reader.GetInt32(1).ToString() + ", " +
                        reader.GetString(2) + ", " +
                        immediateString + ", " +
                        "\n";
                    /*
                    select_result = select_result +
                        reader.GetInt32(0).ToString() + ", " +
                        reader.GetInt32(1).ToString() + ", " +
                        //reader.GetInt32(2).ToString() + ", " + 
                        //reader.GetInt32(3).ToString() + ", " + 
                        //reader.GetInt32(4).ToString() + ", " + 
                        reader.GetString(5) + ", " + //"\n" + 
                        //reader.GetInt32(6).ToString() + ", " + 
                        //reader.GetInt32(7).ToString() + ", " + 
                        //reader.GetInt32(8).ToString() + ", " + 
                        //reader.GetDateTime(9).ToString() + ", " + 
                        //reader.GetDateTime(10).ToString() + ", " + 
                        //reader.GetDateTime(11).ToString() + ", " + 
                        s + ", " +
                        //reader.GetInt32(12).ToString() + ", " + 
                        ////reader.GetInt32(13).ToString() + ", " + //Где с 4-мя слэшами - там закомментарено изначально, т. к. содержит null.
                        ////reader.GetDateTime(14).ToString() + ", " + 
                        //reader.GetInt32(15).ToString() + ", " + 
                        //reader.GetInt32(16).ToString() + ", " + 
                        ////reader.GetInt32(17).ToString() + ", " + 
                        //reader.GetInt32(18).ToString() + ", " + 
                        //reader.GetInt32(19).ToString() + ", " + 
                        //reader.GetInt32(20).ToString() + ", " + 
                        //reader.GetInt32(21).ToString() + ", " + 
                        ////reader.GetDateTime(22).ToString() + ", " + 
                        ////reader.GetInt32(23).ToString() + ", " + 
                        //reader.GetInt32(24).ToString() + ", " + 
                        ////reader.GetDateTime(25).ToString() + ", " + 
                        "\n";
                    */
                }
            }
            finally
            {
                //Всегда необходимо вызывать метод Close(), когда чтение данных завершено:
                reader.Close();
                fb.Close(); //Закрываем соединение, т. к. оно нам больше не нужно.
            }
            SelectSQL.Dispose();    //В документации написано, что очень рекомендуется удалять объекты этого типа, если они больше не нужны.

            //MessageBox.Show(select_result); //Выводим результат запроса.
            //MessageBox.Show("Найдено:\n" + foundName + "\n" + foundOrderNumber); //Выводим результат запроса.

            xlWorkSheet1.Cells[iii, 8] = foundName;
            xlWorkSheet1.Cells[iii, 22] = foundOrderNumber;

            fb.Close();

            #endregion

            #region Работа с визуальной частью (с элементами формы).

            progressBar1.Minimum = 0;
            progressBar1.Maximum = countOfAll;
            progressBar1.Value = countOfProcessed;
            label11.Location = new System.Drawing.Point(((this.Size.Width - label11.Size.Width) / 2) - 6, 265);
            label11.BringToFront(); //Поместить контрол поверх прочих контролов, чтобы его было видно.

            this.Cursor = Cursors.WaitCursor;   //Ожидающий курсор.
            this.Cursor = Cursors.Default;      //Курсор по умолчанию.

            #endregion

            #region Получение Xml-файла. Получение результатов фетч-запросов.

            Console.WriteLine("Wait! you'll get a fetch result.");

            //Код по примеру из справки CRM SDK.
            //Создание сервиса здесь приведено для примера, а для использования оно реализовано функцией createCrmService.
            CrmAuthenticationToken token = new CrmAuthenticationToken();
            token.AuthenticationType = 0;
            token.OrganizationName = "CmpnyLab";   //Тестовая - "CmpnyLab", рабочая - "Cmpny".
            CrmService service1 = new CrmService();
            service1.Url = "http://crm4/mscrmservices/2007/crmservice.asmx";
            service1.CrmAuthenticationTokenValue = token;
            service1.Credentials = System.Net.CredentialCache.DefaultCredentials;

            //Как получить результат запроса по сущностям в отношении с данной сущностью:
            //По сущностям в отношении 1:N, N:1 к данной - реализуется через фетч с использованием <link-entity...>.
            //По сущностям в отношении N:N к данной - реализуется через фетч с использованием специальной промежуточной сущности (по образцу listmember).

            string fetch1 = @"
                <fetch mapping=""logical"">
                    <entity name=""account"">
                        <attribute name=""name""/>
                    </entity>
                </fetch>";
                    //<entity name=""account"">
                    //  <all-attributes/>
                    //</entity>
            string fetch2 = @"
                <fetch mapping=""logical"">
                    <entity name=""contact"">
                        <attribute name=""fullname""/>
                        <attribute name=""firstname""/>
                        <attribute name=""lastname""/>
                        <attribute name=""parentcustomerid""/>
                        <attribute name=""gar_function""/>
                        <attribute name=""jobtitle""/>
                        <order attribute=""lastname"" descending=""false""/>
                        <filter type='and'>
                            <condition attribute = 'fullname' operator='eq' value='Корненко, Ксения Юрьевна'/>                                                       
                        </filter>
                    </entity>
                </fetch>";
            string fetch3 = @"
                <fetch mapping=""logical"">
                    <entity name=""account"">
                        <attribute name=""name""/>
                        <link-entity name=""systemuser"" to=""owninguser"">
                            <filter type=""and"">
                                <condition attribute=""lastname"" operator=""ne"" value=""Cannon""/>
                            </filter>
                        </link-entity>
                    </entity>
                </fetch>";
            string result001 = service1.Fetch(fetch1);                 //Стандартный запрос - выдает 5000 значений максимум.
            string result002 = service1.Fetch(fetch2);                 //Стандартный запрос - выдает 5000 значений максимум.
            string result003 = service1.Fetch(fetch3);                 //Стандартный запрос - выдает 5000 значений максимум.
            string result1 = service1.Fetch(fetch1);                 //Стандартный запрос - выдает 5000 значений максимум.
            string result2 = FunctionsToCrmDataWorking.fetchAll(fetch1, service1);   //Хитрый запрос - выдает все значения.
            //Конец кода по примеру из справки CRM SDK.

            //Запрос с динамически формируемым условием отбора (по заранее неизвестному количеству значений):
            //Получить список list (МСписков), выбрав все удовлетворяющие критериям МС по id из списка listmember'ов для этого БП:
            //Формируем fetch-запрос: сначала начальную часть строки, потом - в цикле - ряд условий, потом конечную часть строки.
            string fetch4 = @"
                <fetch mapping=""logical"">
                    <entity name=""list"">
                        <attribute name=""listid""/>
                        <attribute name=""listname""/>
                        <attribute name=""statuscode""/>
                        <attribute name=""createdon""/>
                        <attribute name=""gar_project""/>
                        <filter type='and'>
                            <condition attribute = 'statuscode' operator='eq' value='0'/>
                            <filter type='or'>";
            foreach (CrmListmember r in resList3)
            {
                fetch4 = fetch4 + @"<condition attribute = 'listid' operator='eq' value='" + r.listid + @"'/>";
            }
            fetch4 = fetch4 + @"
                            </filter>
                        </filter>
                        <link-entity name=""gar_project"" to=""gar_project"">
                            <filter type=""and"">
                                <condition attribute=""gar_name"" operator=""like"" value=""%Гар%""/>
                            </filter>
                        </link-entity>
                    </entity>
                </fetch>";
            string result4 = service.Fetch(fetch4);
            XmlDocument xmlDoc4 = new XmlDocument();
            xmlDoc4.LoadXml(result4);
            List<CrmList> resList4 = getParsedItemsList(xmlDoc4);
            //Конец Получить список list (МСписков), выбрав все удовлетворяющие критериям МС по id из списка listmember'ов для этого БП.
            
            //Запись в Xml-файл результатов запросов:
            string[] result1array = { result1 };
            File.WriteAllLines(@"c:\88888888\2-Контакты\Example1-5000.xml", result1array);
            string[] result2array = { result2 };
            File.WriteAllLines(@"c:\88888888\2-Контакты\Example2-all.xml", result2array);

            Console.WriteLine(result1);
            Console.WriteLine("It was a FETCH!!!");
            Console.ReadLine();

            #endregion

            #region Работа с Xml-файлом.

            //Получение разобранного списка не из файла, а из документа:
            string result3 = FunctionsToCrmDataWorking.fetchAll(fetch3, service1);
            //Если надо получить результат запроса в файл, чтобы посмотреть визуально:
            //string[] result3array = { result3 };
            //File.WriteAllLines(dir1.FullName + @"\Accounts.xml", result3array);
            XmlDocument xmlDoc3 = new XmlDocument();
            xmlDoc3.LoadXml(result3);
            List<CrmAccount> resList3 = getParsedItemsAccount(xmlDoc3);
            
            string filePath2 = @"c:\88888888\2-Контакты\FetchFromCSharp1-5000.xml";
            List<CrmContact> resList = getParsedItemsContact(filePath2);
            int resNumber = 0;
            foreach (CrmContact r in resList)
            {
                resNumber++;
                Console.WriteLine("");
                Console.WriteLine("{0}: r.fullname = {1}",              resNumber, r.fullname);
                Console.WriteLine("{0}: r.firstname = {1}",             resNumber, r.firstname);
                Console.WriteLine("{0}: r.lastname = {1}",              resNumber, r.lastname);
                Console.WriteLine("{0}: r.parentcustomerid = {1}",      resNumber, r.parentcustomerid);
                Console.WriteLine("{0}: r.parentcustomeridname = {1}",  resNumber, r.parentcustomeridname);
                Console.WriteLine("{0}: r.gar_function = {1}",          resNumber, r.gar_function);
                Console.WriteLine("{0}: r.gar_functionname = {1}",      resNumber, r.gar_functionname);
                Console.WriteLine("{0}: r.contactid = {1}",             resNumber, r.contactid);
                Console.WriteLine("{0}: r.statuscode = {1}",            resNumber, r.statuscode);
                Console.WriteLine("{0}: r.statecode = {1}",             resNumber, r.statecode);
                Console.WriteLine("{0}: r.ownerid = {1}",               resNumber, r.ownerid);
                Console.WriteLine("{0}: r.owneridname = {1}",           resNumber, r.owneridname);
            }

            #endregion

            #region Работа с CRM-сервисом - 1 - через статические сущности, НЕ умеет изменять пользовательские пиклисты.

            //Создаем объект CrmService:
            //У нас - имя организации: тестовая база - "CmpnyLab", рабочая база - "Cmpny".
            CrmService service2 = FunctionsToCrmDataWorking.createCrmService("CmpnyLab", "http://crm4/mscrmservices/2007/crmservice.asmx");
            //Работа с Аккаунтом (Бизнес-партнером):
            //for (int iii = 6; iii <= 6; iii++)
            {
                //int number = iii;
                account crAccount = new account();
                crAccount.name = "Тестовый аккаунт для записи в тестовую базу4";
                Guid guidOfNewAccount = service2.Create(crAccount);
                Console.WriteLine("ИД вновь созданного аккаунта (для вывода в Excel-файл): {0}", guidOfNewAccount.ToString());

                contact upContact = new contact();
                upContact.firstname = "Имя-отчество";
                upContact.lastname = "Фамилия";
                upContact.contactid = new Key();
                upContact.contactid.Value = new Guid("{EABBDB13-8AA9-DF11-A3FA-00155D00282B}");
                upContact.familystatuscode = new Picklist();
                upContact.familystatuscode.Value = -1;
                upContact.gendercode = new Picklist(-1);
                service2.Update(upContact);

                annotation note = new annotation();
                note.notetext = "Test string for annotation's notetext.";
                note.subject = "Test Subject";
                note.objectid = new Lookup();
                note.objectid.type = EntityName.incident.ToString();
                note.objectid.Value = new Guid("{4D8B427E-CC0E-E211-B23A-00155D001308}");
                note.objecttypecode = new EntityNameReference();
                note.objecttypecode.Value = EntityName.incident.ToString();
                Guid createdNoteId = service1.Create(note);

                //Создаем Контракт статически (работает):
                contract newContract = new contract();
                newContract.title = "56789-3";
                newContract.contracttemplateid = new Lookup();
                newContract.contracttemplateid.Value = new Guid("{A8CB135E-ACA3-DF11-9FB6-00155D00282B}");
                newContract.billingcustomerid = new Customer();
                newContract.billingcustomerid.type = EntityName.account.ToString();
                newContract.billingcustomerid.Value = new Guid("{FB21E8FB-9FE2-DE11-9387-00155D4E1B14}");
                newContract.customerid = new Customer();
                newContract.customerid.type = EntityName.account.ToString();
                newContract.customerid.Value = new Guid("{FB21E8FB-9FE2-DE11-9387-00155D4E1B14}");
                newContract.activeon = new CrmDateTime(DateTime.Now.ToString(string.Format("yyyy-MM-ddTHH:mm:ss")));
                newContract.expireson = new CrmDateTime(DateTime.Now.ToString(string.Format("yyyy-MM-ddTHH:mm:ss")));
                Guid guidOfNewContract = crmService.Create(newContract);

                //Получить имеющийся аккаунт:
                Guid newGuid = new Guid("{8694E2A3-9EE2-DE11-9387-00155D4E1B14}");
                ColumnSetBase colSet = new AllColumns();
                account getAccount = (account)service1.Retrieve("account", newGuid, colSet);

                //Получить Отправителя и Получателя из Звонка:
                XmlDocument xmlDoc1 = new XmlDocument();
                List<CrmPhonecall> resList1 = getParsedItemsPhonecall(xmlDoc1);
                foreach (CrmPhonecall r in resList1)
                {
                    Guid phonecallGuid = new Guid(r.activityid);
                    ColumnSetBase colSet1 = new ColumnSet();
                    colSet1.AddColumns("from"); //Отправитель.
                    colSet1.AddColumns("to");   //Получатель.
                    phonecall getPhonecall = (phonecall)service2.Retrieve("phonecall", phonecallGuid, colSet1);

                    string nam1 = "";
                    Lookup lkp1 = null;
                    string val1 = "";
                    try
                    {
                        lkp1 = getPhonecall.from[0].partyid;
                        string typ1 = lkp1.type.ToString();     // = "systemuser".
                        nam1 = lkp1.name.ToString();            // = Имя.
                        val1 = lkp1.Value.ToString();           // = ID.
                    }
                    catch { }

                    string nam2 = "";
                    Lookup lkp2 = null;
                    string val2 = "";
                    try
                    {
                        lkp2 = getPhonecall.to[0].partyid;
                        string typ2 = lkp2.type.ToString();     // = "account".
                        nam2 = lkp2.name.ToString();            // = Название (если БП).
                        val2 = lkp2.Value.ToString();    // = ID.
                    }
                    catch { }
                }
                //Конец получения Отправителя и Получателя из Звонка.

                //Перевод System Job в статус "Canceled" с последующим удалением:
                asyncoperation delAsyncoperation = new asyncoperation();
                delAsyncoperation.asyncoperationid = new Key();
                delAsyncoperation.asyncoperationid.Value = new Guid(r.asyncoperationid);
                delAsyncoperation.statecode = new AsyncOperationStateInfo();
                delAsyncoperation.statecode.Value = AsyncOperationState.Completed;
                delAsyncoperation.statuscode = new Status();
                delAsyncoperation.statuscode.Value = 32;
                service.Update(delAsyncoperation);

                Guid guidAsyncoperation = new Guid(r.asyncoperationid);
                service.Delete("asyncoperation", guidAsyncoperation);
                //Конец перевода System Job в статус "Canceled" с последующим удалением.

                //Удалить имеющийся Контакт:
                Guid contactToDelete = new Guid(r.contactid);
                service.Delete("contact", contactToDelete);
            }

            #region Изменение Состояния и Статуса.

            //Изменение состояния и статуса Действия сервиса:
            SetStateServiceAppointmentRequest req = new SetStateServiceAppointmentRequest();
            req.EntityId = new Guid(r.activityid);
            //Сделать доступным перед изменением: Статус - Запланировано, Состояние - 4:
            req.ServiceAppointmentState = ServiceAppointmentState.Scheduled;
            req.ServiceAppointmentStatus = 4;
            SetStateServiceAppointmentResponse res1 = (SetStateServiceAppointmentResponse)service.Execute(req);

            //Внести требуемые изменения...

            //Сделать недоступным после изменения: Статус - Закрыто, Состояние - 8:
            req.ServiceAppointmentState = ServiceAppointmentState.Closed;
            req.ServiceAppointmentStatus = 8;
            SetStateServiceAppointmentResponse res2 = (SetStateServiceAppointmentResponse)service.Execute(req);

            //Аналогично - Изменение Состояния и Статуса Бизнес-партнера:
            SetStateAccountRequest state = new SetStateAccountRequest();
            state.EntityId = new Guid("AD618DB2-F0DB-4A6A-8C4B-2F2213EAA38E");
            state.AccountState = AccountState.Inactive;
            state.AccountStatus = 2;
            SetStateAccountResponse stateSet = (SetStateAccountResponse)service.Execute(state);

            //Аналогично - Изменение Состояния и Статуса Контакта:
            SetStateContactRequest contactState = new SetStateContactRequest();
            contactState.EntityId = new Guid(r.contactid);
            contactState.ContactState = ContactState.Inactive;
            contactState.ContactStatus = 2;
            SetStateContactResponse contactStateSet = (SetStateContactResponse)service.Execute(contactState);

            //Аналогично - Изменение Состояния и Статуса Задачи:
            SetStateTaskRequest tr = new SetStateTaskRequest();
            tr.EntityId = new Guid("{ABA1DD03-DD17-E411-AF6B-00155D001308}");
            tr.TaskState = TaskState.Open;
            tr.TaskStatus = 2;  //"Не начато".
            SetStateTaskResponse res = (SetStateTaskResponse)service.Execute(tr);

            #endregion

            #region Изменение имеющегося Запроса (представления) с публикацией.

            CrmService service = FunctionsToCrmDataWorking.createCrmService("CmpnyLab", "http://crm4/mscrmservices/2007/crmservice.asmx");

            //Получить имеющийся Запрос (представление) "Активные контакты (служебное представление 01)" - один раз, чтобы получить фетч-строку:
            //Guid newGuid = new Guid("{6FF24503-5D6F-E211-91EA-00155D001308}");
            //ColumnSetBase colSet = new AllColumns();
            //savedquery getSQ = (savedquery)service.Retrieve("savedquery", newGuid, colSet);
            //Console.WriteLine(getSQ.fetchxml);
            //И после открытия Excel: xlWorkSheet.Cells[2, 5] = getSQ.fetchxml;

            string fetchString = @"
                <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                    <entity name=""contact"">
                        <attribute name=""fullname"" />
                        <attribute name=""parentcustomerid"" />
                        <attribute name=""telephone1"" />
                        <attribute name=""emailaddress1"" />
                        <attribute name=""gar_function"" />
                        <attribute name=""gar_gar_addresses_contact1"" />
                        <attribute name=""customertypecode"" />
                        <attribute name=""gar_relevance_contact"" />
                        <attribute name=""gar_accountrolecode_3"" />
                        <attribute name=""gar_accountrolecode_2"" />
                        <attribute name=""accountrolecode"" />
                        <order attribute=""fullname"" descending=""false"" />
                        <filter type=""and"">
                            <condition attribute=""parentcustomerid"" operator=""eq"" 
                                uiname=""((СЛУЖЕБНОЕ ПРЕДСТАВЛЕНИЕ, ПРОСЬБА НЕ ИЗМЕНЯТЬ!))"" uitype=""account"" value=""{AC70CBAD-9FE2-DE11-9387-00155D4E1B14}"" />
                        </filter>
                        <attribute name=""contactid"" />
                    </entity>
                </fetch>";

            //Изменить Сохраненный запрос (т. е. представление):
            savedquery setSQ = new savedquery();
            setSQ.fetchxml = fetchString;
            setSQ.savedqueryid = new Key();
            setSQ.savedqueryid.Value = new Guid("{6FF24503-5D6F-E211-91EA-00155D001308}");
            service.Update(setSQ);

            //Опубликовать изменения в сущности Контакт (поскольку Сохраненный запрос относится к Контакту):
            PublishXmlRequest request = new PublishXmlRequest();
            request.ParameterXml = @"
                <importexportxml>
                    <entities>
                        <entity>contact</entity>
                    </entities>
                    <nodes/>
                    <securityroles/>
                    <settings/>
                    <workflows/>
                </importexportxml>";
            PublishXmlResponse response = (PublishXmlResponse)service.Execute(request);

            #endregion

            #endregion

            #region Работа с CRM-сервисом - 2 - через динамические сущности (DynamicEntity), умеет изменять пользовательские пиклисты.

            bool success = false;
            try
            {
                CrmService service4 = FunctionsToCrmDataWorking.createCrmService("CmpnyLab", "http://crm4/mscrmservices/2007/crmservice.asmx");

                //Создает account "Fourth Coffee" и возвращает его id.
                #region Setup Data Required for this Sample
                // Create the account object.
                account account = new account();
                account.name = "Fourth Coffee";
                // Create the target object for the request.
                TargetCreateAccount target = new TargetCreateAccount();
                target.Account = account;
                // Create the request object.
                CreateRequest createRequest = new CreateRequest();
                createRequest.Target = target;
                // Execute the request.
                CreateResponse createResponse = (CreateResponse)service4.Execute(createRequest);
                Guid accountID = createResponse.id;
                #endregion

                //Создает contact "Jesper Aaberg" и возвращает его id.
                #region Create Contact Dynamically
                // Set the properties of the contact using property objects.
                StringProperty firstname = new StringProperty();
                firstname.Name = "firstname";
                firstname.Value = "Jesper";
                StringProperty lastname = new StringProperty();
                lastname.Name = "lastname";
                lastname.Value = "Aaberg";
                // Create the DynamicEntity object.
                DynamicEntity contactEntity = new DynamicEntity();
                // Set the name of the entity type.
                contactEntity.Name = EntityName.contact.ToString();
                // Set the properties of the contact.
                contactEntity.Properties.Add(firstname);
                contactEntity.Properties.Add(lastname);
                //contactEntity.Properties = new Property[] { firstname, lastname };
                // Create the target.
                TargetCreateDynamic targetCreate = new TargetCreateDynamic();
                targetCreate.Entity = contactEntity;
                // Create the request object.
                CreateRequest create = new CreateRequest();
                // Set the properties of the request object.
                create.Target = targetCreate;
                // Execute the request.
                CreateResponse created = (CreateResponse)service4.Execute(create);
                Guid contactID = created.id;
                #endregion

                //Создаем Контракт динамически (работает):
                #region Create Contract Dynamically
                DynamicEntity entityToCreate = new DynamicEntity();
                entityToCreate.Name = "contract";
                entityToCreate.Properties.Add(new StringProperty("title", "КОПИЯ"));
                entityToCreate.Properties.Add(new LookupProperty("contracttemplateid", new Lookup("contracttemplate", new Guid("{A8CB135E-ACA3-DF11-9FB6-00155D00282B}"))));
                entityToCreate.Properties.Add(new CustomerProperty("billingcustomerid", new Customer("account", new Guid("{FB21E8FB-9FE2-DE11-9387-00155D4E1B14}"))));
                entityToCreate.Properties.Add(new CustomerProperty("customerid", new Customer("account", new Guid("{FB21E8FB-9FE2-DE11-9387-00155D4E1B14}"))));
                entityToCreate.Properties.Add(new CrmDateTimeProperty("activeon", new CrmDateTime(DateTime.Now.ToString(string.Format("yyyy-MM-ddTHH:mm:ss")))));
                entityToCreate.Properties.Add(new CrmDateTimeProperty("expireson", new CrmDateTime(DateTime.Now.ToString(string.Format("yyyy-MM-ddTHH:mm:ss")))));
                TargetCreateDynamic orderTCD = new TargetCreateDynamic();
                orderTCD.Entity = entityToCreate;
                CreateRequest orderCR = new CreateRequest();
                orderCR.Target = orderTCD;
                ICrmService service = context.CreateCrmService(false);
                CreateResponse response = (CreateResponse)service.Execute(orderCR);
                #endregion

                //Возвращает contact по заданному id (задаем id контакта "Jesper Aaberg").
                #region Retrieve Contact Dynamically
                // Create the retrieve target.
                TargetRetrieveDynamic targetRetrieve = new TargetRetrieveDynamic();
                // Set the properties of the target.
                targetRetrieve.EntityName = EntityName.contact.ToString();
                targetRetrieve.EntityId = created.id;
                // Create the request object.
                RetrieveRequest retrieve = new RetrieveRequest();
                // Set the properties of the request object.
                retrieve.Target = targetRetrieve;
                // Be aware that using AllColumns may adversely affect
                // performance and cause unwanted cascading in subsequent 
                // updates. A best practice is to retrieve the least amount of 
                // data required.
                retrieve.ColumnSet = new AllColumns();
                // Indicate that the BusinessEntity should be retrieved as a DynamicEntity.
                retrieve.ReturnDynamicEntities = true;
                // Execute the request.
                RetrieveResponse retrieved = (RetrieveResponse)service4.Execute(retrieve);
                // Extract the DynamicEntity from the request.
                DynamicEntity entity = (DynamicEntity)retrieved.BusinessEntity;
                // Extract the fullname from the dynamic entity
                string fullname = "";
                foreach (Property pr in entity.Properties)
                {
                    if (pr.Name.ToLower() == "fullname")
                    {
                        StringProperty property = (StringProperty)pr;
                        fullname = fullname + property.Value;
                        break;
                    }
                }
                #endregion

                //Возвращает Обращение по заданному id.
                #region Retrieve Incident Dynamically
                TargetRetrieveDynamic targetRetrieve2 = new TargetRetrieveDynamic();
                targetRetrieve2.EntityName = EntityName.incident.ToString();
                targetRetrieve2.EntityId = new Guid("Write a GUID here.");
                RetrieveRequest retrieve2 = new RetrieveRequest();
                retrieve2.Target = targetRetrieve2;
                // Be aware that using AllColumns may adversely affect performance and cause unwanted cascading
                // in subsequent updates. A best practice is to retrieve the least amount of data required.
                retrieve2.ColumnSet = new ColumnSet();   // = new AllColumns();
                retrieve2.ColumnSet.AddColumn("ticketnumber");
                retrieve2.ColumnSet.AddColumn("gar_source_incident_1");
                retrieve2.ColumnSet.AddColumn("gar_source_incident_2");
                retrieve2.ColumnSet.AddColumn("gar_source_incident_3");
                retrieve2.ColumnSet.AddColumn("gar_comments_source_1");
                retrieve2.ColumnSet.AddColumn("gar_comments_source_2");
                retrieve2.ColumnSet.AddColumn("gar_comments_source_3");
                retrieve2.ReturnDynamicEntities = true;
                RetrieveResponse retrieved2 = new RetrieveResponse();
                try
                {
                    retrieved2 = (RetrieveResponse)service1.Execute(retrieve2);
                }
                catch
                {
                    Console.WriteLine("Не удалось получить Обращение. Наверное, оно удалено.");
                    //continue;   //Перейти к следующей итерации ближайшего foreach.
                }
                DynamicEntity entity2 = (DynamicEntity)retrieved2.BusinessEntity;
                //Выбираем Источники и Комментарии:
                string source1 = "";
                string komment1 = "";
                foreach (Property pr in entity2.Properties)
                {
                    if (pr.Name.ToLower() == "gar_source_incident_1")
                    {
                        LookupProperty property = (LookupProperty)pr;
                        source1 = property.Value.name;
                    }
                    if (pr.Name.ToLower() == "gar_comments_source_1")
                    {
                        StringProperty property = (StringProperty)pr;
                        komment1 = property.Value;
                    }
                }
                #endregion

                //Изменяет возвращенный Контакт: добавляет значения свойств типов money, picklist, customer.
                #region Update the DynamicEntity
                // This part of the example demonstrates how to update properties of a DynamicEntity.
                // Set the contact properties dynamically.
                // Contact Credit Limit
                CrmMoneyProperty money = new CrmMoneyProperty();
                // Specify the property name of the DynamicEntity.
                money.Name = "creditlimit";
                money.Value = new CrmMoney();
                // Specify a $10000 credit limit.
                money.Value.Value = 10000M;
                // Contact PreferredContactMethodCode property
                PicklistProperty picklist = new PicklistProperty();
                //   Specify the property name of the DynamicEntity. 
                picklist.Name = "preferredcontactmethodcode";
                picklist.Value = new Picklist();
                //   Set the property's picklist index to 1.
                picklist.Value.Value = 5;
                // Contact ParentCustomerId property.
                CustomerProperty parentCustomer = new CustomerProperty();
                //   Specify the property name of the DynamicEntity.
                parentCustomer.Name = "parentcustomerid";
                parentCustomer.Value = new Customer();
                //   Set the customer type to account.
                parentCustomer.Value.type = EntityName.account.ToString();
                //   Specify the GUID of an existing CRM account.
                // SDK:parentCustomer.Value.Value = new Guid("A0F2D8FE-6468-DA11-B748-000D9DD8CDAC");
                parentCustomer.Value.Value = accountID;
                //   Update the DynamicEntities properties collection to add new properties.
                //   Add properties to ArrayList.
                entity.Properties.Add(money);
                entity.Properties.Add(picklist);
                entity.Properties.Add(parentCustomer);
                // Create the update target.
                TargetUpdateDynamic updateDynamic = new TargetUpdateDynamic();
                // Set the properties of the target.
                updateDynamic.Entity = entity;
                //   Create the update request object.
                UpdateRequest update = new UpdateRequest();
                //   Set request properties.
                update.Target = updateDynamic;
                //   Execute the request.
                UpdateResponse updated = (UpdateResponse)service4.Execute(update);
                #endregion

                //Назначает динамическую сущность пользователю.
                #region Assign the DynamicEntity
                AssignRequest assignEntity = new AssignRequest();
                assignEntity.Assignee = new SecurityPrincipal();
                assignEntity.Assignee.Type = SecurityPrincipalType.User;
                assignEntity.Assignee.PrincipalId = new Guid("289D02C4-FC19-E311-8C4C-00155D001308");
                TargetOwnedDynamic target = new TargetOwnedDynamic();
                target.EntityId = newGar_kitID;
                target.EntityName = "gar_kit";
                assignEntity.Target = target;
                service.Execute(assignEntity);
                #endregion

                //Изменяет возвращенное Обращение динамически (хитро, через изменение статуса Обращения и закрытие Обращения специальными методами).
                #region Update the Incident

                LookupProperty gar_source_incident_1 = new LookupProperty();
                gar_source_incident_1.Name = "gar_source_incident_1";
                gar_source_incident_1.Value = new Lookup();
                gar_source_incident_1.Value.type = "gar_source";
                gar_source_incident_1.Value.Value = new Guid("Write a GUID here.");
                entity2.Properties.Add(gar_source_incident_1);
                StringProperty gar_comments_source_1 = new StringProperty();
                gar_comments_source_1.Name = "gar_comments_source_1";
                gar_comments_source_1.Value = "Write some string here.";
                entity2.Properties.Add(gar_comments_source_1);

                //Получить имеющееся Обращение:
                Guid guidIncident1 = new Guid("Write a GUID here.");
                ColumnSetBase colSet = new ColumnSet();
                colSet.AddColumn("statecode");
                colSet.AddColumn("statuscode");
                incident getIncident = (incident)service1.Retrieve("incident", guidIncident1, colSet);

                //Сохранить из него Статус и Состояние:
                IncidentStateInfo state = getIncident.statecode;
                Status status = getIncident.statuscode;
                //Создать Разрешение обращения для сохранения туда Разрешения, если будет надо:
                incidentresolution incidentResolution = new incidentresolution();

                //Сделать Статус и Состояние разрешенными перед динамическим изменением:
                if (status.Value != 1)  //Если Состояние еще не является разрешенным.
                {
                    SetStateIncidentRequest ssiRequest1 = new SetStateIncidentRequest();
                    ssiRequest1.EntityId = guidIncident1;
                    ssiRequest1.IncidentState = 0;
                    ssiRequest1.IncidentStatus = 1;
                    SetStateIncidentResponse ssiResponsed1 = (SetStateIncidentResponse)service1.Execute(ssiRequest1);

                    //Если разрешено - получить Разрешение обращения:
                    if (status.Value == 5)
                    {
                        string fetch4 = @"
                                        <fetch mapping=""logical"">
                                            <entity name=""incidentresolution"">
                                                <attribute name=""incidentid""/>
                                                <attribute name=""subject""/>
                                                <attribute name=""statuscode""/>
                                                <attribute name=""statecode""/>
                                                <filter type='and'>
                                                    <condition attribute = 'incidentid' operator='eq' value='" + guidIncident1 + @"'/>
                                                </filter>
                                            </entity>
                                        </fetch>";
                        string result4 = FunctionsToCrmDataWorking.fetchAll(fetch4, service1);
                        XmlDocument xmlDoc4 = new XmlDocument();
                        xmlDoc4.LoadXml(result4);
                        List<CrmIncidentresolution> resList4 = getParsedItemsIncidentresolution(xmlDoc4);
                        foreach (CrmIncidentresolution r3 in resList4)  //Их тут 1 штука, но тем не менее.
                        {
                            //Получить имеющееся Разрешение обращения:
                            Guid guidIncidentresolution1 = new Guid("Write a GUID here.");
                            ColumnSetBase colSet2 = new AllColumns();
                            incidentResolution = (incidentresolution)service1.Retrieve("incidentresolution", guidIncidentresolution1, colSet2);
                        }
                    }
                }

                //Провести динамическое изменение:
                TargetUpdateDynamic updateDynamic2 = new TargetUpdateDynamic();
                updateDynamic2.Entity = entity2;
                UpdateRequest update2 = new UpdateRequest();
                update2.Target = updateDynamic2;
                try
                {
                    UpdateResponse updated2 = (UpdateResponse)service1.Execute(update2);
                }
                catch
                {
                    Console.WriteLine("Error.");
                    Console.ReadLine();
                    //continue;   //Перейти к следующей итерации ближайшего foreach.
                }
                
                //Вернуть Статус и Состояние обратно, если необходимо, т. е.
                //...если они были изменены перед динамическим изменением:
                switch (status.Value)
                {
                    case 1: //Ничего не изменять - и так все как надо.
                        break;
                    case 5: //Перевести в Разрешенные (statecode = 1 = Resolved, statuscode = 5).
                        incidentresolution newIncidentResolution = new incidentresolution();
                        Lookup lookupGuidIncident1 = new Lookup();
                        lookupGuidIncident1.type = EntityName.incident.ToString();
                        lookupGuidIncident1.Value = new Guid("Write a GUID here.");
                        newIncidentResolution.incidentid = lookupGuidIncident1;
                        newIncidentResolution.subject = "Были перенесены данные из Комментария к источнику.";
                        Guid createdIncidentResolutionId = service1.Create(newIncidentResolution);

                        CloseIncidentRequest closeIncidentRequest = new CloseIncidentRequest();
                        closeIncidentRequest.IncidentResolution = newIncidentResolution;
                        closeIncidentRequest.Status = 5;
                        try
                        {
                            CloseIncidentResponse closeIncResponse = (CloseIncidentResponse)service1.Execute(closeIncidentRequest);
                        }
                        catch
                        {
                            Console.WriteLine("Error.");
                            Console.ReadLine();
                            //continue;   //Перейти к следующей итерации ближайшего foreach.
                        }
                        
                        break;
                    case 6: //Перевести в Отмененные (statecode = 2 = Canceled, statuscode = 6).
                        SetStateIncidentRequest ssiRequest2 = new SetStateIncidentRequest();
                        ssiRequest2.EntityId = guidIncident1;
                        ssiRequest2.IncidentState = IncidentState.Canceled;
                        ssiRequest2.IncidentStatus = 6;
                        SetStateIncidentResponse ssiResponsed2 = (SetStateIncidentResponse)service1.Execute(ssiRequest2);
                        
                        break;
                }
                #endregion

                //Возвращение + изменение: Возвращает ФУ по заданному id + изменяет возвращенное ФУ:
                #region Retrieve + Update Gar_fixation Dynamically

                //Возвращает ФУ по заданному id:
                TargetRetrieveDynamic targetRetrieve = new TargetRetrieveDynamic();
                targetRetrieve.EntityName = "gar_fixation";
                targetRetrieve.EntityId = new Guid(r.gar_fixationid);
                RetrieveRequest retrieve = new RetrieveRequest();
                retrieve.Target = targetRetrieve;
                retrieve.ColumnSet = new ColumnSet();   // = new AllColumns();
                retrieve.ColumnSet.AddColumn("gar_service_parameter");
                retrieve.ReturnDynamicEntities = true;
                RetrieveResponse retrieved = (RetrieveResponse)service.Execute(retrieve);
                DynamicEntity entity = (DynamicEntity)retrieved.BusinessEntity;

                //Изменяет возвращенное ФУ:
                CrmBooleanProperty gar_service_parameter = new CrmBooleanProperty();
                gar_service_parameter.Name = "gar_service_parameter";
                gar_service_parameter.Value = new CrmBoolean();
                gar_service_parameter.Value.Value = true;
                entity.Properties.Add(gar_service_parameter);
                TargetUpdateDynamic updateDynamic = new TargetUpdateDynamic();
                updateDynamic.Entity = entity;
                UpdateRequest update = new UpdateRequest();
                update.Target = updateDynamic;
                try
                {
                    UpdateResponse updated = (UpdateResponse)service.Execute(update);
                }
                catch
                {
                }

                #endregion

                //Возвращение + изменение: Возвращает Контракт по заданному id + изменяет возвращенный Контракт:
                #region Retrieve + Update Contract Dynamically

                //Возвращает Контракт по заданному id:
                TargetRetrieveDynamic targetRetrieve = new TargetRetrieveDynamic();
                targetRetrieve.EntityName = "contract";
                targetRetrieve.EntityId = new Guid(r.contractid);
                RetrieveRequest retrieve = new RetrieveRequest();
                retrieve.Target = targetRetrieve;
                retrieve.ColumnSet = new ColumnSet();   // = new AllColumns();
                retrieve.ColumnSet.AddColumn("gar_date_contract_expiry");
                retrieve.ReturnDynamicEntities = true;
                RetrieveResponse retrieved = (RetrieveResponse)service.Execute(retrieve);
                DynamicEntity entity = (DynamicEntity)retrieved.BusinessEntity;

                //Изменяет возвращенный Контракт:
                CrmDateTimeProperty gar_date_contract_expiry_fact = new CrmDateTimeProperty();
                gar_date_contract_expiry_fact.Name = "gar_date_contract_expiry_fact";
                gar_date_contract_expiry_fact.Value = new CrmDateTime(r.gar_date_contract_expiry.ToString());
                //CrmBooleanProperty gar_service_parameter = new CrmBooleanProperty();
                //gar_service_parameter.Name = "gar_service_parameter";
                //gar_service_parameter.Value = new CrmBoolean();
                //gar_service_parameter.Value.Value = true;
                entity.Properties.Add(gar_date_contract_expiry_fact);
                TargetUpdateDynamic updateDynamic = new TargetUpdateDynamic();
                updateDynamic.Entity = entity;
                UpdateRequest update = new UpdateRequest();
                update.Target = updateDynamic;
                try
                {
                    UpdateResponse updated = (UpdateResponse)service.Execute(update);
                }
                catch
                {
                }

                #endregion

                //Возвращение + изменение: Возвращает Строку контракта по заданному id + изменяет возвращенную Строку контракта:
                #region Retrieve + Update Contractdetail Dynamically

                //Возвращает Строку контракта по заданному id:
                TargetRetrieveDynamic targetRetrieve = new TargetRetrieveDynamic();
                targetRetrieve.EntityName = "contractdetail";
                targetRetrieve.EntityId = new Guid(r.contractdetailid);
                RetrieveRequest retrieve = new RetrieveRequest();
                retrieve.Target = targetRetrieve;
                retrieve.ColumnSet = new ColumnSet();   // = new AllColumns();
                retrieve.ColumnSet.AddColumn("gar_end_date_renovation");
                retrieve.ReturnDynamicEntities = true;
                RetrieveResponse retrieved = (RetrieveResponse)service.Execute(retrieve);
                DynamicEntity entity = (DynamicEntity)retrieved.BusinessEntity;

                //Изменяет возвращенную Строку контракта:
                CrmDateTimeProperty gar_end_date_renovation_fact = new CrmDateTimeProperty();
                gar_end_date_renovation_fact.Name = "gar_end_date_renovation_fact";
                gar_end_date_renovation_fact.Value = new CrmDateTime(r.gar_end_date_renovation.ToString());
                entity.Properties.Add(gar_end_date_renovation_fact);
                TargetUpdateDynamic updateDynamic = new TargetUpdateDynamic();
                updateDynamic.Entity = entity;
                UpdateRequest update = new UpdateRequest();
                update.Target = updateDynamic;
                try
                {
                    UpdateResponse updated = (UpdateResponse)service.Execute(update);
                }
                catch
                {
                }

                #endregion

                //Связывает Отношением две Сущности.
                #region Associate two Entities
                
                Guid Gar_kitId = new Guid("00000000-0000-0000-0000-000000000000");
                Guid ProductId = new Guid("00000000-0000-0000-0000-000000000000");
                
                AssociateEntitiesRequest myRequest = new AssociateEntitiesRequest();
                myRequest.Moniker1 = new Moniker();
                myRequest.Moniker1.Id = Gar_kitId;
                myRequest.Moniker1.Name = "gar_kit";
                myRequest.Moniker2 = new Moniker();
                myRequest.Moniker2.Id = ProductId;
                myRequest.Moniker2.Name = "product";
                myRequest.RelationshipName = "gar_gar_kit_product";
                service.Execute(myRequest);
                
                #endregion

                //Создает Историю работ gar_history_jobs_contact и возвращает ее id.
                #region Create JobsHistory gar_history_jobs_contact Dynamically
                string someDate = "7.6.2012";
                string someGuidString = "{9236886A-A0E2-DE11-9387-00155D4E1B14}";
                StringProperty gar_name = new StringProperty();
                gar_name.Name = "gar_name";
                gar_name.Value = "Сертификация";
                PicklistProperty gar_type_jobs = new PicklistProperty();
                gar_type_jobs.Name = "gar_type_jobs";
                gar_type_jobs.Value = new Picklist();
                gar_type_jobs.Value.Value = 1;
                CrmDateTimeProperty gar_date_works = new CrmDateTimeProperty();
                gar_date_works.Name = "gar_date_works";
                gar_date_works.Value = new CrmDateTime();
                gar_date_works.Value.Value = someDate.ToString();
                LookupProperty gar_systemuser = new LookupProperty();
                gar_systemuser.Name = "gar_systemuser";
                gar_systemuser.Value = new Lookup();
                gar_systemuser.Value.type = EntityName.systemuser.ToString();
                gar_systemuser.Value.Value = new Guid(someGuidString);
                PicklistProperty gar_certificate = new PicklistProperty();
                gar_certificate.Name = "gar_certificate";
                gar_certificate.Value = new Picklist();
                gar_certificate.Value.Value = 2;
                CrmBooleanProperty gar_for_lpr = new CrmBooleanProperty();
                gar_for_lpr.Name = "gar_for_lpr";
                gar_for_lpr.Value = new CrmBoolean();
                gar_for_lpr.Value.Value = true; //gar_for_lpr.Value.Value = false;
                LookupProperty gar_contact_history_jobs = new LookupProperty();
                gar_contact_history_jobs.Name = "gar_contact_history_jobs";
                gar_contact_history_jobs.Value = new Lookup();
                gar_contact_history_jobs.Value.type = EntityName.contact.ToString();
                gar_contact_history_jobs.Value.Value = new Guid(someGuidString);
                // Create the DynamicEntity object.
                DynamicEntity jbhistoryEntity = new DynamicEntity();
                // Set the name of the entity type.
                jbhistoryEntity.Name = "gar_history_jobs_contact";
                // Set the properties of the contact.
                jbhistoryEntity.Properties.Add(gar_name);
                jbhistoryEntity.Properties.Add(gar_type_jobs);
                jbhistoryEntity.Properties.Add(gar_date_works);
                jbhistoryEntity.Properties.Add(gar_systemuser);
                jbhistoryEntity.Properties.Add(gar_certificate);
                jbhistoryEntity.Properties.Add(gar_for_lpr);
                jbhistoryEntity.Properties.Add(gar_contact_history_jobs);
                // Create the target.
                TargetCreateDynamic targetCreate2 = new TargetCreateDynamic();
                targetCreate2.Entity = jbhistoryEntity;
                // Create the request object.
                CreateRequest create2 = new CreateRequest();
                // Set the properties of the request object.
                create2.Target = targetCreate2;
                // Execute the request.
                CreateResponse created2 = (CreateResponse)service4.Execute(create2);
                Guid jhID = created2.id;
                #endregion

                //Проверяет, все ли в порядке.
                #region check success
                if (retrieved.BusinessEntity is DynamicEntity)
                {
                    success = true;
                }
                #endregion

                //Удаляет созданные сущности (закомментарено).
                #region Remove Data Required for this Sample
                //service4.Delete(EntityName.contact.ToString(), created.id);
                //service4.Delete(EntityName.account.ToString(), accountID);
                #endregion
            }
            catch (System.Web.Services.Protocols.SoapException ex)
            {
                // Add your error handling code here...
                Console.WriteLine(ex.Message + ex.Detail.InnerXml);
            }

            if (success)
            {
                Console.WriteLine("Работа с сущностями завершена успешно, success==true.");
            }
            Console.WriteLine("Произошла работа с CRM-сервисом.");

            #endregion

            #region Работа с Metadata-сервисом.

            //Создаем сначала объект CrmService:
            CrmService service31 = FunctionsToCrmDataWorking.createCrmService("CmpnyLab", "http://crm4/mscrmservices/2007/crmservice.asmx");

            //Создаем объект Metadata:
            //У нас - имя организации: тестовая база - "CmpnyLab", рабочая база - "Cmpny".
            MetadataService service32 = FunctionsToCrmDataWorking.createMetadataService("CmpnyLab", "http://crm4/mscrmservices/2007/metadataservice.asmx");
            //Работа с Аккаунтом (Бизнес-партнером) (или другими метаданными):
            try
            {
                //Вернуть все возможные значения поля типа "пиклист":
                //Retrieve the attribute metadata.
                RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest();
                attributeRequest.EntityLogicalName = "account";
                attributeRequest.LogicalName = "customertypecode";  //Relationship Type picklist.
                RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)service32.Execute(attributeRequest);
                //Cast the attribute metadata to a picklist metadata.
                PicklistAttributeMetadata picklist = (PicklistAttributeMetadata)attributeResponse.AttributeMetadata;
                //Смотрим, что в пиклисте:
                Console.WriteLine("picklist.Options.Length.ToString(): {0}", picklist.Options.Length.ToString());
                Console.WriteLine("picklist.DefaultValue.ToString(): {0}", picklist.DefaultValue.ToString());
                Console.WriteLine("picklist.SchemaName.ToString(): {0}", picklist.SchemaName.ToString());
                Console.WriteLine("picklist.LogicalName: {0}", picklist.LogicalName);
                Option[] optList = picklist.Options;
                for (int jjj = 0; jjj <= optList.Length - 1; jjj++)
                {
                    Console.WriteLine("№: {0}, Строка: {1}, Значение: {2}", jjj, optList[jjj].Label.UserLocLabel.Label, optList[jjj].Value.Value);
                }

                //Вернуть все возможные значения поля типа "статус":
                RetrieveAttributeRequest attributeRequest2 = new RetrieveAttributeRequest();
                attributeRequest2.EntityLogicalName = "asyncoperation";
                attributeRequest2.LogicalName = "statuscode";  //Relationship Type status.
                RetrieveAttributeResponse attributeResponse2 = (RetrieveAttributeResponse)service32.Execute(attributeRequest2);
                StatusAttributeMetadata st = (StatusAttributeMetadata)attributeResponse2.AttributeMetadata;
                ArrayList optList2 = st.Options;
                Console.WriteLine(optList2.Count);
                for (int jjj = 0; jjj <= optList2.Count - 1; jjj++)
                {
                    Console.WriteLine("d");
                    StatusOption so = (StatusOption)optList2[jjj];
                    Console.WriteLine("№: {0}, Строка: {1}, Значение: {2}, Статус: {3}", jjj, so.Label.UserLocLabel.Label, so.Value.Value, so.State.Value.ToString());
                }

                //Set the default value to "customer" (3).
                picklist.DefaultValue = (object)3;

                //Update the attribute metadata.
                UpdateAttributeRequest updateRequest = new UpdateAttributeRequest();
                updateRequest.Attribute = picklist;
                updateRequest.EntityName = "account";
                updateRequest.MergeLabels = false;

                service32.Execute(updateRequest);

                //Publish the changes.
                PublishXmlRequest request1 = new PublishXmlRequest();
                request.ParameterXml = @"<importexportxml>  
                                       <entities>  
                                          <entity>account</entity>  
                                       </entities>  
                                       <nodes/>  
                                       <securityroles/>  
                                       <settings/>  
                                       <workflows/>  
                                    </importexportxml>";
                PublishXmlResponse response1 = (PublishXmlResponse)service31.Execute(request1);

                Console.WriteLine("Работа с метаданными прошла успешно!");
            }
            catch (System.Web.Services.Protocols.SoapException sex)
            {
                Console.WriteLine("Ошибка работы с метаданными 1: {0}", sex.Detail.OuterXml);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка работы с метаданными 2: {0}", ex.ToString());
            }

            #endregion

            Console.WriteLine("");
            Console.WriteLine("Done!!!");
            Console.WriteLine("");
            Console.WriteLine(prompt);
            Console.ReadLine();
        }

        //Класс для получения объектов "Элемент списка comboBox - Entity".
        public class comboBoxEntity
        {
            int _id;
            string _nameRus;
            string _nameEng;
            public comboBoxEntity(int id, string nameRus, string nameEng)
            {
                this._id = id;
                this._nameRus = nameRus;
                this._nameEng = nameEng;
            }
            public int Id
            {
                get { return this._id; }
            }
            public string NameRus
            {
                get { return this._nameRus; }
            }
            public string NameEng
            {
                get { return this._nameEng; }
            }
            public override string ToString()
            {
                return this._nameRus;   // +" (" + this._nameEng + ")";
            }
        }

        //Класс для получения объектов "Элемент списка comboBox - Attribute".
        public class comboBoxAttribute
        {
            int _id;
            string _nameRus;
            string _nameEng;
            string _type;
            public comboBoxAttribute(int id, string nameRus, string nameEng, string type)
            {
                this._id = id;
                this._nameRus = nameRus;
                this._nameEng = nameEng;
                this._type = type;
            }
            public int Id
            {
                get { return this._id; }
            }
            public string NameRus
            {
                get { return this._nameRus; }
            }
            public string NameEng
            {
                get { return this._nameEng; }
            }
            public string Type
            {
                get { return this._type; }
            }
            public override string ToString()
            {
                return this._nameRus + " (" + this._nameEng + ")" + " (" + this._type + ")";  // +" (" + this._id + ")";
            }
        }

        //Класс для получения объектов "Элемент списка comboBox - Пользователь".
        public class comboBoxSystemuser
        {
            int _id;            //Порядковый номер.
            string _crmGUID;    //ID из CRM.
            string _fullUserName;   //Фамилия.
            public comboBoxSystemuser(int id, string crmGUID, string fullUserName)
            {
                this._id = id;
                this._crmGUID = crmGUID;
                this._fullUserName = fullUserName;
            }
            public int Id
            {
                get { return this._id; }
            }
            public string CrmGUID
            {
                get { return this._crmGUID; }
            }
            public string FullUserName
            {
                get { return this._fullUserName; }
            }
            public override string ToString()
            {
                return this._fullUserName;  // +", " + this._firstname; // +")";  // +" (" + this._id + ")";
            }
        }

        //Класс для получения объектов "Элемент списка".
        public class elementOfList
        {
            public string name;
            public int quantity;
            public elementOfList(string Name, int Quantity)
            {
                this.name = Name;
                this.quantity = Quantity;
            }
        }

        //Класс для получения объектов "Бизнес-партнер" из CRM.
        public class CrmAccount
        {
            public string accountid { get; set; }
            public string name { get; set; }                //Краткое наименование.
            public string telephone1 { get; set; }
            public string telephone2 { get; set; }
            public string fax { get; set; }
            public string gar_telefons { get; set; }
            public string gar_tel_om { get; set; }
            public string address2_telephone1 { get; set; }
            public string address2_telephone2 { get; set; }
            public string address2_fax { get; set; }
            public string address1_telephone1 { get; set; }
            public string address1_telephone2 { get; set; }
            public string address1_fax { get; set; }
            public string statecode { get; set; }           //Статус. 0 - "Активный", 1 - "Неактивный".
            public string statecodename { get; set; }
            public string statuscode { get; set; }          //Состояние.
            public string statuscodename { get; set; }
            public string ownerid { get; set; }             //Ответственный.
            public string owneridname { get; set; }
            public string gar_project { get; set; }         //Проект.
            public string gar_projectname { get; set; }
            public string gar_resultjob { get; set; }       //Результат работы (Г).
            public string gar_resultjobname { get; set; }
            public string gar_date_closed { get; set; }     //Дата завершения работы с организацией (Г).
            public string gar_verification_date { get; set; }       //Контрольная дата (Г).
            public string gar_marketingovogom_list { get; set; }    //Состояние в МС (Г).
            public string gar_marketingovogom_listname { get; set; }
            public string gar_last_call { get; set; }       //Последний по обзвону (Г).
            public string gar_last_callname { get; set; }
            public string gar_latest_roll_call { get; set; }        //Последний по обзвону (СБ).
            public string gar_latest_roll_callname { get; set; }
            public string gar_state_mar_spiske { get; set; }        //Состояние в МС (СБ).
            public string gar_state_mar_spiskename { get; set; }
            public string gar_result_work_mr { get; set; }  //Результат работы (СБ).
            public string gar_result_work_mrname { get; set; }
            public string gar_date_completion { get; set; } //Дата завершения работы с организацией (СБ).
            public string createdby { get; set; } //Создано.
            public string createdbyname { get; set; }
            public string createdon { get; set; } //Дата создания.
            public string modifiedby { get; set; } //Изменено.
            public string modifiedbyname { get; set; }
            public string modifiedon { get; set; } //Дата изменения.
            public string gar_channel_appearance { get; set; } //Источник появления.
            public string gar_channel_appearancename { get; set; }
            public string gar_fixed { get; set; } //Закреплен за.
            public string gar_fixedname { get; set; }
        }

        //Класс для получения объектов "Контакт" из CRM.
        public class CrmContact
        {
            public string fullname { get; set; }
            public string firstname { get; set; }           //Имя Отчество.
            public string lastname { get; set; }            //Фамилия контакта.
            public string parentcustomerid { get; set; }
            public string parentcustomeridname { get; set; }
            public string gar_function { get; set; }        //Должность.
            public string gar_functionname { get; set; }
            public string contactid { get; set; }
            public string statuscode { get; set; }
            public string statecode { get; set; }
            public string ownerid { get; set; }
            public string owneridname { get; set; }
            public string telephone1 { get; set; }          //Рабочий телефон.
            public string mobilephone { get; set; }         //Мобильный телефон.
            public string fax { get; set; }
            public string managerphone { get; set; }
            public string assistantphone { get; set; }
            public string gar_interoffice_telephone { get; set; }
            public string telephone2 { get; set; }
            public string address1_telephone1 { get; set; }
            public string gar_additional_tlefony { get; set; }
            public string gar_telefons { get; set; }
            public string jobtitle { get; set; }            //Комментарий к должности.
            public string gar_department_all { get; set; }  //Отдел.
            public string gar_department_allname { get; set; }
            public string createdon { get; set; }           //Дата создания.
            public string createdby { get; set; }           //Создано.
            public string createdbyname { get; set; }
            public string birthdate { get; set; }           //Дата рождения.
            public string gar_birthday { get; set; }        //День рождения.
            public string gar_birthdayname { get; set; }
            public string gar_birth_month { get; set; }     //Месяц рождения.
            public string gar_birth_monthname { get; set; }
            public string accountrolecode { get; set; }         //Роль 1.
            public string accountrolecodename { get; set; }
            public string gar_accountrolecode_2 { get; set; }   //Роль 2.
            public string gar_accountrolecode_2name { get; set; }
            public string gar_accountrolecode_3 { get; set; }   //Роль 3.
            public string gar_accountrolecode_3name { get; set; }
        }

        //Класс для получения объектов "Контракт" из CRM.
        public class CrmContract
        {
            public string contractid { get; set; }  //GUID.
            public string statuscode { get; set; }  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }   //Статус.
            public string title { get; set; }       //Номер заключенного контракта.
            public string contractnumber { get; set; }                  //Номер контракта.
            public string customerid { get; set; }  //Клиент (Лукап).
            public string customeridname { get; set; }
            public string activeon { get; set; }                        //Дата договора.
            public string gar_activeon_fact { get; set; }               //Дата начала договора.
            public string gar_date_contract_expiry_fact { get; set; }   //Дата окончания договора.
            public string gar_date_contract_expiry { get; set; }        //Дата фактического окончания договора.
        }

        //Класс для получения объектов "Строка контракта" из CRM.
        public class CrmContractdetail
        {
            public string contractdetailid { get; set; }    //GUID.
            public string title { get; set; }               //Наименование.
            public string contractid { get; set; }          //Контракт.
            public string statuscode { get; set; }          //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }           //Статус.
            public string customerid { get; set; }          //Клиент, на кого заключен контракт.
            public string customeridname { get; set; }
            public string gar_gar_stomost_without_discounts_abs { get; set; }   //Стоимость без скидок в АБС.
            public string gar_discount_from_contract_percentage { get; set; }   //Скидка из контракта (%).
            public string gar_discount_from_contract { get; set; }              //Скидка из контракта.
            public string gar_additional_discount_percentage { get; set; }      //Скидка вручную (%).
            public string gar_more_discount { get; set; }                       //Скидка вручную.
            public string gar_contract_number { get; set; } //Номер контракта.
            public string activeon { get; set; }                                //Дата начала.
            public string gar_activeon_fact { get; set; }                       //Дата фактического начала.
            public string gar_end_date_renovation_fact { get; set; }            //Дата окончания.
            public string gar_end_date_renovation { get; set; }                 //Дата фактического окончания.
        }

        //Класс для получения объектов "Соответствие Актов (Услуг) Строкам контракта" из CRM.
        public class CrmGar_correspondence
        {
            public string gar_correspondenceid { get; set; }
            public string gar_name_in_contractdetail { get; set; }  //Наименование Продукта или Типа комплекта в Строке.
            public string gar_name_in_service { get; set; }         //Наименование Ном. или Характ. в Услуге.
            public string gar_nom_to_product { get; set; }          //Это соответствие Номенклатуры Продукту.
            public string gar_char_to_type_set { get; set; }        //Это соответствие Характеристики Типу комплекта.
        }

        //Класс для получения объектов "Возможная сделка" из CRM.
        public class CrmOpportunity
        {
            public string opportunityid { get; set; }
            public string name { get; set; }
            public string customerid { get; set; }
            public string statecode { get; set; }
            public string statuscode { get; set; }
            public string actualclosedate { get; set; }
            public string originatingleadid { get; set; }
            public string campaignid { get; set; }
        }

        //Класс для получения объектов "Тип СПС" из CRM.
        public class CrmGar_sps
        {
            public string gar_spsid { get; set; }
            public string gar_name { get; set; }
        }

        //Класс для получения объектов "Сущность" из CRM.
        public class CrmEntity
        {
            public string Name { get; set; }
            public string LocalizedName { get; set; }
            public string OriginalName { get; set; }
        }

        //Класс для получения объектов "Действие" из CRM.
        public class CrmActivitypointer
        {
            public string activityid { get; set; }
            public string regardingobjectid { get; set; }
            public string subject { get; set; }
            public string createdon { get; set; }
            public string statecode { get; set; }
            public string statuscode { get; set; }
            public string modifiedon { get; set; }                  //Дата последнего изменения.
            public string activitytypecode { get; set; }            //Тип действия.
            public string activitytypecodename { get; set; }
        }

        //Класс для получения объектов "Маркетинговый список" из CRM.
        public class CrmList
        {
            public string listid { get; set; }      //GUID.
            public string listname { get; set; }    //Имя.
            public string statecode { get; set; }   //Статус.
            public string statuscode { get; set; }  //Состояние.
            public string createdon { get; set; }   //Дата создания.
            public string gar_project { get; set; } //Проект.
            public string gar_projectname { get; set; }
        }

        //Класс для получения объектов "Listmember" из CRM.
        //Listmember - это отношение N:N между Бизнес-партнером (account) и Маркетинговым списком (list).
        //Параметр "Name":                      listaccount_association
        //Параметр "Relationship Entity Name":  listmember
        public class CrmListmember
        {
            public string listmemberid { get; set; }
            public string listid { get; set; }      //Идентификатор связанного Маркетингового списка.
            public string entityid { get; set; }    //Идентификатор связанной сущности.
        }

        //Класс для получения объектов "Роль" из CRM.
        public class CrmRole
        {
            public string roleid { get; set; }              //GUID.
            public string name { get; set; }                //Имя.
        }

        //Класс для получения объектов "Systemuserroles" из CRM.
        //Systemuserroles - это отношение N:N между Пользователем (systemuser) и Ролью (role).
        //Параметр "Name":                      systemuserroles_association
        //Параметр "Relationship Entity Name":  systemuserroles
        public class CrmSystemuserroles
        {
            public string systemuserroleid { get; set; }    //GUID.
            public string systemuserid { get; set; }        //Идентификатор связанного Пользователя.
            public string roleid { get; set; }              //Идентификатор связанной Роли.
        }

        //Класс для получения объектов "Note" из CRM.
        //Note - это примечание, которые выводятся в разделе "Примечания" на форме кастомных сущностей.
        public class CrmAnnotation
        {
            public string annotationid { get; set; }    //ID.
            public string notetext { get; set; }        //Description (текст примечания).
            public string subject { get; set; }         //Title (заголовок примечания).
            public string objectid { get; set; }        //Regarding (ID записи, к которой относится примечание).
            public string objecttypecode { get; set; }  //Object Type (тип сущности, к которой относится примечание).
        }

        //Класс для получения объектов "Обращение" из CRM.
        public class CrmIncident
        {
            public string incidentid { get; set; }
            public string ticketnumber { get; set; }            //Номер обращения.
            public string gar_source_incident_1 { get; set; }   //Источник 1.
            public string gar_source_incident_1name { get; set; }
            public string gar_source_incident_2 { get; set; }   //Источник 2.
            public string gar_source_incident_2name { get; set; }
            public string gar_source_incident_3 { get; set; }   //Источник 3.
            public string gar_source_incident_3name { get; set; }
            public string gar_comments_source_1 { get; set; }   //Комментарий к Источнику 1.
            public string gar_comments_source_2 { get; set; }   //Комментарий к Источнику 2.
            public string gar_comments_source_3 { get; set; }   //Комментарий к Источнику 3.
        }

        //Класс для получения объектов "Разрешение обращения" из CRM.
        public class CrmIncidentresolution
        {
            public string activityid { get; set; }              //ID.
            public string incidentid { get; set; }              //Обращение.
            public string subject { get; set; }                 //Тема.
            public string statecode { get; set; }               //Статус.
            public string statuscode { get; set; }              //Состояние.
        }

        //Класс для получения объектов "Звонок" из CRM.
        public class CrmPhonecall
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Статус.
            public string regardingobjectid { get; set; }           //В отношении.
            public string regardingobjectidname { get; set; }
            public string gar_main_account_phonecall { get; set; }  //Родительский клиент.
            public string gar_main_account_phonecallname { get; set; }
            public string gar_contacts_call { get; set; }           //Контакт БП.
            public string gar_contacts_callname { get; set; }
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string gar_1_date { get; set; }                  //Дата семинара_1.
            public string gar_1_answer_client { get; set; }         //Отклик клиента по 1 теме.
            public string gar_1_answer_clientname { get; set; }
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string description { get; set; }                 //Описание.
        }

        //Класс для получения объектов "Отношение с клиентом" из CRM.
        public class CrmCustomerrelationship
        {
            public string customerrelationshipid { get; set; }      //GUID.
            public string customerid { get; set; }                  //ID Клиента № 1.
            public string customeridname { get; set; }
            public string partnerid { get; set; }                   //ID Клиента № 2.
            public string partneridname { get; set; }
        }

        //Класс для получения объектов "Интерес" из CRM.
        public class CrmLead
        {
            public string leadid { get; set; }                      //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Статус.
            public string statecodename { get; set; }
            public string gar_info_povod_lead { get; set; }         //Информационный повод.
            public string gar_info_povod_leadname { get; set; }
            public string estimatedclosedate { get; set; }          //Дата закрытия.
            public string gar_reference_legal_system { get; set; }  //СПС.
            public string gar_reference_legal_systemname { get; set; }
            public string description { get; set; }                 //Описание ТМЦ.
            public string gar_ground_return { get; set; }           //Информация от ОП.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Описание.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string customerid { get; set; }                  //Клиент.
            public string customeridname { get; set; }
            public string gar_result_tmc { get; set; }              //Результат работы ТМЦ.
            public string gar_result_tmcname { get; set; }
        }

        //Класс для получения объектов "Действие сервиса" из CRM.
        public class CrmServiceappointment
        {
            public string activityid { get; set; }          //GUID.
            public string statuscode { get; set; }          //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }           //Статус.
            public string statecodename { get; set; }
            public string regardingobjectid { get; set; }       //В отношении.
            public string regardingobjectidname { get; set; }
            public string gar_priority_training { get; set; }   //Очередность обучения.
            public string gar_priority_trainingname { get; set; }
            public string gar_item1 { get; set; }
            public string gar_item2 { get; set; }
            public string gar_item4 { get; set; }
            public string gar_item5 { get; set; }
            public string gar_item11 { get; set; }
            public string gar_item12 { get; set; }
            public string gar_item13 { get; set; }
            public string gar_item14 { get; set; }
            public string gar_item18 { get; set; }
            public string gar_item19 { get; set; }
            public string gar_item21 { get; set; }
            public string gar_item22 { get; set; }
            public string gar_item23 { get; set; }
            public string gar_item24 { get; set; }
            public string gar_item25 { get; set; }
            public string gar_item26 { get; set; }
            public string gar_item27 { get; set; }
            public string gar_item28 { get; set; }
        }

        //Класс для получения объектов "Проект" из CRM.
        public class CrmGar_project
        {
            public string gar_projectid { get; set; }       //GUID.
            public string gar_name { get; set; }            //Название.
        }

        //Класс для получения объектов "Пользователь" из CRM.
        public class CrmSystemuser
        {
            public string systemuserid { get; set; }        //GUID.
            public string domainname { get; set; }          //Имя пользователя в домене.
            public string firstname { get; set; }           //Имя.
            public string lastname { get; set; }            //Фамилия.
            public string address1_telephone1 { get; set; } //Основной телефон.
            public string address1_telephone2 { get; set; } //Внутренний телефон.
            public string homephone { get; set; }           //Домашний телефон.
            public string mobilephone { get; set; }         //Мобильный телефон.
            public string preferredphonecode { get; set; }  //Основной телефон.
            public string preferredphonecodename { get; set; }
            public string internalemailaddress { get; set; }//Основной адрес эл. почты.
            public string isdisabled { get; set; }          //Статус.
            public string isdisabledname { get; set; }
            public string parentsystemuserid { get; set; }  //Руководитель.
            public string parentsystemuseridname { get; set; }
        }

        //Класс для получения объектов "System Job" из CRM.
        public class CrmAsyncoperation
        {
            public string asyncoperationid { get; set; }    //GUID.
            public string operationtype { get; set; }       //System Job Type.
            public string operationtypename { get; set; }
            public string startedon { get; set; }           //Started On.
            public string createdon { get; set; }           //Created On.
            public string statecode { get; set; }           //Status.
            public string statuscode { get; set; }          //Status Reason.
            public string statuscodename { get; set; }
        }

        //Класс для получения объектов "Финансовые условия" из CRM.
        public class CrmGar_fixation
        {
            public string gar_fixationid { get; set; }
            public string gar_expiration_of { get; set; }       //Окончание действия.
            public string gar_service_parameter { get; set; }   //Служебный параметр.
            public string gar_service_parametername { get; set; }
        }

        //Класс для получения объектов "Задача" из CRM.
        public class CrmTask
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Состояние действия.
            public string statecodename { get; set; }
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string regardingobjectid { get; set; }           //В отношении.
            public string gar_contact { get; set; }                 //Контакт.
            public string gar_contactname { get; set; }
        }

        //Класс для получения объектов "Факс" из CRM.
        public class CrmFax
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Состояние действия.
            public string statecodename { get; set; }
            public string description { get; set; }                 //Описание.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string regardingobjectid { get; set; }           //В отношении.
            public string gar_contact { get; set; }                 //Контакт БП.
            public string gar_contactname { get; set; }
            public string from { get; set; }                        //Отправитель.
            public string to { get; set; }                          //Получатель.
        }

        //Класс для получения объектов "Электронная почта" из CRM.
        public class CrmEmail
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Состояние действия.
            public string statecodename { get; set; }
            public string description { get; set; }                 //Описание.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string regardingobjectid { get; set; }           //В отношении.
            public string from { get; set; }                        //От.
            public string to { get; set; }                          //Кому.
        }

        //Класс для получения объектов "Письмо" из CRM.
        public class CrmLetter
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Состояние действия.
            public string statecodename { get; set; }
            public string description { get; set; }                 //Описание.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string regardingobjectid { get; set; }           //В отношении.
            public string from { get; set; }                        //Отправитель.
            public string to { get; set; }                          //Получатель.
            public string cc { get; set; }                          //Список обучаемых.
        }

        //Класс для получения объектов "Встреча" из CRM.
        public class CrmAppointment
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Статус.
            public string statecodename { get; set; }
            public string description { get; set; }                 //Описание.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string regardingobjectid { get; set; }           //В отношении.
            public string requiredattendees { get; set; }           //Обязательные участники.
            public string optionalattendees { get; set; }           //Необязательные участники.
        }

        //Класс для получения объектов "Контракт от кампании" из CRM.
        public class CrmCampaignresponse
        {
            public string activityid { get; set; }                  //GUID.
            public string statuscode { get; set; }                  //Состояние.
            public string statuscodename { get; set; }
            public string statecode { get; set; }                   //Статус.
            public string statecodename { get; set; }
            public string description { get; set; }                 //Описание.
            public string createdon { get; set; }                   //Дата создания.
            public string modifiedon { get; set; }                  //Дата изменения.
            public string subject { get; set; }                     //Тема.
            public string ownerid { get; set; }                     //Ответственный.
            public string owneridname { get; set; }
            public string gar_contact { get; set; }                 //Контакт.
            public string gar_contactname { get; set; }
        }

        //Класс для получения объектов "Activity Party" из CRM.
        public class CrmActivityparty
        {
            public string activitypartyid { get; set; }             //GUID.
            public string activityid { get; set; }                  //Связанное Действие.
            public string partyid { get; set; }                     //Связанный участник.
        }

        //Класс для получения объектов "Акт" из XML-выдачи из 1С.
        public class CrmAct_DocumentFrom1C
        {
            public string datetime { get; set; }
            public string number { get; set; }
            public string ismark { get; set; }          //Удален.
            public string closed { get; set; }          //Закрыт.
            public string returned { get; set; }
            public string returndate { get; set; }
            public string createdate { get; set; }
            public string maintenance_period_start { get; set; }
            public string maintenance_period_end { get; set; }
            public string uslugi_sum { get; set; }
            public string uslugi_sumNDS { get; set; }
            public string integral { get; set; }
            public string comment { get; set; }
            public string manager { get; set; }
            public string managerdescription { get; set; }
            public string managercode { get; set; }
            public string author { get; set; }
            public string authordescription { get; set; }
            public string authorcode { get; set; }
            public string kontragent { get; set; }
            public string kontragentdescription { get; set; }
            public string kontragentcode { get; set; }
            public string dogovor { get; set; }
            public string dogovordescription { get; set; }
            public string dogovorcode { get; set; }
            public string organization { get; set; }
            public string organizationdescription { get; set; }
            public string organizationcode { get; set; }
            public string usluga { get; set; }
            public string pay { get; set; }
            public List<CrmUsluga_DocumentFrom1C> listUsluga { get; set; }
            public List<CrmScheme_DocumentFrom1C> listScheme { get; set; }
        }

        //Класс для получения объектов "Услуга" из XML-выдачи из 1С.
        public class CrmUsluga_DocumentFrom1C
        {
            public string sim { get; set; }
            public string lineno { get; set; }
            public string amount { get; set; }
            public string sumNDS { get; set; }
            public string NDS { get; set; }
            public string sum { get; set; }
            public string price { get; set; }
            public string actiondescription { get; set; }
            public string nomenklature { get; set; }
            public string nomenklaturedescription { get; set; }
            public string nomenklaturecode { get; set; }
            public string characteristic { get; set; }
            public string characteristicdescription { get; set; }
            public string characteristiccode { get; set; }
        }

        //Класс для получения объектов "Схема оплаты" из XML-выдачи из 1С.
        public class CrmScheme_DocumentFrom1C
        {
            public string lineno { get; set; }
            public string characteristiccode { get; set; }
            public string nomenklaturecode { get; set; }
            public string actiondescr { get; set; }
            public string periodstart { get; set; }
            public string periodend { get; set; }
            public string sum { get; set; }
        }

        //Класс для получения объектов "Акт" из CRM.
        public class CrmGar_act
        {
            public string gar_actid { get; set; }
            public string gar_datetime { get; set; }        //От.
            public string gar_number { get; set; }          //Номер.
            public string gar_kontragent { get; set; }      //Контрагент.
            public string gar_organization { get; set; }    //Организация.
        }

        //Класс для получения объектов "Услуга" из CRM.
        public class CrmGar_service
        {
            public string gar_serviceid { get; set; }
            public string gar_actid { get; set; }           //Акт.
        }

        //Класс для получения объектов "Схема" из CRM.
        public class CrmGar_scheme
        {
            public string gar_schemeid { get; set; }
            public string gar_actid { get; set; }           //Акт.
        }

        //Класс для получения объектов "Разнесение оплаты" из CRM.
        public class CrmGar_paysplit
        {
            public string gar_paysplitid { get; set; }
            public string gar_payid { get; set; }           //Оплата.
            public string gar_payidname { get; set; }
            public string gar_paytype { get; set; }         //Тип оплаты.
            public string gar_paytypename { get; set; }
            public string gar_month { get; set; }           //Месяц.
            public string gar_monthname { get; set; }
            public string gar_year { get; set; }            //Год.
            public string gar_yearname { get; set; }
            public string gar_amount { get; set; }          //Сумма.
            public string gar_contract { get; set; }        //Контракт.
            public string gar_contractname { get; set; }
            public string gar_accountid { get; set; }       //Бизнес-партнер.
            public string gar_accountidname { get; set; }
        }

        //Класс для получения объектов "Объект неопределенного заранее типа" из CRM.
        //Используется, когда в ходе выполнения программы определяется, какого типа объект из нескольких заранее предусмотренных следует выбрать из базы в данную переменную.
        //После того, как будет определен тип объекта, фетч-запрос вернет список объектов этого типа, при этом в наличии будут атрибуты только относящиеся к выбранному типу объекта.
        public class CrmIndeterminatedObject
        {
            public string accountcategorycode { get; set; }         //Категория.                /"Бизнес-партнер".
            public string accountcategorycodename { get; set; }
            public string gar_arhsb { get; set; }                   //Основной день визита.     /"Бизнес-партнер".
            public string gar_arhsbname { get; set; }
            public string preferredappointmentdaycode { get; set; } //Основной день обновления. /"Бизнес-партнер".
            public string preferredappointmentdaycodename { get; set; }
            public string ownerid { get; set; }                     //Ответственный.            /"Бизнес-партнер".
            //Ответственный.            /"Контакт".
            //Ответственный №1.         /"Обращение".
            public string owneridname { get; set; }
            public string businesstypecode { get; set; }            //Режим налогообложения.    /"Бизнес-партнер".
            public string businesstypecodename { get; set; }
            public string statuscode { get; set; }                  //Состояние.                /"Бизнес-партнер".
            public string statuscodename { get; set; }
            public string createdby { get; set; }                   //Создано.                  /"Бизнес-партнер".
            public string createdbyname { get; set; }
            public string createdon { get; set; }                   //Дата создания.            /"Бизнес-партнер".
            public string modifiedby { get; set; }                  //Изменено.                 /"Бизнес-партнер".
            public string modifiedbyname { get; set; }
            public string modifiedon { get; set; }                  //Дата изменения.           /"Бизнес-партнер".
            public string gar_channel_appearance { get; set; }      //Источник появления.       /"Бизнес-партнер".
            public string gar_channel_appearancename { get; set; }

            public string preferredcontactmethodcode { get; set; }  //Основной способ связи.    /"Контакт".
            public string preferredcontactmethodcodename { get; set; }
            public string accountrolecode { get; set; }             //Роль 1.                   /"Контакт".
            public string accountrolecodename { get; set; }
            public string customertypecode { get; set; }            //Тип отношений.            /"Контакт".
            public string customertypecodename { get; set; }

            public string caseorigincode { get; set; }              //Канал.                    /"Обращение".
            public string caseorigincodename { get; set; }
            public string subjectid { get; set; }                   //Тема.                     /"Обращение".
            public string subjectidname { get; set; }
            public string casetypecode { get; set; }                //Тип.                      /"Обращение".
            public string casetypecodename { get; set; }

            public string gar_info_povod_lead { get; set; }         //Информационный повод.     /"Интерес".
            public string gar_info_povod_leadname { get; set; }
            public string leadsourcecode { get; set; }              //Источник.                 /"Интерес".
            public string leadsourcecodename { get; set; }
            public string gar_reference_legal_system { get; set; }  //СПС.                      /"Интерес".
            public string gar_reference_legal_systemname { get; set; }
            public string gar_result_tmc { get; set; }              //Результат работы ТМЦ.     /"Интерес".
            public string gar_result_tmcname { get; set; }

            public string gar_businessunit { get; set; }            //Отдел.                    /"История работы".
            public string gar_businessunitname { get; set; }
            public string gar_systemuser { get; set; }              //Сотрудник.                /"История работы".
            public string gar_systemusername { get; set; }
            public string gar_projects { get; set; }                //Проект.                   /"История работы".
            public string gar_projectsname { get; set; }
            public string gar_result { get; set; }                  //Результат.                /"История работы".
            public string gar_resultname { get; set; }
            public string gar_list { get; set; }                    //Маркетинговый список.     /"История работы".
            public string gar_listname { get; set; }

            public string gar_item_1 { get; set; }                  //Знаете как зовут обслуживающего Вас сотрудника?       /"Действие сервиса".
            public string gar_item_1name { get; set; }
            public string gar_item3 { get; set; }                   //Как Вы оцениваете его профессиональный уровень?       /"Действие сервиса".
            public string gar_item3name { get; set; }
            public string gar_item17 { get; set; }                  //Как Вы оцениваете качество обслуживания?              /"Действие сервиса".
            public string gar_item17name { get; set; }
            public string gar_item9 { get; set; }                   //Как часто в работе Вы используете ИПО ГАРАНТ?         /"Действие сервиса".
            public string gar_item9name { get; set; }
            public string gar_item16 { get; set; }                  //Как планируете получать правовую информацию далее?    /"Действие сервиса".
            public string gar_item16name { get; set; }
        }

        //Класс для получения объектов неизвестного заранее типа (без параметров) из CRM.
        public class CrmAnyObjectWithNoParameters
        {
            //Нет обрабатываемых параметров для объектов этого типа.
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Бизнес-партнер".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAccount> ParsedItems - список объектов CrmAccount, представляющих собой объекты "Бизнес-партнер" из CRM.
        public static List<CrmAccount> getParsedItemsAccount(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAccount> DataCollection = getParsedItemsAccount(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAccount> getParsedItemsAccount(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmAccount> DataCollection = new List<CrmAccount>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAccount currentNode = new CrmAccount();
                    currentNode.accountid               = FunctionsToXmlDataWorking.ParseNodeValue(node, "accountid");
                    currentNode.name                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "name");
                    currentNode.telephone1              = FunctionsToXmlDataWorking.ParseNodeValue(node, "telephone1");
                    currentNode.telephone2              = FunctionsToXmlDataWorking.ParseNodeValue(node, "telephone2");
                    currentNode.fax                     = FunctionsToXmlDataWorking.ParseNodeValue(node, "fax");
                    currentNode.gar_telefons            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_telefons");
                    currentNode.gar_tel_om              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_tel_om");
                    currentNode.address2_telephone1     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address2_telephone1");
                    currentNode.address2_telephone2     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address2_telephone2");
                    currentNode.address2_fax            = FunctionsToXmlDataWorking.ParseNodeValue(node, "address2_fax");
                    currentNode.address1_telephone1     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_telephone1");
                    currentNode.address1_telephone2     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_telephone2");
                    currentNode.address1_fax            = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_fax");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.gar_project             = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_project");//Проект (Лукап).
                    currentNode.gar_projectname         = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_project", "name");
                    currentNode.gar_resultjob           = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_resultjob");//Результат работы (Г).
                    currentNode.gar_resultjobname       = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_resultjob", "name");
                    currentNode.gar_date_closed         = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_date_closed");//Дата завершения работы с организацией (Г).
                    currentNode.gar_verification_date   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_verification_date");//Контрольная дата (Г).
                    currentNode.gar_marketingovogom_list        = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_marketingovogom_list");//Состояние в МС (Г).
                    currentNode.gar_marketingovogom_listname    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_marketingovogom_list", "name");
                    currentNode.gar_last_call           = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_last_call");//Последний по обзвону (Г).
                    currentNode.gar_last_callname       = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_last_call", "name");
                    currentNode.gar_latest_roll_call            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_latest_roll_call");//Последний по обзвону (СБ).
                    currentNode.gar_latest_roll_callname        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_latest_roll_call", "name");
                    currentNode.gar_state_mar_spiske            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_state_mar_spiske");//Состояние в МС (СБ).
                    currentNode.gar_state_mar_spiskename        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_state_mar_spiske", "name");
                    currentNode.gar_result_work_mr      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_result_work_mr");//Результат работы (СБ).
                    currentNode.gar_result_work_mrname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_result_work_mr", "name");
                    currentNode.gar_date_completion     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_date_completion");//Дата завершения работы с организацией (СБ).
                    currentNode.createdby               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdby");//Создано.
                    currentNode.createdbyname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "createdby", "name");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedby              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedby");//Изменено.
                    currentNode.modifiedbyname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "modifiedby", "name");
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.gar_channel_appearance  = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_channel_appearance");//Источник появления.
                    currentNode.gar_channel_appearancename      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_channel_appearance", "name");
                    currentNode.gar_fixed               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_fixed");//Закреплен за.
                    currentNode.gar_fixedname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_fixed", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Контакт".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmContact> ParsedItems - список объектов CrmContact, представляющих собой объекты "Контакт" из CRM.
        public static List<CrmContact> getParsedItemsContact(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmContact> DataCollection = getParsedItemsContact(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmContact> getParsedItemsContact(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmContact> DataCollection = new List<CrmContact>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmContact currentNode = new CrmContact();
                    currentNode.fullname                = FunctionsToXmlDataWorking.ParseNodeValue(node, "fullname");
                    currentNode.firstname               = FunctionsToXmlDataWorking.ParseNodeValue(node, "firstname");
                    currentNode.lastname                = FunctionsToXmlDataWorking.ParseNodeValue(node, "lastname");
                    currentNode.parentcustomerid        = FunctionsToXmlDataWorking.ParseNodeValue(node, "parentcustomerid");
                    currentNode.parentcustomeridname    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "parentcustomerid", "name");
                    currentNode.gar_function            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_function");
                    currentNode.gar_functionname        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_function", "name");
                    currentNode.contactid               = FunctionsToXmlDataWorking.ParseNodeValue(node, "contactid");
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.telephone1              = FunctionsToXmlDataWorking.ParseNodeValue(node, "telephone1");
                    currentNode.mobilephone             = FunctionsToXmlDataWorking.ParseNodeValue(node, "mobilephone");
                    currentNode.fax                     = FunctionsToXmlDataWorking.ParseNodeValue(node, "fax");
                    currentNode.managerphone            = FunctionsToXmlDataWorking.ParseNodeValue(node, "managerphone");
                    currentNode.assistantphone          = FunctionsToXmlDataWorking.ParseNodeValue(node, "assistantphone");
                    currentNode.gar_interoffice_telephone   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_interoffice_telephone");
                    currentNode.telephone2              = FunctionsToXmlDataWorking.ParseNodeValue(node, "telephone2");
                    currentNode.address1_telephone1     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_telephone1");
                    currentNode.gar_additional_tlefony  = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_additional_tlefony");
                    currentNode.gar_telefons            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_telefons");
                    currentNode.jobtitle                = FunctionsToXmlDataWorking.ParseNodeValue(node, "jobtitle");
                    currentNode.gar_department_all      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_department_all"); //Отдел.
                    currentNode.gar_department_allname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_department_all", "name");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");          //Дата создания.
                    currentNode.createdby               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdby");          //Создано.
                    currentNode.createdbyname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "createdby", "name");
                    currentNode.birthdate               = FunctionsToXmlDataWorking.ParseNodeValue(node, "birthdate");//Дата рождения.
                    currentNode.gar_birthday            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_birthday");//День рождения.
                    currentNode.gar_birthdayname        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_birthday", "name");
                    currentNode.gar_birth_month         = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_birth_month");//Месяц рождения.
                    currentNode.gar_birth_monthname     = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_birth_month", "name");
                    currentNode.accountrolecode         = FunctionsToXmlDataWorking.ParseNodeValue(node, "accountrolecode");//Роль 1.
                    currentNode.accountrolecodename     = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "accountrolecode", "name");
                    currentNode.gar_accountrolecode_2   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_accountrolecode_2");//Роль 2.
                    currentNode.gar_accountrolecode_2name   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_accountrolecode_2", "name");
                    currentNode.gar_accountrolecode_3   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_accountrolecode_3");//Роль 3.
                    currentNode.gar_accountrolecode_3name   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_accountrolecode_3", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Контракт".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmContract> ParsedItems - список объектов CrmContract, представляющих собой объекты "Контракт" из CRM.
        public static List<CrmContract> getParsedItemsContract(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmContract> DataCollection = getParsedItemsContract(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmContract> getParsedItemsContract(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmContract> DataCollection = new List<CrmContract>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmContract currentNode = new CrmContract();
                    currentNode.contractid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.title                   = FunctionsToXmlDataWorking.ParseNodeValue(node, "title");//Номер заключенного контракта.
                    currentNode.contractnumber          = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractnumber");//Номер контракта.
                    currentNode.customerid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");//Клиент.
                    currentNode.customeridname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customerid", "name");
                    currentNode.activeon                = FunctionsToXmlDataWorking.ParseNodeValue(node, "activeon");//Дата договора.
                    currentNode.gar_activeon_fact       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_activeon_fact");//Дата начала договора.
                    currentNode.gar_date_contract_expiry_fact   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_date_contract_expiry_fact");//Дата окончания договора.
                    currentNode.gar_date_contract_expiry        = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_date_contract_expiry");//Дата фактического окончания договора.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Строка контракта".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmContractdetail> ParsedItems - список объектов CrmContractdetail, представляющих собой объекты "Строка контракта" из CRM.
        public static List<CrmContractdetail> getParsedItemsContractdetail(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmContractdetail> DataCollection = getParsedItemsContractdetail(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmContractdetail> getParsedItemsContractdetail(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmContractdetail> DataCollection = new List<CrmContractdetail>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmContractdetail currentNode = new CrmContractdetail();
                    currentNode.contractdetailid        = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractdetailid");//GUID.
                    currentNode.title                   = FunctionsToXmlDataWorking.ParseNodeValue(node, "title");//Наименование.
                    currentNode.contractid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractid");//Контракт.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.customerid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");//Клиент, на кого заключен контракт.
                    currentNode.customeridname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customerid", "name");
                    currentNode.gar_gar_stomost_without_discounts_abs   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_gar_stomost_without_discounts_abs");//Стоимость без скидок в АБС.
                    currentNode.gar_discount_from_contract_percentage   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_discount_from_contract_percentage");//Скидка из контракта (%).
                    currentNode.gar_discount_from_contract              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_discount_from_contract");//Скидка из контракта.
                    currentNode.gar_additional_discount_percentage      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_additional_discount_percentage");//Скидка вручную (%).
                    currentNode.gar_more_discount                       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_more_discount");//Скидка вручную.
                    currentNode.gar_contract_number     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contract_number");//Номер контракта.
                    currentNode.activeon                = FunctionsToXmlDataWorking.ParseNodeValue(node, "activeon");//Дата начала.
                    currentNode.gar_activeon_fact       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_activeon_fact");//Дата фактического начала.
                    currentNode.gar_end_date_renovation_fact            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_end_date_renovation_fact");//Дата окончания.
                    currentNode.gar_end_date_renovation = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_end_date_renovation");//Дата фактического окончания.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Соответствие Актов (Услуг) Строкам контракта".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_correspondence> ParsedItems - список объектов CrmGar_correspondence, представляющих собой объекты "Соответствие Актов (Услуг) Строкам контракта" из CRM.
        public static List<CrmGar_correspondence> getParsedItemsGar_correspondence(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_correspondence> DataCollection = getParsedItemsGar_correspondence(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_correspondence> getParsedItemsGar_correspondence(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_correspondence> DataCollection = new List<CrmGar_correspondence>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_correspondence currentNode = new CrmGar_correspondence();
                    currentNode.gar_correspondenceid    = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_correspondenceid");
                    currentNode.gar_name_in_contractdetail  = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_name_in_contractdetail");//Наименование Продукта или Типа комплекта в Строке.
                    currentNode.gar_name_in_service     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_name_in_service");//Наименование Ном. или Характ. в Услуге.
                    currentNode.gar_nom_to_product      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_nom_to_product");//Это соответствие Номенклатуры Продукту.
                    currentNode.gar_char_to_type_set    = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_char_to_type_set");//Это соответствие Характеристики Типу комплекта.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Возможная сделка".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmOpportunity> ParsedItems - список объектов CrmOpportunity, представляющих собой объекты "Возможная сделка" из CRM.
        public static List<CrmOpportunity> getParsedItemsOpportunity(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmOpportunity> DataCollection = getParsedItemsOpportunity(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmOpportunity> getParsedItemsOpportunity(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmOpportunity> DataCollection = new List<CrmOpportunity>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmOpportunity currentNode = new CrmOpportunity();
                    currentNode.opportunityid           = FunctionsToXmlDataWorking.ParseNodeValue(node, "opportunityid");
                    currentNode.name                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "name");
                    currentNode.customerid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");
                    currentNode.actualclosedate         = FunctionsToXmlDataWorking.ParseNodeValue(node, "actualclosedate");
                    currentNode.originatingleadid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "originatingleadid");
                    currentNode.campaignid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "campaignid");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Тип СПС".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_sps> ParsedItems - список объектов CrmGar_sps, представляющих собой объекты "Тип СПС" из CRM.
        public static List<CrmGar_sps> getParsedItemsGar_sps(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_sps> DataCollection = getParsedItemsGar_sps(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_sps> getParsedItemsGar_sps(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_sps> DataCollection = new List<CrmGar_sps>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_sps currentNode = new CrmGar_sps();
                    currentNode.gar_spsid               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_spsid");
                    currentNode.gar_name                = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-файле в список объектов "Сущность".
        //Входные параметры:
        //string Url - путь к Xml-файлу.
        //Выходные параметры:
        //List<CrmEntity> ParsedItems - список объектов CrmEntity, представляющих собой объекты "Сущность" из CRM.
        public static List<CrmEntity> getParsedItemsEntity(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                XmlNodeList Itemslist = xmlDoc.SelectNodes("ImportExportXml/Entities/Entity");

                List<CrmEntity> DataCollection = new List<CrmEntity>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmEntity currentNode = new CrmEntity();
                    currentNode.Name                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "Name");
                    currentNode.LocalizedName           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "Name", "LocalizedName");
                    currentNode.OriginalName            = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "Name", "OriginalName");

                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Действие".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmActivitypointer> ParsedItems - список объектов CrmActivitypointer, представляющих собой объекты "Действие" из CRM.
        public static List<CrmActivitypointer> getParsedItemsActivitypointer(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmActivitypointer> DataCollection = getParsedItemsActivitypointer(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmActivitypointer> getParsedItemsActivitypointer(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmActivitypointer> DataCollection = new List<CrmActivitypointer>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmActivitypointer currentNode = new CrmActivitypointer();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата последнего изменения.
                    currentNode.activitytypecode        = FunctionsToXmlDataWorking.ParseNodeValue(node, "activitytypecode");//Тип действия.
                    currentNode.activitytypecodename    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "activitytypecode", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Маркетинговый список".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmList> ParsedItems - список объектов CrmList, представляющих собой объекты "Маркетинговый список" из CRM.
        public static List<CrmList> getParsedItemsList(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmList> DataCollection = getParsedItemsList(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmList> getParsedItemsList(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmList> DataCollection = new List<CrmList>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmList currentNode = new CrmList();
                    currentNode.listid                  = FunctionsToXmlDataWorking.ParseNodeValue(node, "listid");//GUID.
                    currentNode.listname                = FunctionsToXmlDataWorking.ParseNodeValue(node, "listname");//Имя.
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.gar_project             = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_project");//Проект.
                    currentNode.gar_projectname         = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_project", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Listmember".
        //Listmember - это отношение N:N между Бизнес-партнером (account) и Маркетинговым списком (list).
        //Параметр "Name":                      listaccount_association
        //Параметр "Relationship Entity Name":  listmember
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmListmember> ParsedItems - список объектов CrmListmember, представляющих собой объекты "Listmember" из CRM.
        public static List<CrmListmember> getParsedItemsListmember(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmListmember> DataCollection = getParsedItemsListmember(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmListmember> getParsedItemsListmember(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmListmember> DataCollection = new List<CrmListmember>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmListmember currentNode = new CrmListmember();
                    currentNode.listmemberid            = FunctionsToXmlDataWorking.ParseNodeValue(node, "listmemberid");
                    currentNode.listid                  = FunctionsToXmlDataWorking.ParseNodeValue(node, "listid");
                    currentNode.entityid                = FunctionsToXmlDataWorking.ParseNodeValue(node, "entityid");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Роль".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmRole> ParsedItems - список объектов CrmRole, представляющих собой объекты "Роль" из CRM.
        public static List<CrmRole> getParsedItemsRole(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmRole> DataCollection = getParsedItemsRole(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmRole> getParsedItemsRole(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmRole> DataCollection = new List<CrmRole>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmRole currentNode = new CrmRole();
                    currentNode.roleid                  = FunctionsToXmlDataWorking.ParseNodeValue(node, "roleid");
                    currentNode.name                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Systemuserroles".
        //Systemuserroles - это отношение N:N между Пользователем (systemuser) и Ролью (role).
        //Параметр "Name":                      systemuserroles_association
        //Параметр "Relationship Entity Name":  systemuserroles
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmSystemuserroles> ParsedItems - список объектов CrmSystemuserroles, представляющих собой объекты "Systemuserroles" из CRM.
        public static List<CrmSystemuserroles> getParsedItemsSystemuserroles(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmSystemuserroles> DataCollection = getParsedItemsSystemuserroles(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmSystemuserroles> getParsedItemsSystemuserroles(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmSystemuserroles> DataCollection = new List<CrmSystemuserroles>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmSystemuserroles currentNode = new CrmSystemuserroles();
                    currentNode.systemuserroleid        = FunctionsToXmlDataWorking.ParseNodeValue(node, "systemuserroleid");
                    currentNode.systemuserid            = FunctionsToXmlDataWorking.ParseNodeValue(node, "systemuserid");
                    currentNode.roleid                  = FunctionsToXmlDataWorking.ParseNodeValue(node, "roleid");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Note".
        //Note - это примечание, которые выводятся в разделе "Примечания" на форме кастомных сущностей.
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAnnotation> ParsedItems - список объектов CrmAnnotation, представляющих собой объекты "Note" из CRM.
        public static List<CrmAnnotation> getParsedItemsAnnotation(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAnnotation> DataCollection = getParsedItemsAnnotation(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAnnotation> getParsedItemsAnnotation(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmAnnotation> DataCollection = new List<CrmAnnotation>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAnnotation currentNode = new CrmAnnotation();
                    currentNode.annotationid            = FunctionsToXmlDataWorking.ParseNodeValue(node, "annotationid");
                    currentNode.notetext                = FunctionsToXmlDataWorking.ParseNodeValue(node, "notetext");
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");
                    currentNode.objectid                = FunctionsToXmlDataWorking.ParseNodeValue(node, "objectid");
                    currentNode.objecttypecode          = FunctionsToXmlDataWorking.ParseNodeValue(node, "objecttypecode");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Обращение".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmIncident> ParsedItems - список объектов CrmIncident, представляющих собой объекты "Обращение" из CRM.
        public static List<CrmIncident> getParsedItemsIncident(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmIncident> DataCollection = getParsedItemsIncident(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmIncident> getParsedItemsIncident(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmIncident> DataCollection = new List<CrmIncident>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmIncident currentNode = new CrmIncident();
                    currentNode.incidentid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "incidentid");
                    currentNode.ticketnumber            = FunctionsToXmlDataWorking.ParseNodeValue(node, "ticketnumber");//Номер обращения.
                    currentNode.gar_source_incident_1   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_source_incident_1");//Источник 1.
                    currentNode.gar_source_incident_1name   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_source_incident_1name", "name");
                    currentNode.gar_source_incident_2   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_source_incident_2");//Источник 2.
                    currentNode.gar_source_incident_2name   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_source_incident_2name", "name");
                    currentNode.gar_source_incident_3   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_source_incident_3");//Источник 3.
                    currentNode.gar_source_incident_3name   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_source_incident_3name", "name");
                    currentNode.gar_comments_source_1   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_comments_source_1");//Комментарий к Источнику 1.
                    currentNode.gar_comments_source_2   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_comments_source_2");//Комментарий к Источнику 2.
                    currentNode.gar_comments_source_3   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_comments_source_3");//Комментарий к Источнику 3.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Разрешение обращения".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmIncidentresolution> ParsedItems - список объектов CrmIncidentresolution, представляющих собой объекты "Разрешение обращения" из CRM.
        public static List<CrmIncidentresolution> getParsedItemsIncidentresolution(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmIncidentresolution> DataCollection = getParsedItemsIncidentresolution(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmIncidentresolution> getParsedItemsIncidentresolution(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmIncidentresolution> DataCollection = new List<CrmIncidentresolution>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmIncidentresolution currentNode = new CrmIncidentresolution();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//ID.
                    currentNode.incidentid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "incidentid");//Обращение.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Звонок".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmPhonecall> ParsedItems - список объектов CrmPhonecall, представляющих собой объекты "Звонок" из CRM.
        public static List<CrmPhonecall> getParsedItemsPhonecall(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmPhonecall> DataCollection = getParsedItemsPhonecall(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmPhonecall> getParsedItemsPhonecall(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmPhonecall> DataCollection = new List<CrmPhonecall>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmPhonecall currentNode = new CrmPhonecall();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.regardingobjectidname   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "regardingobjectid", "name");
                    currentNode.gar_main_account_phonecall      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_main_account_phonecall");//Родительский клиент.
                    currentNode.gar_main_account_phonecallname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_main_account_phonecall", "name");
                    currentNode.gar_contacts_call       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contacts_call");//Контакт БП.
                    currentNode.gar_contacts_callname   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_contacts_call", "name");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.gar_1_date              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_1_date");//Дата семинара_1.
                    currentNode.gar_1_answer_client     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_1_answer_client");//Отклик клиента по 1 теме.
                    currentNode.gar_1_answer_clientname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_1_answer_client", "name");
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Отношение с клиентом".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmCustomerrelationship> ParsedItems - список объектов CrmCustomerrelationship, представляющих собой объекты "Отношение с клиентом" из CRM.
        public static List<CrmCustomerrelationship> getParsedItemsCustomerrelationship(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmCustomerrelationship> DataCollection = getParsedItemsCustomerrelationship(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmCustomerrelationship> getParsedItemsCustomerrelationship(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmCustomerrelationship> DataCollection = new List<CrmCustomerrelationship>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmCustomerrelationship currentNode = new CrmCustomerrelationship();
                    currentNode.customerrelationshipid  = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerrelationshipid");//GUID.
                    currentNode.customerid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");//ID Клиента № 1.
                    currentNode.customeridname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customerid", "name");
                    currentNode.partnerid               = FunctionsToXmlDataWorking.ParseNodeValue(node, "partnerid");//ID Клиента № 2.
                    currentNode.partneridname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "partnerid", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Интерес".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmLead> ParsedItems - список объектов CrmLead, представляющих собой объекты "Интерес" из CRM.
        public static List<CrmLead> getParsedItemsLead(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmLead> DataCollection = getParsedItemsLead(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmLead> getParsedItemsLead(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmLead> DataCollection = new List<CrmLead>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmLead currentNode = new CrmLead();
                    currentNode.leadid                  = FunctionsToXmlDataWorking.ParseNodeValue(node, "leadid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.gar_info_povod_lead     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_info_povod_lead");//Информационный повод.
                    currentNode.gar_info_povod_leadname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_info_povod_lead", "name");
                    currentNode.estimatedclosedate      = FunctionsToXmlDataWorking.ParseNodeValue(node, "estimatedclosedate");//Дата закрытия.
                    currentNode.gar_reference_legal_system      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_reference_legal_system");//СПС.
                    currentNode.gar_reference_legal_systemname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_reference_legal_system", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание ТМЦ.
                    currentNode.gar_ground_return       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_ground_return");//Информация от ОП.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Описание.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.customerid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");//Клиент.
                    currentNode.customeridname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customerid", "name");
                    currentNode.gar_result_tmc          = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_result_tmc");//Результат работы ТМЦ.
                    currentNode.gar_result_tmcname      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_result_tmc", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Действие сервиса".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmServiceappointment> ParsedItems - список объектов CrmServiceappointment, представляющих собой объекты "Действие сервиса" из CRM.
        public static List<CrmServiceappointment> getParsedItemsServiceappointment(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmServiceappointment> DataCollection = getParsedItemsServiceappointment(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmServiceappointment> getParsedItemsServiceappointment(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmServiceappointment> DataCollection = new List<CrmServiceappointment>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmServiceappointment currentNode = new CrmServiceappointment();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.regardingobjectidname   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "regardingobjectid", "name");
                    currentNode.gar_priority_training   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_priority_training");//Очередность обучения.
                    currentNode.gar_priority_trainingname   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_priority_training", "name");
                    currentNode.gar_item1               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item1");
                    currentNode.gar_item2               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item2");
                    currentNode.gar_item4               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item4");
                    currentNode.gar_item5               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item5");
                    currentNode.gar_item11              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item11");
                    currentNode.gar_item12              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item12");
                    currentNode.gar_item13              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item13");
                    currentNode.gar_item14              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item14");
                    currentNode.gar_item18              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item18");
                    currentNode.gar_item19              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item19");
                    currentNode.gar_item21              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item21");
                    currentNode.gar_item22              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item22");
                    currentNode.gar_item23              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item23");
                    currentNode.gar_item24              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item24");
                    currentNode.gar_item25              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item25");
                    currentNode.gar_item26              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item26");
                    currentNode.gar_item27              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item27");
                    currentNode.gar_item28              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item28");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Проект".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_project> ParsedItems - список объектов CrmGar_project, представляющих собой объекты "Проект" из CRM.
        public static List<CrmGar_project> getParsedItemsGar_project(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_project> DataCollection = getParsedItemsGar_project(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_project> getParsedItemsGar_project(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_project> DataCollection = new List<CrmGar_project>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_project currentNode = new CrmGar_project();
                    currentNode.gar_projectid           = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_projectid");//GUID.
                    currentNode.gar_name                = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_name");//Название.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Пользователь".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmSystemuser> ParsedItems - список объектов CrmSystemuser, представляющих собой объекты "Пользователь" из CRM.
        public static List<CrmSystemuser> getParsedItemsSystemuser(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmSystemuser> DataCollection = getParsedItemsSystemuser(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmSystemuser> getParsedItemsSystemuser(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmSystemuser> DataCollection = new List<CrmSystemuser>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmSystemuser currentNode = new CrmSystemuser();
                    currentNode.systemuserid            = FunctionsToXmlDataWorking.ParseNodeValue(node, "systemuserid");//GUID.
                    currentNode.domainname              = FunctionsToXmlDataWorking.ParseNodeValue(node, "domainname");//Имя пользователя в домене.
                    currentNode.firstname               = FunctionsToXmlDataWorking.ParseNodeValue(node, "firstname");//Имя.
                    currentNode.lastname                = FunctionsToXmlDataWorking.ParseNodeValue(node, "lastname");//Фамилия.
                    currentNode.address1_telephone1     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_telephone1");//Основной телефон.
                    currentNode.address1_telephone2     = FunctionsToXmlDataWorking.ParseNodeValue(node, "address1_telephone2");//Внутренний телефон.
                    currentNode.homephone               = FunctionsToXmlDataWorking.ParseNodeValue(node, "homephone");//Домашний телефон.
                    currentNode.mobilephone             = FunctionsToXmlDataWorking.ParseNodeValue(node, "mobilephone");//Мобильный телефон.
                    currentNode.preferredphonecode      = FunctionsToXmlDataWorking.ParseNodeValue(node, "preferredphonecode");//Основной телефон.
                    currentNode.preferredphonecodename  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "preferredphonecode", "name");
                    currentNode.internalemailaddress    = FunctionsToXmlDataWorking.ParseNodeValue(node, "internalemailaddress");//Основной адрес эл. почты.
                    currentNode.isdisabled              = FunctionsToXmlDataWorking.ParseNodeValue(node, "isdisabled");//Статус.
                    currentNode.isdisabledname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "isdisabled", "name");
                    currentNode.parentsystemuserid      = FunctionsToXmlDataWorking.ParseNodeValue(node, "parentsystemuserid");//Руководитель.
                    currentNode.parentsystemuseridname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "parentsystemuserid", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "System Job".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAsyncoperation> ParsedItems - список объектов CrmAsyncoperation, представляющих собой объекты "System Job" из CRM.
        public static List<CrmAsyncoperation> getParsedItemsAsyncoperation(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAsyncoperation> DataCollection = getParsedItemsAsyncoperation(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAsyncoperation> getParsedItemsAsyncoperation(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmAsyncoperation> DataCollection = new List<CrmAsyncoperation>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAsyncoperation currentNode = new CrmAsyncoperation();
                    currentNode.asyncoperationid        = FunctionsToXmlDataWorking.ParseNodeValue(node, "asyncoperationid");//GUID.
                    currentNode.operationtype           = FunctionsToXmlDataWorking.ParseNodeValue(node, "operationtype");//System Job Type.
                    currentNode.operationtypename       = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "operationtype", "name");
                    currentNode.startedon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "startedon");//Started On.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Created On.
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Status.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Status Reason.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-файле в список объектов "Финансовые условия".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_fixation> ParsedItems - список объектов CrmGar_fixation, представляющих собой объекты "Финансовые условия" из CRM.
        public static List<CrmGar_fixation> getParsedItemsGar_fixation(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_fixation> DataCollection = getParsedItemsGar_fixation(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_fixation> getParsedItemsGar_fixation(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_fixation> DataCollection = new List<CrmGar_fixation>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_fixation currentNode = new CrmGar_fixation();
                    currentNode.gar_fixationid          = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_fixationid");
                    currentNode.gar_expiration_of       = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_expiration_of");//Окончание действия.
                    currentNode.gar_service_parameter   = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_service_parameter");//Служебный параметр.
                    currentNode.gar_service_parametername   = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_service_parameter", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Задача".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmTask> ParsedItems - список объектов CrmTask, представляющих собой объекты "Задача" из CRM.
        public static List<CrmTask> getParsedItemsTask(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmTask> DataCollection = getParsedItemsTask(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmTask> getParsedItemsTask(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmTask> DataCollection = new List<CrmTask>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmTask currentNode = new CrmTask();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Состояние действия.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.gar_contact             = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contact");//Контакт.
                    currentNode.gar_contactname         = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_contact", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Факс".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmFax> ParsedItems - список объектов CrmFax, представляющих собой объекты "Факс" из CRM.
        public static List<CrmFax> getParsedItemsFax(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmFax> DataCollection = getParsedItemsFax(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmFax> getParsedItemsFax(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmFax> DataCollection = new List<CrmFax>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmFax currentNode = new CrmFax();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Состояние действия.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.gar_contact             = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contact");//Контакт БП.
                    currentNode.gar_contactname         = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_contact", "name");
                    currentNode.from                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "from");//Отправитель.
                    currentNode.to                      = FunctionsToXmlDataWorking.ParseNodeValue(node, "to");//Получатель.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Электронная почта".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmEmail> ParsedItems - список объектов CrmEmail, представляющих собой объекты "Электронная почта" из CRM.
        public static List<CrmEmail> getParsedItemsEmail(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmEmail> DataCollection = getParsedItemsEmail(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmEmail> getParsedItemsEmail(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmEmail> DataCollection = new List<CrmEmail>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmEmail currentNode = new CrmEmail();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Состояние действия.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.from                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "from");//От.
                    currentNode.to                      = FunctionsToXmlDataWorking.ParseNodeValue(node, "to");//Кому.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Письмо".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmLetter> ParsedItems - список объектов CrmLetter, представляющих собой объекты "Письмо" из CRM.
        public static List<CrmLetter> getParsedItemsLetter(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmLetter> DataCollection = getParsedItemsLetter(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmLetter> getParsedItemsLetter(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmLetter> DataCollection = new List<CrmLetter>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmLetter currentNode = new CrmLetter();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Состояние действия.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.from                    = FunctionsToXmlDataWorking.ParseNodeValue(node, "from");//Отправитель.
                    currentNode.to                      = FunctionsToXmlDataWorking.ParseNodeValue(node, "to");//Получатель.
                    currentNode.cc                      = FunctionsToXmlDataWorking.ParseNodeValue(node, "cc");//Список обучаемых.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Встреча".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAppointment> ParsedItems - список объектов CrmAppointment, представляющих собой объекты "Встреча" из CRM.
        public static List<CrmAppointment> getParsedItemsAppointment(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAppointment> DataCollection = getParsedItemsAppointment(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAppointment> getParsedItemsAppointment(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmAppointment> DataCollection = new List<CrmAppointment>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAppointment currentNode = new CrmAppointment();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.regardingobjectid       = FunctionsToXmlDataWorking.ParseNodeValue(node, "regardingobjectid");//В отношении.
                    currentNode.requiredattendees       = FunctionsToXmlDataWorking.ParseNodeValue(node, "requiredattendees");//Обязательные участники.
                    currentNode.optionalattendees       = FunctionsToXmlDataWorking.ParseNodeValue(node, "optionalattendees");//Необязательные участники.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Контракт от кампании".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmCampaignresponse> ParsedItems - список объектов CrmCampaignresponse, представляющих собой объекты "Контракт от кампании" из CRM.
        public static List<CrmCampaignresponse> getParsedItemsCampaignresponse(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmCampaignresponse> DataCollection = getParsedItemsCampaignresponse(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmCampaignresponse> getParsedItemsCampaignresponse(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmCampaignresponse> DataCollection = new List<CrmCampaignresponse>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmCampaignresponse currentNode = new CrmCampaignresponse();
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//GUID.
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode               = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statecodename           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statecode", "name");
                    currentNode.description             = FunctionsToXmlDataWorking.ParseNodeValue(node, "description");//Описание.
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.
                    currentNode.subject                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "subject");//Тема.
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный.
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.gar_contact             = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contact");//Контакт.
                    currentNode.gar_contactname         = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_contact", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Activity Party".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmActivityparty> ParsedItems - список объектов CrmActivityparty, представляющих собой объекты "Activity Party" из CRM.
        public static List<CrmActivityparty> getParsedItemsActivityparty(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmActivityparty> DataCollection = getParsedItemsActivityparty(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmActivityparty> getParsedItemsActivityparty(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmActivityparty> DataCollection = new List<CrmActivityparty>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmActivityparty currentNode = new CrmActivityparty();
                    currentNode.activitypartyid         = FunctionsToXmlDataWorking.ParseNodeValue(node, "activitypartyid");//GUID.
                    currentNode.activityid              = FunctionsToXmlDataWorking.ParseNodeValue(node, "activityid");//Связанное Действие.
                    currentNode.partyid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "partyid");//Связанный участник.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Акт" из XML-выдачи из 1С.
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAct_DocumentFrom1C> ParsedItems - список объектов CrmAct_DocumentFrom1C, представляющих собой объекты "Акт" из XML-выдачи из 1С.
        public static List<CrmAct_DocumentFrom1C> getParsedItemsAct_DocumentFrom1C(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAct_DocumentFrom1C> DataCollection = getParsedItemsAct_DocumentFrom1C(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAct_DocumentFrom1C> getParsedItemsAct_DocumentFrom1C(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("documents/document");   //В варианте для XML Из CRM было "resultset/result".

                List<CrmAct_DocumentFrom1C> DataCollection = new List<CrmAct_DocumentFrom1C>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAct_DocumentFrom1C currentNode = new CrmAct_DocumentFrom1C();
                    currentNode.datetime = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "datetime");
                    currentNode.number = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "number");
                    currentNode.ismark = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "ismark");//Удален.
                    currentNode.closed = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "closed");//Закрыт.
                    currentNode.returned = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "returned");
                    currentNode.returndate = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "returndate");
                    currentNode.createdate = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "createdate");
                    currentNode.maintenance_period_start = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "maintenance_period_start");
                    currentNode.maintenance_period_end = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "maintenance_period_end");
                    currentNode.uslugi_sum = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "uslugi_sum");
                    currentNode.uslugi_sumNDS = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "uslugi_sumNDS");
                    currentNode.integral = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node, "integral");
                    currentNode.comment = FunctionsToXmlDataWorking.ParseNodeValue(node, "comment");
                    currentNode.manager = FunctionsToXmlDataWorking.ParseNodeValue(node, "manager");
                    currentNode.managerdescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "manager", "description");
                    currentNode.managercode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "manager", "code");
                    currentNode.author = FunctionsToXmlDataWorking.ParseNodeValue(node, "author");
                    currentNode.authordescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "author", "description");
                    currentNode.authorcode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "author", "code");
                    currentNode.kontragent = FunctionsToXmlDataWorking.ParseNodeValue(node, "kontragent");
                    currentNode.kontragentdescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "kontragent", "description");
                    currentNode.kontragentcode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "kontragent", "code");
                    currentNode.dogovor = FunctionsToXmlDataWorking.ParseNodeValue(node, "dogovor");
                    currentNode.dogovordescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "dogovor", "description");
                    currentNode.dogovorcode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "dogovor", "code");
                    currentNode.organization = FunctionsToXmlDataWorking.ParseNodeValue(node, "organization");
                    currentNode.organizationdescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "organization", "description");
                    currentNode.organizationcode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "organization", "code");
                    currentNode.usluga = FunctionsToXmlDataWorking.ParseNodeValue(node, "usluga");
                    currentNode.pay = FunctionsToXmlDataWorking.ParseNodeValue(node, "pay");

                    XmlNodeList Itemslist2 = node.SelectNodes("uslugi/usluga");
                    List<CrmUsluga_DocumentFrom1C> DataCollection2 = new List<CrmUsluga_DocumentFrom1C>();
                    foreach (XmlNode node2 in Itemslist2)
                    {
                        CrmUsluga_DocumentFrom1C currentNode2 = new CrmUsluga_DocumentFrom1C();
                        currentNode2.sim = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "sim");
                        currentNode2.lineno = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "lineno");
                        currentNode2.amount = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "amount");
                        currentNode2.sumNDS = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "sumNDS");
                        currentNode2.NDS = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "NDS");
                        currentNode2.sum = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "sum");
                        currentNode2.price = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node2, "price");
                        currentNode2.actiondescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node2, "action", "description");
                        currentNode2.nomenklature = FunctionsToXmlDataWorking.ParseNodeValue(node2, "nomenklature");
                        currentNode2.nomenklaturedescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node2, "nomenklature", "description");
                        currentNode2.nomenklaturecode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node2, "nomenklature", "code");
                        currentNode2.characteristic = FunctionsToXmlDataWorking.ParseNodeValue(node2, "characteristic");
                        currentNode2.characteristicdescription = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node2, "characteristic", "description");
                        currentNode2.characteristiccode = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node2, "characteristic", "code");
                        DataCollection2.Add(currentNode2);
                    }
                    currentNode.listUsluga = DataCollection2;

                    XmlNodeList Itemslist3 = node.SelectNodes("scheme/pay");
                    List<CrmScheme_DocumentFrom1C> DataCollection3 = new List<CrmScheme_DocumentFrom1C>();
                    foreach (XmlNode node3 in Itemslist3)
                    {
                        CrmScheme_DocumentFrom1C currentNode3 = new CrmScheme_DocumentFrom1C();
                        currentNode3.lineno = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "lineno");
                        currentNode3.characteristiccode = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "characteristiccode");
                        currentNode3.nomenklaturecode = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "nomenklaturecode");
                        currentNode3.actiondescr = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "actiondescr");
                        currentNode3.periodstart = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "periodstart");
                        currentNode3.periodend = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "periodend");
                        currentNode3.sum = FunctionsToXmlDataWorking.ParseItselfAttributeValue(node3, "sum");
                        DataCollection3.Add(currentNode3);
                    }
                    currentNode.listScheme = DataCollection3;

                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Акт".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_act> ParsedItems - список объектов CrmGar_act, представляющих собой объекты "Акт" из CRM.
        public static List<CrmGar_act> getParsedItemsGar_act(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_act> DataCollection = getParsedItemsGar_act(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_act> getParsedItemsGar_act(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_act> DataCollection = new List<CrmGar_act>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_act currentNode = new CrmGar_act();
                    currentNode.gar_actid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_actid");
                    currentNode.gar_datetime = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_datetime");//От.
                    currentNode.gar_number = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_number");//Номер.
                    currentNode.gar_kontragent = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_kontragent");//Контрагент.
                    currentNode.gar_organization = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_organization");//Организация.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Услуга".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_service> ParsedItems - список объектов CrmGar_service, представляющих собой объекты "Услуга" из CRM.
        public static List<CrmGar_service> getParsedItemsGar_service(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_service> DataCollection = getParsedItemsGar_service(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_service> getParsedItemsGar_service(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_service> DataCollection = new List<CrmGar_service>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_service currentNode = new CrmGar_service();
                    currentNode.gar_serviceid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_serviceid");
                    currentNode.gar_actid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_actid");//Акт.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Схема".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_scheme> ParsedItems - список объектов CrmGar_scheme, представляющих собой объекты "Схема" из CRM.
        public static List<CrmGar_scheme> getParsedItemsGar_scheme(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_scheme> DataCollection = getParsedItemsGar_scheme(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_scheme> getParsedItemsGar_scheme(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_scheme> DataCollection = new List<CrmGar_scheme>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_scheme currentNode = new CrmGar_scheme();
                    currentNode.gar_schemeid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_schemeid");
                    currentNode.gar_actid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_actid");//Акт.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Разнесение оплаты".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmGar_paysplit> ParsedItems - список объектов CrmGar_paysplit, представляющих собой объекты "Разнесение оплаты" из CRM.
        public static List<CrmGar_paysplit> getParsedItemsGar_paysplit(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmGar_paysplit> DataCollection = getParsedItemsGar_paysplit(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmGar_paysplit> getParsedItemsGar_paysplit(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmGar_paysplit> DataCollection = new List<CrmGar_paysplit>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmGar_paysplit currentNode = new CrmGar_paysplit();
                    currentNode.gar_paysplitid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_paysplitid");
                    currentNode.gar_payid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_payid");//Оплата.
                    currentNode.gar_payidname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_payid", "name");
                    currentNode.gar_paytype = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_paytype");//Тип оплаты.
                    currentNode.gar_paytypename = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_paytype", "name");
                    currentNode.gar_month = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_month");//Месяц.
                    currentNode.gar_monthname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_month", "name");
                    currentNode.gar_year = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_year");//Год.
                    currentNode.gar_yearname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_year", "name");
                    currentNode.gar_amount = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_amount");//Сумма.
                    currentNode.gar_contract = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contract");//Контракт.
                    currentNode.gar_contractname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_contract", "name");
                    currentNode.gar_accountid = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_accountid");//Бизнес-партнер.
                    currentNode.gar_accountidname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_accountid", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Объект неопределенного заранее типа".
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmIndeterminatedObject> ParsedItems - список объектов CrmIndeterminatedObject, представляющих собой объекты "Объект неопределенного заранее типа" из CRM.
        public static List<CrmIndeterminatedObject> getParsedItemsIndeterminatedObject(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmIndeterminatedObject> DataCollection = getParsedItemsIndeterminatedObject(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmIndeterminatedObject> getParsedItemsIndeterminatedObject(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmIndeterminatedObject> DataCollection = new List<CrmIndeterminatedObject>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmIndeterminatedObject currentNode = new CrmIndeterminatedObject();
                    currentNode.accountcategorycode     = FunctionsToXmlDataWorking.ParseNodeValue(node, "accountcategorycode");//Категория.                /"Бизнес-партнер".
                    currentNode.accountcategorycodename = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "accountcategorycode", "name");
                    currentNode.gar_arhsb               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_arhsb");//Основной день визита.     /"Бизнес-партнер".
                    currentNode.gar_arhsbname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_arhsb", "name");
                    currentNode.preferredappointmentdaycode     = FunctionsToXmlDataWorking.ParseNodeValue(node, "preferredappointmentdaycode");//Основной день обновления. /"Бизнес-партнер".
                    currentNode.preferredappointmentdaycodename = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "preferredappointmentdaycode", "name");
                    currentNode.ownerid                 = FunctionsToXmlDataWorking.ParseNodeValue(node, "ownerid");//Ответственный./"Бизнес-партнер". //Ответственный./"Контакт". //Ответственный №1./"Обращение".
                    currentNode.owneridname             = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "ownerid", "name");
                    currentNode.businesstypecode        = FunctionsToXmlDataWorking.ParseNodeValue(node, "businesstypecode");//Режим налогообложения.    /"Бизнес-партнер".
                    currentNode.businesstypecodename    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "businesstypecode", "name");
                    currentNode.statuscode              = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.                /"Бизнес-партнер".
                    currentNode.statuscodename          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.createdby               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdby");//Создано.                  /"Бизнес-партнер".
                    currentNode.createdbyname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "createdby", "name");
                    currentNode.createdon               = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.            /"Бизнес-партнер".
                    currentNode.modifiedby              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedby");//Изменено.                 /"Бизнес-партнер".
                    currentNode.modifiedbyname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "modifiedby", "name");
                    currentNode.modifiedon              = FunctionsToXmlDataWorking.ParseNodeValue(node, "modifiedon");//Дата изменения.           /"Бизнес-партнер".
                    currentNode.gar_channel_appearance  = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_channel_appearance");//Источник появления.       /"Бизнес-партнер".
                    currentNode.gar_channel_appearancename      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_channel_appearance", "name");

                    currentNode.preferredcontactmethodcode      = FunctionsToXmlDataWorking.ParseNodeValue(node, "preferredcontactmethodcode");//Основной способ связи.    /"Контакт".
                    currentNode.preferredcontactmethodcodename  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "preferredcontactmethodcode", "name");
                    currentNode.accountrolecode         = FunctionsToXmlDataWorking.ParseNodeValue(node, "accountrolecode");//Роль 1.                   /"Контакт".
                    currentNode.accountrolecodename     = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "accountrolecode", "name");
                    currentNode.customertypecode        = FunctionsToXmlDataWorking.ParseNodeValue(node, "customertypecode");//Тип отношений.            /"Контакт".
                    currentNode.customertypecodename    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customertypecode", "name");

                    currentNode.caseorigincode          = FunctionsToXmlDataWorking.ParseNodeValue(node, "caseorigincode");//Канал.                    /"Обращение".
                    currentNode.caseorigincodename      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "caseorigincode", "name");
                    currentNode.subjectid               = FunctionsToXmlDataWorking.ParseNodeValue(node, "subjectid");//Тема.                     /"Обращение".
                    currentNode.subjectidname           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "subjectid", "name");
                    currentNode.casetypecode            = FunctionsToXmlDataWorking.ParseNodeValue(node, "casetypecode");//Тип.                      /"Обращение".
                    currentNode.casetypecodename        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "casetypecode", "name");

                    currentNode.gar_info_povod_lead     = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_info_povod_lead");//Информационный повод.     /"Интерес".
                    currentNode.gar_info_povod_leadname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_info_povod_lead", "name");
                    currentNode.leadsourcecode          = FunctionsToXmlDataWorking.ParseNodeValue(node, "leadsourcecode");//Источник.                 /"Интерес".
                    currentNode.leadsourcecodename      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "leadsourcecode", "name");
                    currentNode.gar_reference_legal_system      = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_reference_legal_system");//СПС.                      /"Интерес".
                    currentNode.gar_reference_legal_systemname  = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_reference_legal_system", "name");
                    currentNode.gar_result_tmc          = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_result_tmc");//Результат работы ТМЦ.     /"Интерес".
                    currentNode.gar_result_tmcname      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_result_tmc", "name");

                    currentNode.gar_businessunit        = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_businessunit");//Отдел.                    /"История работы".
                    currentNode.gar_businessunitname    = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_businessunit", "name");
                    currentNode.gar_systemuser          = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_systemuser");//Сотрудник.                /"История работы".
                    currentNode.gar_systemusername      = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_systemuser", "name");
                    currentNode.gar_projects            = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_projects");//Проект.                   /"История работы".
                    currentNode.gar_projectsname        = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_projects", "name");
                    currentNode.gar_result              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_result");//Результат.                /"История работы".
                    currentNode.gar_resultname          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_result", "name");
                    currentNode.gar_list                = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_list");//Маркетинговый список.     /"История работы".
                    currentNode.gar_listname            = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_list", "name");

                    currentNode.gar_item_1              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item_1");//Знаете как зовут обслуживающего Вас сотрудника?       /"Действие сервиса".
                    currentNode.gar_item_1name          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_item_1", "name");
                    currentNode.gar_item3               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item3");//Как Вы оцениваете его профессиональный уровень?       /"Действие сервиса".
                    currentNode.gar_item3name           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_item3", "name");
                    currentNode.gar_item17              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item17");//Как Вы оцениваете качество обслуживания?              /"Действие сервиса".
                    currentNode.gar_item17name          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_item17", "name");
                    currentNode.gar_item9               = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item9");//Как часто в работе Вы используете ИПО ГАРАНТ?         /"Действие сервиса".
                    currentNode.gar_item9name           = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_item9", "name");
                    currentNode.gar_item16              = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_item16");//Как планируете получать правовую информацию далее?    /"Действие сервиса".
                    currentNode.gar_item16name          = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_item16", "name");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        //Функция разбирает объекты в Xml-документе в список объектов неизвестного заранее типа (без параметров).
        //Входные параметры:
        //у перегрузки 1: string Url - путь к Xml-файлу, откуда считывается Xml-документ, с именем которого вызывается перегрузка 2.
        //у перегрузки 2: XmlDocument xmlDoc - имя Xml-документа.
        //Выходные параметры:
        //List<CrmAnyObjectWithNoParameters> ParsedItems - список объектов CrmAnyObjectWithNoParameters, представляющих собой объекты неизвестного заранее типа (без параметров) из CRM.
        public static List<CrmAnyObjectWithNoParameters> getParsedItemsAnyObjectWithNoParameters(string Url)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Url);
                List<CrmAnyObjectWithNoParameters> DataCollection = getParsedItemsAnyObjectWithNoParameters(xmlDoc);
                return DataCollection;
            }
            catch { return null; }
        }
        public static List<CrmAnyObjectWithNoParameters> getParsedItemsAnyObjectWithNoParameters(XmlDocument xmlDoc)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmAnyObjectWithNoParameters> DataCollection = new List<CrmAnyObjectWithNoParameters>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmAnyObjectWithNoParameters currentNode = new CrmAnyObjectWithNoParameters();
                    //Нет обрабатываемых параметров для объектов этого типа.
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }
    }

    public class FunctionsToStringWorking
    {
        //Функция возвращает строку, представляющую собой номер телефона в правильном формате (со скобками и дефисами) в зависимости от количества цифр.
        //Входные параметры:
        //string inputedPhone - исходная строка, представляющая собой неформатированный номер (любую совокупность цифровых и нецифровых символов).
        //Выходные параметры:
        //string formatPhone - обработанная строка.
        public static string formatPhone(string inputedPhone)
        {
            string inputedPhoneDigitsOnly = Regex.Replace(inputedPhone, @"[^0-9]", String.Empty);   //Требует using System.Text.RegularExpressions;
            string newString = "";
            switch (inputedPhoneDigitsOnly.Length)
            {
                case 11:
                    if (inputedPhoneDigitsOnly[1] == '9')
                    {
                        newString = "+7-" + inputedPhoneDigitsOnly.Substring(1, 3) + "-" +
                            inputedPhoneDigitsOnly.Substring(4, 3) + "-" + inputedPhoneDigitsOnly.Substring(7, 2) + "-" + inputedPhoneDigitsOnly.Substring(9, 2);
                    }
                    else
                    {
                        newString = inputedPhoneDigitsOnly.Substring(0, 1) + " (" + inputedPhoneDigitsOnly.Substring(1, 4) + ") " +
                            inputedPhoneDigitsOnly.Substring(5, 2) + "-" + inputedPhoneDigitsOnly.Substring(7, 2) + "-" + inputedPhoneDigitsOnly.Substring(9, 2);
                    }
                    break;

                case 10:
                    if (inputedPhoneDigitsOnly[0] == '9')
                    {
                        newString = "+7-" + inputedPhoneDigitsOnly.Substring(0, 3) + "-" +
                            inputedPhoneDigitsOnly.Substring(3, 3) + "-" + inputedPhoneDigitsOnly.Substring(6, 2) + "-" + inputedPhoneDigitsOnly.Substring(8, 2);
                    }
                    else
                    {
                        newString = "8 (" + inputedPhoneDigitsOnly.Substring(0, 4) + ") " +
                            inputedPhoneDigitsOnly.Substring(4, 2) + "-" + inputedPhoneDigitsOnly.Substring(6, 2) + "-" + inputedPhoneDigitsOnly.Substring(8, 2);
                    }
                    break;

                case 7:
                    newString = inputedPhoneDigitsOnly.Substring(0, 3) + "-" +
                        inputedPhoneDigitsOnly.Substring(3, 2) + "-" + inputedPhoneDigitsOnly.Substring(5, 2);
                    break;

                case 6:
                    newString = inputedPhoneDigitsOnly.Substring(0, 2) + "-" +
                        inputedPhoneDigitsOnly.Substring(2, 2) + "-" + inputedPhoneDigitsOnly.Substring(4, 2);
                    break;

                case 5:
                    newString = inputedPhoneDigitsOnly.Substring(0, 1) + "-" +
                        inputedPhoneDigitsOnly.Substring(1, 2) + "-" + inputedPhoneDigitsOnly.Substring(3, 2);
                    break;

                case 4:
                    newString = inputedPhoneDigitsOnly.Substring(0, 2) + "-" + inputedPhoneDigitsOnly.Substring(2, 2);
                    break;

                default:
                    newString = inputedPhoneDigitsOnly;
                    break;
            }
            return newString;
        }

        //Функция возвращает дату в формате, подходящем для использования в фетч-запросе, т. е. в формате строки вида "2012-06-01".
        //Входные параметры:
        //DateTime datetimeDate - дата в формате DateTime, из которой получаем день, месяц, год для преобразования в нужный формат.
        //Выходные параметры:
        //string getDateFormattedToUsingInTheFetchRequest - дата в требуемом формате.
        public static string getDateFormattedToUsingInTheFetchRequest(DateTime datetimeDate)
        {
            int dDay = datetimeDate.Day;
            int dMon = datetimeDate.Month;
            int dYar = datetimeDate.Year;
            string requiredDay = dDay.ToString();
            if (dDay < 10)
            {
                requiredDay = "0" + requiredDay;
            }
            string requiredMon = dMon.ToString();
            if (dMon < 10)
            {
                requiredMon = "0" + requiredMon;
            }
            string requiredYar = dYar.ToString();
            string requiredString = requiredYar + "-" + requiredMon + "-" + requiredDay;
            return requiredString;
        }

        //Функция возвращает дату и время в формате, подходящем для удобочитаемого вывода, т. е. в формате строки вида "2015.01.27-15.43.28".
        //Входные параметры:
        //DateTime datetimeDate - дата в формате DateTime, из которой получаем день, месяц, год, часы, минуты, секунды для преобразования в нужный формат.
        //Выходные параметры:
        //string getDateTimeFormattedToShowToTheUser - дата в требуемом формате.
        public static string getDateTimeFormattedToShowToTheUser(DateTime datetimeDate)
        {
            int dYar = datetimeDate.Year;
            int dMon = datetimeDate.Month;
            int dDay = datetimeDate.Day;
            int dHou = datetimeDate.Hour;
            int dMin = datetimeDate.Minute;
            int dSec = datetimeDate.Second;
            string requiredYar = dYar.ToString();
            string requiredMon = dMon.ToString();
            if (dMon < 10)
            {
                requiredMon = "0" + requiredMon;
            }
            string requiredDay = dDay.ToString();
            if (dDay < 10)
            {
                requiredDay = "0" + requiredDay;
            }
            string requiredHou = dHou.ToString();
            if (dHou < 10)
            {
                requiredHou = "0" + requiredHou;
            }
            string requiredMin = dMin.ToString();
            if (dMin < 10)
            {
                requiredMin = "0" + requiredMin;
            }
            string requiredSec = dSec.ToString();
            if (dSec < 10)
            {
                requiredSec = "0" + requiredSec;
            }
            string requiredString = requiredYar + "." + requiredMon + "." + requiredDay + "-" + requiredHou + "." + requiredMin + "." + requiredSec;
            return requiredString;
        }

        //Функция возвращает строку, где первые буквы каждого слова - верхнего регистра, остальные - нижнего.
        //Входные параметры:
        //string inputedFIO - исходная строка.
        //Выходные параметры:
        //string getStringAsFIO - обработанная строка.
        public static string getStringAsFIO(string inputedFIO)
        {
            string returnFIO = "";  //Изначально пустая строка для вывода.
            bool beforeCurrencyLetterWasSpaceOrDot = true;
            for (int iii = 0; iii <= inputedFIO.Length - 1; iii++)
            {
                char ch = inputedFIO[iii];   //Получить текущий символ из строки.
                if (beforeCurrencyLetterWasSpaceOrDot)   //Если перед текущим стоял пробел/точка, сделать текущий большим.
                {
                    ch = char.ToUpper(ch);
                }
                else   //Если перед текущим НЕ стоял пробел, сделать текущий маленьким.
                {
                    ch = char.ToLower(ch);
                }
                returnFIO = returnFIO + ch;    //Добавить обработанный символ к строке для вывода.

                if ((ch == ' ') || (ch == '.'))  //Запомнить для следующей итерации, пробел/точка это или нет.
                {
                    beforeCurrencyLetterWasSpaceOrDot = true;
                }
                else
                {
                    beforeCurrencyLetterWasSpaceOrDot = false;
                }
            }
            return returnFIO;
        }

        //Функция возвращает ответ на вопрос, является ли символ буквой английского алфавита (строчной или прописной) или нет.
        //Входные параметры:
        //char letter - исходный символ.
        //Выходные параметры:
        //bool isEnglishLetter - true, если исходный символ является буквой английского алфавита (строчной или прописной), иначе - false.
        public static bool isEnglishLetter(char letter)
        {
            bool answer = false;
            switch (letter)
            {
                case 'A':
                case 'B':
                case 'C':
                case 'D':
                case 'E':
                case 'F':
                case 'G':
                case 'H':
                case 'I':
                case 'J':
                case 'K':
                case 'L':
                case 'M':
                case 'N':
                case 'O':
                case 'P':
                case 'Q':
                case 'R':
                case 'S':
                case 'T':
                case 'U':
                case 'V':
                case 'W':
                case 'X':
                case 'Y':
                case 'Z':
                case 'a':
                case 'b':
                case 'c':
                case 'd':
                case 'e':
                case 'f':
                case 'g':
                case 'h':
                case 'i':
                case 'j':
                case 'k':
                case 'l':
                case 'm':
                case 'n':
                case 'o':
                case 'p':
                case 'q':
                case 'r':
                case 's':
                case 't':
                case 'u':
                case 'v':
                case 'w':
                case 'x':
                case 'y':
                case 'z':
                    answer = true;
                    break;

                default:
                    answer = false;
                    break;
            }
            return answer;
        }
    }

    public class FunctionsToAnyDataWorking
    {
        //Функция получает порядковый номер заданного дня с 01.01.2000 г.
        //Входные параметры:
        //у перегрузки 1: string stringDate - дата в формате строки вида "2012-06-01", которая преобразуется в DateTime, с которым вызывается перегрузка 2.
        //у перегрузки 2: DateTime datetimeDate - дата в формате DateTime, из которой получаются int day, month, year, с которыми вызывается перегрузка 3.
        //у перегрузки 3:
        //int day - заданный день (день заданной даты).
        //int month - заданный месяц (месяц заданной даты).
        //int year - заданный год (год заданной даты).
        //Выходные параметры:
        //int numberOfDaySince01012000 - порядковый номер дня. Если не удалось получить значение, то 0.
        public static int numberOfDaySince01012000(string stringDate)
        {
            DateTime wantedDate = DateTime.Parse(stringDate);
            int wantedNumber = numberOfDaySince01012000(wantedDate);
            return wantedNumber;
        }
        public static int numberOfDaySince01012000(DateTime datetimeDate)
        {
            int wantedDay       = datetimeDate.Day;
            int wantedMonth     = datetimeDate.Month;
            int wantedYear      = datetimeDate.Year;
            int wantedNumber = numberOfDaySince01012000(wantedDay, wantedMonth, wantedYear);
            return wantedNumber;
        }
        public static int numberOfDaySince01012000(int day, int month, int year)
        {
            if ((year < 2000) || (day < 1) || (day > 31) || (month < 1) || (month > 12))
            {
                return 0;
            }
            else
            {
                int number = 0;
                //Считаем дни в полных годах с 2000 года:
                for (int iii = 2000; iii < year; iii++)
                {
                    if (DateTime.IsLeapYear(iii))   //Если високосный.
                    {
                        number = number + 366;
                    }
                    else
                    {
                        number = number + 365;
                    }
                }
                //Считаем дни в полных месяцах с 01.01.этого года:
                for (int iii = 1; iii < month; iii++)
                {
                    switch (iii)
                    {
                        case 1:
                        case 3:
                        case 5:
                        case 7:
                        case 8:
                        case 10:
                        case 12:
                            number = number + 31;
                            break;
                        case 4:
                        case 6:
                        case 9:
                        case 11:
                            number = number + 30;
                            break;
                        case 2:
                            number = number + 28;
                            if (DateTime.IsLeapYear(year))  //Если текущий год - високосный.
                            {
                                number = number + 1;
                            }
                            break;
                    }
                }
                //Считаем дни в текущем месяце:
                number = number + day;
                //Возвращаем результат:
                return number;
            }
        }
        
        //Процедура удаляет объект и освобождает в памяти место из-под него.
        //Входные параметры:
        //object obj - удаляемый объект.
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.Write("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }

    public class FunctionsToXmlDataWorking
    {
        //Функция получает значение элемента (свойства) из записи (узла) Xml.
        //Входные параметры:
        //XmlNode item - запись (узел) Xml.
        //string propertyName - имя элемента (свойства) из записи (узла) Xml, значение которого необходимо получить.
        //Выходные параметры:
        //string ParseNodeValue - значение элемента (свойства) с заданным именем из заданной записи (узла) Xml.
        public static string ParseNodeValue(XmlNode item, string propertyName)
        {
            XmlNode propertyNode = item.SelectSingleNode(propertyName);
            string value = string.Empty;
            if (propertyNode != null) value = propertyNode.InnerText;
            return value;
        }

        //Функция получает значение атрибута элемента (свойства) из записи (узла) Xml.
        //Входные параметры:
        //XmlNode item - запись (узел) Xml.
        //string propertyName - имя элемента (свойства) из записи (узла) Xml, значение атрибута которого необходимо получить.
        //string attributeName - имя атрибута, значение которого необходимо получить.
        //Выходные параметры:
        //string ParseNodeAttributeValue - значение атрибута с заданным именем из заданного элемента (свойства) из заданной записи (узла) Xml.
        public static string ParseNodeAttributeValue(XmlNode item, string propertyName, string attributeName)
        {
            XmlNode propertyNode = item.SelectSingleNode(propertyName);
            string value = string.Empty;
            if (propertyNode != null) value = propertyNode.Attributes[attributeName].Value;
            return value;
        }
    }

    public class FunctionsToXmlDataWorking_ForWorkWith1S
    {
        //Функция получает значение атрибута из записи (узла) Xml.
        //Входные параметры:
        //XmlNode item - запись (узел) Xml.
        //string attributeName - имя атрибута, значение которого необходимо получить.
        //Выходные параметры:
        //string ParseItselfAttributeValue - значение атрибута с заданным именем из заданной записи (узла) Xml.
        public static string ParseItselfAttributeValue(XmlNode item, string attributeName)
        {
            XmlNode propertyNode = item;    //Узел - этот же, остальное по шаблону.
            string value = string.Empty;
            if (propertyNode != null) value = propertyNode.Attributes[attributeName].Value;
            return value;
        }

        //Функция получает значение атрибута элемента (свойства) из записи (узла) Xml.
        //Входные параметры:
        //XmlNode item - запись (узел) Xml.
        //string propertyName - имя элемента (свойства) из записи (узла) Xml, значение атрибута которого необходимо получить.
        //string attributeName - имя атрибута, значение которого необходимо получить.
        //Выходные параметры:
        //string ParseNodeAttributeValue - значение атрибута с заданным именем из заданного элемента (свойства) из заданной записи (узла) Xml.
        public static string ParseNodeAttributeValue(XmlNode item, string propertyName, string attributeName)
        {
            XmlNode propertyNode = item.SelectSingleNode(propertyName);
            string value = string.Empty;
            if (propertyNode != null) value = propertyNode.Attributes[attributeName].Value;
            return value;
        }
    }

    public class FunctionsToCrmDataWorking
    {
        //Функция получает экземпляр указанной сущности с указанным набором атрибутов по GUID с использованием динамического запроса.
        //Входные параметры:
        //у перегрузки 1:
        //ICrmService service - сервис для работы с базой CRM, из которой получить искомый экземпляр сущности.
        //string entityname - имя сущности.
        //string entityidStr - строковое представление GUID искомого экземпляра сущности, которое преобразуется в Guid entityid, с которым вызывается перегрузка 2.
        //ColumnSet entitycolumnset - набор атрибутов, которые требуется вернуть в найденном экземпляре.
        //у перегрузки 2:
        //ICrmService service - сервис для работы с базой CRM, из которой получить искомый экземпляр сущности.
        //string entityname - имя сущности.
        //Guid entityid - GUID искомого экземпляра сущности.
        //ColumnSet entitycolumnset - набор атрибутов, которые требуется вернуть в найденном экземпляре.
        //Выходные параметры:
        //DynamicEntity getInstanceOfDynamicEntityById - искомый экземпляр сущности.
        public DynamicEntity getInstanceOfDynamicEntityById(ICrmService service, string entityname, string entityidStr, ColumnSet entitycolumnset)
        {
            Guid entityid = new Guid(entityidStr);
            DynamicEntity wantedEntity = getInstanceOfDynamicEntityById(service, entityname, entityid, entitycolumnset);
            return wantedEntity;
        }
        public DynamicEntity getInstanceOfDynamicEntityById(ICrmService service, string entityname, Guid entityid, ColumnSet entitycolumnset)
        {
            TargetRetrieveDynamic tr = new TargetRetrieveDynamic();
            tr.EntityName = entityname;
            tr.EntityId = entityid;
            RetrieveRequest rr = new RetrieveRequest();
            rr.ColumnSet = entitycolumnset;
            rr.ReturnDynamicEntities = true;
            rr.Target = tr;
            RetrieveResponse resp = (RetrieveResponse)service.Execute(rr);
            return (DynamicEntity)resp.BusinessEntity;
        }
        
        //Функция создает объект CrmService стандартным способом из справки CRM SDK.
        //Входные параметры:
        //string organizationName - имя организации (у нас: тестовая база - "CmpnyLab", рабочая база - "Cmpny").
        //string serviceUrl - путь к CRM-сервису (у нас: "http://crm4/mscrmservices/2007/crmservice.asmx").
        //Выходные параметры:
        //CrmService createCrmService - полученный сервис для работы с базой CRM.
        public static CrmService createCrmService(string organizationName, string serviceUrl)
        {
            CrmAuthenticationToken token = new CrmAuthenticationToken();
            token.AuthenticationType = 0;
            token.OrganizationName = organizationName;  //У нас - имя организации: тестовая база - "CmpnyLab", рабочая база - "Cmpny".
            CrmService newService = new CrmService();
            newService.Url = serviceUrl;
            newService.CrmAuthenticationTokenValue = token;
            newService.Credentials = System.Net.CredentialCache.DefaultCredentials;
            return newService;
        }

        //Функция создает объект MetadataService стандартным способом из справки CRM SDK.
        //Входные параметры:
        //string organizationName - имя организации (у нас: тестовая база - "CmpnyLab", рабочая база - "Cmpny").
        //string serviceUrl - путь к Metadata-сервису (у нас: "http://crm4/mscrmservices/2007/metadataservice.asmx").
        //Выходные параметры:
        //MetadataService createMetadataService - полученный сервис для работы с метаданными CRM.
        public static MetadataService createMetadataService(string organizationName, string serviceUrl)
        {
            CrmAuthenticationToken token = new CrmAuthenticationToken();
            token.AuthenticationType = 0;
            token.OrganizationName = organizationName;  //У нас - имя организации: тестовая база - "CmpnyLab", рабочая база - "Cmpny".
            MetadataService newService = new MetadataService();
            newService.Url = serviceUrl;
            newService.CrmAuthenticationTokenValue = token;
            newService.Credentials = System.Net.CredentialCache.DefaultCredentials;
            newService.PreAuthenticate = true;
            return newService;
        }

        //Функция возвращает все записи Fetch-запроса, даже если их больше чем значение верхнего предела (5000 по умолчанию).
        //В этом ее отличие от стандартной функции fetch, которая возвращает не более определенного количества значений (5000 по умолчанию).
        //Входные параметры:
        //string sFetchXml - строка fetch-запроса.
        //CrmService sCrmService - сервис для работы с базой CRM, в котором искать записи по fetch-запросу.
        //Выходные параметры:
        //string fetchAll - результат fetch-запроса.
        public static string fetchAll(string sFetchXml, CrmService sCrmService)
        {
            bool bComplete = false; //Переменная, опеределяющая необходимость продолжать цикл.
            int iPage = 1;          //Начинаем просмотр с первой страницы.
            XmlDocument oResults = new XmlDocument();   //Общий массив, в коториый будем помещать результаты временных выборок.
            while (!bComplete)                          //Повторяем постраничную выборку, пока bComplete равен false.
            {
                XmlDocument oFetchXml = new XmlDocument();  //XML документ для "разбора" строки Fetch-запроса.
                oFetchXml.LoadXml(sFetchXml);               //Загружаем строку Fetch-запроса в XML документ.
                XmlNode oFetchNode = oFetchXml.SelectSingleNode("/fetch");  //Находим в XML документе узел fetch.
                
                //Если в узле fetch нет атрибута, то создаем его. В любом случае помещаем в этот атрибут номер текущей страницы (итерации):
                if (oFetchNode.Attributes["page"] == null)
                {
                    XmlAttribute oPageAttribute = oFetchXml.CreateAttribute("page");
                    oPageAttribute.Value = iPage.ToString();
                    oFetchNode.Attributes.Append(oPageAttribute);
                }
                else
                {
                    oFetchNode.Attributes["page"].Value = iPage.ToString();
                }

                string sResultsXml = sCrmService.Fetch(oFetchXml.InnerXml); //Выполняем Fetch-запрос.
                XmlDocument oTempXml = new XmlDocument();           //Создаем временный массив, в который мы помещаем...
                oTempXml.LoadXml(sResultsXml);                      //...до 5000 записей за одну итерацию.

                //Затем мы добавляем результаты временной выборки в собирательный XML объект oResults:
                if (iPage == 1)
                {
                    oResults.LoadXml(oTempXml.InnerXml);
                }
                else
                {
                    XmlNodeList oNodes = oTempXml.SelectNodes("/resultset/result");
                    XmlNode oResultsetNode = oResults.SelectSingleNode("/resultset");
                    foreach (XmlNode oNode in oNodes)
                    {
                        XmlNode oNewNode = oResults.ImportNode(oNode, true);
                        oResultsetNode.AppendChild(oNewNode);
                    }
                }

                //Теперь мы должны выяснить, есть ли еще страницы для выборки.
                //Для этого мы проверяем атрибут morerecords узла resultset.
                //Если такие страницы есть, мы только инкриминируем iPage, иначе устанавливаем bComplete в true, чем прерываем цикл.
                XmlNode oMore = oTempXml.SelectSingleNode("/resultset");
                if (oMore.Attributes["morerecords"].Value == "0")
                {
                    bComplete = true;
                }
                else
                {
                    iPage++;
                }
            }
            return oResults.InnerXml;   //Возвращаем содержимое собирательного массива oResults.
        }

        //Комментарий неотлажен:
        //Получить значение атрибута типа "статус" по его коду из метаданных:
        private string getStatusPropertyTextFromMetadata(IMetadataService service, int value, string entityLogicalName, string propertyLogicalName)
        {
            RetrieveAttributeRequest req = new RetrieveAttributeRequest();
            req.EntityLogicalName = entityLogicalName;
            req.LogicalName = propertyLogicalName;
            RetrieveAttributeResponse resp = (RetrieveAttributeResponse)service.Execute(req);
            StatusAttributeMetadata met = (StatusAttributeMetadata)resp.AttributeMetadata;
            foreach (Option o in met.Options)
            {
                if (o.Value.Value == value)
                {
                    return o.Label.UserLocLabel.Label;
                }
            }
            return "";
        }

        //Комментарий неотлажен:
        //Получить значение атрибута типа "пиклист" по его коду из метаданных:
        private string getPicklistPropertyTextFromMetadata(IMetadataService service, int value, string entityLogicalName, string propertyLogicalName)
        {
            RetrieveAttributeRequest req = new RetrieveAttributeRequest();
            req.EntityLogicalName = entityLogicalName;
            req.LogicalName = propertyLogicalName;
            RetrieveAttributeResponse resp = (RetrieveAttributeResponse)service.Execute(req);
            PicklistAttributeMetadata met = (PicklistAttributeMetadata)resp.AttributeMetadata;
            foreach (Option o in met.Options)
            {
                if (o.Value.Value == value)
                {
                    return o.Label.UserLocLabel.Label;
                }
            }
            return "";
        }
    }
}
