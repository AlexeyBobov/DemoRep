using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//Добавлено вручную:
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using System.Xml.Serialization;                 //Добавлен для разбора Xml-файла.
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Crm.Sdk;                        //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.Sdk.Metadata;               //Файл microsoft.crm.sdk.dll.
using Microsoft.Crm.SdkTypeProxy;               //Файл microsoft.crm.sdktypeproxy.dll.
using Microsoft.Crm.SdkTypeProxy.Metadata;      //Файл microsoft.crm.sdktypeproxy.dll.

namespace Data_and_metadata_getting_tool_002
{
    public partial class Form1 : Form
    {
        //Название по-русски:
        //MS Dynamix CRM 4.0. Просмотр статистики значений полей в сущностях
        //Название по-английски:
        //MS Dynamix CRM 4.0. Entities fields values statistics view
        
        //Тестовая - "CmpnyLab", рабочая - "Cmpny".
        CrmService crmService1;
        MetadataService metadataService1;
        RetrieveAllEntitiesResponse metadata;
        AttributeMetadata currentAttribute;

        public Form1()
        {
            InitializeComponent();
        }

        //Функция, выполняемая при срабатывании таймера.
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false; //При создании формы таймер включен; при первом срабатывании отключаем, т. к. дальнейшие срабатывания не нужны.
            
            //Получить данные:
            getData();
        }

        //Получить данные:
        private void button1_Click(object sender, EventArgs e)
        {
            getData();
        }

        //Функция получения данных.
        //Выполняется при открытии формы и при нажатии button1.
        public void getData()
        {
            button1.Enabled = false;
            button1.Text = "Получение данных ...";
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            button3.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            //Показать форму еще до того, как весь код инициализации исполнился:
            this.Show();
            this.Refresh();    //Перерисовать немедленно, иначе изменение отобразится только после того как весь код кнопки отработает.

            //Очистить список "Сущность":
            comboBox1.Items.Clear();
            comboBox1.Items.Add("");
            comboBox1.Items.Clear();
            comboBox1.Text = "";
            comboBox1.Refresh();    //Перерисовать немедленно, иначе изменение отобразится только после того как весь код кнопки отработает.
            //Очистить список "Поле (атрибут)":
            comboBox2.Items.Clear();
            comboBox2.Items.Add("");
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            comboBox2.Refresh();    //Перерисовать немедленно, иначе изменение отобразится только после того как весь код кнопки отработает.
            
            try //Создаем сервисы данных и метаданных:
            {
                //Тестовая - "CmpnyLab", рабочая - "Cmpny".
                crmService1 = FunctionsToCrmDataWorking.createCrmService(textBox3.Text, textBox1.Text);
                metadataService1 = FunctionsToCrmDataWorking.createMetadataService(textBox3.Text, textBox2.Text);

                //Получить все сущности:
                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest();
                request.MetadataItems = MetadataItems.IncludeAttributes;
                metadata = (RetrieveAllEntitiesResponse)metadataService1.Execute(request);

                //Получить все сущности поэлементно в список:
                int resNumber1 = 0;
                foreach (EntityMetadata currentEntity in metadata.CrmMetadata)
                {
                    resNumber1++;
                    string nameEng = currentEntity.LogicalName;
                    string nameRus;
                    if (currentEntity.DisplayName.UserLocLabel != null)
                    {
                        nameRus = currentEntity.DisplayName.UserLocLabel.Label;
                    }
                    else
                    {
                        nameRus = "[Нет]";
                    }

                    comboBox1.Items.Add(new comboBoxEntity(resNumber1, nameRus, nameEng));
                }

                //Отсортировать список:
                int len = comboBox1.Items.Count;    //Длина списка.
                object immediateObject; //Объект для хранения промежуточного значения при обмене элементов списка значениями.
                bool isNeedNextIteration = true; //Нужна ли следующая итерация сортировки (если был обмен значениями на этой итерации, то нужна следующая).
                while (isNeedNextIteration)
                {
                    isNeedNextIteration = false;

                    for (int iii = 0; iii <= len - 2; iii++)    //Рассматриваем пары "этот-следующий", "этот" - от 0 до предпоследнего.
                    {
                        bool isNeedToBeChanged = false; //Надо обменять значения.

                        if ((comboBox1.Items[iii].ToString()[0] == '[') && (comboBox1.Items[iii + 1].ToString()[0] != '[')) //Если первый элемент начинается с "[" ("[Нет]"), а второй нет, то обменять их:
                        {
                            isNeedToBeChanged = true;
                        }
                        else //Иначе проверить, они по-русски или по-английски.
                        {
                            //Если первый элемент назван по-английски (первый символ английский), а второй нет (первый символ не английский) и второй начинается не с "[":
                            if ((FunctionsToStringWorking.isEnglishLetter(comboBox1.Items[iii].ToString()[0])) && (!(FunctionsToStringWorking.isEnglishLetter(comboBox1.Items[iii + 1].ToString()[0]))) && (comboBox1.Items[iii + 1].ToString()[0] != '['))
                            {
                                isNeedToBeChanged = true;
                            }
                            else //Иначе (если оба на одном языке и второй начинается не с "[") проверить, кто больше по алфавиту.
                            {
                                //Если (оба на одном языке и второй начинается не с "[")
                                //или (оба начинаются с "[") (первое "if") - 
                                //проверить, кто больше по алфавиту (второе "if"):
                                if
                                    (
                                    ((FunctionsToStringWorking.isEnglishLetter(comboBox1.Items[iii].ToString()[0]) == FunctionsToStringWorking.isEnglishLetter(comboBox1.Items[iii + 1].ToString()[0])) && (comboBox1.Items[iii + 1].ToString()[0] != '['))
                                    ||
                                    ((comboBox1.Items[iii].ToString()[0] == '[') && (comboBox1.Items[iii + 1].ToString()[0] == '['))
                                    )
                                    if (comboBox1.Items[iii].ToString().CompareTo(comboBox1.Items[iii + 1].ToString()) == 1)
                                    {
                                        isNeedToBeChanged = true;
                                    }
                            }
                        }

                        if (isNeedToBeChanged)  //Если надо обменять значения - обменять значения.
                        {
                            immediateObject = comboBox1.Items[iii + 1];
                            comboBox1.Items[iii + 1] = comboBox1.Items[iii];
                            comboBox1.Items[iii] = immediateObject;
                            isNeedNextIteration = true;
                        }
                    }
                }

                comboBox1.SelectedIndex = 0;

                //textBox4.Text = "Получен список сущностей.";
                //textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7); //Возвращаем шрифт какой был.

                button1.Enabled = true;
                button1.Text = "Получить данные еще раз";
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                button3.Enabled = true;
                this.Cursor = Cursors.Default;

                comboBox1.Focus();
            }
            catch   //Если не удалось создать сервисы данных и метаданных:
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("");
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox2.Items.Add("");
                comboBox2.Items.Clear();

                button1.Enabled = true;
                button1.Text = "Получить данные еще раз";
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                button3.Enabled = true;
                this.Cursor = Cursors.Default;
                
                MessageBox.Show("Не удалось создать CRM-сервис и/или Metadata-сервис.");
                
                textBox3.Focus();    //Назначение фокуса должно стоять после всех операций с внешним видом контролов, иначе фокус не назначится.
            }
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
                return this._nameRus + " (" + this._nameEng + ")";  // +" (" + this._id + ")";
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

        //Изменено поле "Сущность":
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = false;
            button3.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            //Очистить список "Поле (атрибут)":
            comboBox2.Items.Clear();
            comboBox2.Items.Add("");
            comboBox2.Items.Clear();
            comboBox2.Text = "";

            this.Refresh();    //Перерисовать форму немедленно, иначе изменение отобразится только после того как весь код кнопки отработает.

            if (comboBox1.SelectedItem != null)
            {
                string entityName = ((comboBoxEntity)comboBox1.SelectedItem).NameEng;   //Имя сущности - из списка 1.

                /*
                //Создать список типов. Атрибуты только этих типов будем отображать:
                List<Type> listType = new List<Type>()
                {
                    //Если тип - Picklist, Lookup, Status, State или другие выбранные:
                    typeof(PicklistAttributeMetadata), typeof(LookupAttributeMetadata), typeof(StatusAttributeMetadata), 
                    typeof(StateAttributeMetadata), typeof(BooleanAttributeMetadata), typeof(DateTimeAttributeMetadata)
                };
                */

                try //Получаем атрибуты для выбранной сущности:
                {
                    //Получить все нужные атрибуты нужной сущности поэлементно в список:
                    int resNumber2 = 0;
                    foreach (EntityMetadata currentEntity in metadata.CrmMetadata)
                    {
                        if (((comboBoxEntity)comboBox1.SelectedItem).NameEng == currentEntity.LogicalName)
                        {
                            #region Attributes
                            for (int j = 0; j < currentEntity.Attributes.Length; j++)
                            {
                                currentAttribute = currentEntity.Attributes[j];

                                Type attributeType = currentAttribute.GetType();

                                //if (listType.Contains(attributeType)) //Отображать только атрибуты определенных типов.
                                {
                                    resNumber2++;

                                    string nameRus;
                                    if (currentAttribute.DisplayName.UserLocLabel != null)
                                    {
                                        nameRus = currentAttribute.DisplayName.UserLocLabel.Label;
                                    }
                                    else
                                    {
                                        nameRus = "[Нет]";
                                    }

                                    string nameEng = currentAttribute.LogicalName;

                                    string type;
                                    switch (attributeType.Name)
                                    {
                                        case ("PicklistAttributeMetadata"): //"Picklist".
                                            type = "Picklist";
                                            break;
                                        case ("LookupAttributeMetadata"):   //"Lookup".
                                            type = "Lookup";
                                            break;
                                        case ("StatusAttributeMetadata"):   //"Status".
                                            type = "Status";
                                            break;
                                        case ("StateAttributeMetadata"):    //"State".
                                            type = "State";
                                            break;
                                        case ("BooleanAttributeMetadata"):  //"Boolean".
                                            type = "Boolean";
                                            break;
                                        case ("DateTimeAttributeMetadata"): //"DateTime".
                                            type = "DateTime";
                                            break;
                                        case ("StringAttributeMetadata"):   //"String".
                                            type = "String";
                                            break;
                                        case ("FloatAttributeMetadata"):    //"Float".
                                            type = "Float";
                                            break;
                                        case ("IntegerAttributeMetadata"):  //"Integer".
                                            type = "Integer";
                                            break;
                                        case ("MoneyAttributeMetadata"):    //"Money".
                                            type = "Money";
                                            break;
                                        case ("DecimalAttributeMetadata"):  //"Decimal".
                                            type = "Decimal";
                                            break;
                                        case ("MemoAttributeMetadata"):     //"Memo".
                                            type = "Memo";
                                            break;
                                        case ("AttributeMetadata"):         //"AttributeMetadata".
                                            type = "AttributeMetadata";
                                            break;
                                        default:
                                            type = "Another type (" + attributeType.Name.ToString() + ")";
                                            break;
                                    }

                                    comboBox2.Items.Add(new comboBoxAttribute(resNumber2, nameRus, nameEng, type));
                                }
                            }
                            #endregion

                            //Отсортировать список:
                            int len = comboBox2.Items.Count;    //Длина списка.
                            object immediateObject; //Объект для хранения промежуточного значения при обмене элементов списка значениями.
                            bool isNeedNextIteration = true; //Нужна ли следующая итерация сортировки (если был обмен значениями на этой итерации, то нужна следующая).
                            while (isNeedNextIteration)
                            {
                                isNeedNextIteration = false;

                                for (int iii = 0; iii <= len - 2; iii++)    //Рассматриваем пары "этот-следующий", "этот" - от 0 до предпоследнего.
                                {
                                    bool isNeedToBeChanged = false; //Надо обменять значения.

                                    if ((comboBox2.Items[iii].ToString()[0] == '[') && (comboBox2.Items[iii + 1].ToString()[0] != '[')) //Если первый элемент начинается с "[" ("[Нет]"), а второй нет, то обменять их:
                                    {
                                        isNeedToBeChanged = true;
                                    }
                                    else //Иначе проверить, они по-русски или по-английски.
                                    {
                                        //Если первый элемент назван по-английски (первый символ английский), а второй нет (первый символ не английский) и второй начинается не с "[":
                                        if ((FunctionsToStringWorking.isEnglishLetter(comboBox2.Items[iii].ToString()[0])) && (!(FunctionsToStringWorking.isEnglishLetter(comboBox2.Items[iii + 1].ToString()[0]))) && (comboBox2.Items[iii + 1].ToString()[0] != '['))
                                        {
                                            isNeedToBeChanged = true;
                                        }
                                        else //Иначе (если оба на одном языке и второй начинается не с "[") проверить, кто больше по алфавиту.
                                        {
                                            //Если (оба на одном языке и второй начинается не с "[")
                                            //или (оба начинаются с "[") (первое "if") - 
                                            //проверить, кто больше по алфавиту (второе "if"):
                                            if
                                                (
                                                ((FunctionsToStringWorking.isEnglishLetter(comboBox2.Items[iii].ToString()[0]) == FunctionsToStringWorking.isEnglishLetter(comboBox2.Items[iii + 1].ToString()[0])) && (comboBox2.Items[iii + 1].ToString()[0] != '['))
                                                ||
                                                ((comboBox2.Items[iii].ToString()[0] == '[') && (comboBox2.Items[iii + 1].ToString()[0] == '['))
                                                )
                                                if (comboBox2.Items[iii].ToString().CompareTo(comboBox2.Items[iii + 1].ToString()) == 1)
                                                {
                                                    isNeedToBeChanged = true;
                                                }
                                        }
                                    }

                                    if (isNeedToBeChanged)  //Если надо обменять значения - обменять значения.
                                    {
                                        immediateObject = comboBox2.Items[iii + 1];
                                        comboBox2.Items[iii + 1] = comboBox2.Items[iii];
                                        comboBox2.Items[iii] = immediateObject;
                                        isNeedNextIteration = true;
                                    }
                                }
                            }
                        }
                    }

                    comboBox2.SelectedIndex = 0;
                    comboBox2.Focus();
                }
                catch   //Если не удалось получить атрибуты для выбранной сущности:
                {
                    comboBox2.Items.Clear();
                    comboBox2.Items.Add("");
                    comboBox2.Items.Clear();
                    MessageBox.Show("Не удалось получить атрибуты для сущности: " + ((comboBoxEntity)comboBox1.SelectedItem).ToString());
                    comboBox1.Focus();
                }
            }
            else   //Эта ветвь не сработает в коде на comboBox1_SelectedIndexChanged. Она осталась от кода на кнопке.
            {
                MessageBox.Show("Не выбрана сущность!");
                comboBox1.Focus();
            }

            comboBox2.Enabled = true;
            button3.Enabled = true;
            this.Cursor = Cursors.Default;
        }

        //Изменено поле "Атрибут":
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                //string entityName = ((comboBoxEntity)comboBox1.SelectedItem).NameEng;       //Имя сущности - из списка 1.
                //string attributeName = ((comboBoxAttribute)comboBox2.SelectedItem).NameEng; //Имя атрибута - из списка 2.
                //string attributeType = ((comboBoxAttribute)comboBox2.SelectedItem).Type;    //Тип атрибута - из списка 2.

                button3.Focus();    //Назначение фокуса должно стоять после всех операций с внешним видом контролов, иначе фокус не назначится.
            }
            else   //Эта ветвь не сработает в коде на comboBox2_SelectedIndexChanged. Она осталась от кода на кнопке.
            {
                MessageBox.Show("Не выбран атрибут!");

                comboBox2.Focus();
            }
        }

        //Показать в Excel:
        private void button3_Click(object sender, EventArgs e)
        {
            if ((comboBox1.SelectedItem != null) && (comboBox2.SelectedItem != null))
            {
                this.Cursor = Cursors.WaitCursor;
                progressBar1.Value = 0;
                progressBar1.Visible = true;

                //listBox1.Items.Add("Выбранная сущность: " + ((comboBoxEntity)comboBox1.SelectedItem).NameRus + " (" + ((comboBoxEntity)comboBox1.SelectedItem).NameEng + ").");
                //listBox1.Items.Add("Выбранный атрибут: " + ((comboBoxAttribute)comboBox2.SelectedItem).NameRus + " (" + ((comboBoxAttribute)comboBox2.SelectedItem).NameEng + ").");

                //string entityName = ((comboBoxEntity)comboBox1.SelectedItem).NameEng;       //Имя сущности - из списка 1.
                //string attributeName = ((comboBoxAttribute)comboBox2.SelectedItem).NameEng; //Имя атрибута - из списка 2.
                //string attributeType = ((comboBoxAttribute)comboBox2.SelectedItem).Type;    //Тип атрибута - из списка 2.
                
                //Атрибуты текущей папки:
                DirectoryInfo dir1 = new DirectoryInfo(".");
                
                //Выбираем все карточки нужной сущности: ((comboBoxEntity)comboBox1.SelectedItem).NameEng, а из них - атрибут ((comboBoxAttribute)comboBox2.SelectedItem).NameEng:
                string fetch1 = @"
                    <fetch mapping=""logical"">
                        <entity name='" + ((comboBoxEntity)comboBox1.SelectedItem).NameEng + @"'>
                            <attribute name='" + ((comboBoxAttribute)comboBox2.SelectedItem).NameEng + @"'/>";

                /*
                //Дополнительное условие при выборе определенного атрибута:
                switch (((comboBoxAttribute)comboBox2.SelectedItem).Id)
                {
                    case 5004:        //"Интерес".  //"Результат работы ТМЦ", "gar_result_tmc".
                        fetch1 = fetch1 + @"<filter type='and'>";

                        //Добавляем условие "Маркетинговый список: Содержит любые данные":
                        fetch1 = fetch1 + @"<condition attribute = 'gar_list' operator='not-null'/>";

                        //Добавляем условие "Дата окончания работы с: ... по: ...":
                        fetch1 = fetch1 + @"<condition attribute = 'gar_end_date' operator='on-or-after' value='" + FunctionsToAnyDataWorking.getDataFormattedToUsingInTheFetchRequest(dateTimePicker1.Value) + @"'/>";
                        fetch1 = fetch1 + @"<condition attribute = 'gar_end_date' operator='on-or-before' value='" + FunctionsToAnyDataWorking.getDataFormattedToUsingInTheFetchRequest(dateTimePicker2.Value) + @"'/>";

                        if (comboBox3.SelectedItem != null)
                        {
                            //listBox1.Items.Add(((comboBoxAttribute)comboBox3.SelectedItem).Id);
                            //listBox1.Items.Add(((comboBoxAttribute)comboBox3.SelectedItem).NameEng);
                            //listBox1.Items.Add(((comboBoxAttribute)comboBox3.SelectedItem).NameRus);
                            //listBox1.Items.Add(dateTimePicker1.Value.ToString());

                            //Добавляем условие "Проект: Содержит выбранное значение":
                            fetch1 = fetch1 + @"<condition attribute = 'gar_projects' operator='eq' value='" + ((comboBoxAttribute)comboBox3.SelectedItem).NameEng + @"'/>";
                        }

                        fetch1 = fetch1 + @"</filter>";
                        break;
                    default:
                        break;
                }
                */

                fetch1 = fetch1 + @"
                        </entity>
                    </fetch>";
                string result1 = FunctionsToCrmDataWorking.fetchAll(fetch1, crmService1);
                XmlDocument xmlDoc1 = new XmlDocument();
                xmlDoc1.LoadXml(result1);
                List<CrmIndeterminatedObject> resList1 = getParsedItemsIndeterminatedObject(xmlDoc1, ((comboBoxAttribute)comboBox2.SelectedItem).NameEng, ((comboBoxAttribute)comboBox2.SelectedItem).Type);



                //Создать Excel-файл:
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlWorkBook;
                Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                string filePath1 = dir1.FullName + @"\Results.xlsx";
                //Открытие файла (если такой существует) или создание (если не существует):
                FileInfo fInfo = new FileInfo(filePath1);
                if (!fInfo.Exists)
                {
                    xlWorkBook = xlApp.Workbooks.Add(misValue); //Добавить новый Book в файл.
                    //Console.WriteLine("Файл создан.");
                }
                else //Открыть существующий файл.
                {
                    xlWorkBook = xlApp.Workbooks.Open(filePath1, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //Console.WriteLine("Файл открыт.");
                }
                //Открытие первой вкладки:
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);



                #region Работа с Xml.

                int resNumber1 = 0;
                int resNumber2 = 0;

                int writeIntoRow = 1;   //Строка, в которую писать.
                xlWorkSheet.Cells[writeIntoRow, 1] = "Выбранная сущность:";
                xlWorkSheet.Cells[writeIntoRow, 2] = ((comboBoxEntity)comboBox1.SelectedItem).NameRus + " (" + ((comboBoxEntity)comboBox1.SelectedItem).NameEng + ").";
                writeIntoRow++;
                xlWorkSheet.Cells[writeIntoRow, 1] = "Выбранный атрибут:";
                xlWorkSheet.Cells[writeIntoRow, 2] = ((comboBoxAttribute)comboBox2.SelectedItem).NameRus + " (" + ((comboBoxAttribute)comboBox2.SelectedItem).NameEng + ").";
                writeIntoRow++;
                writeIntoRow++;
                xlWorkSheet.Cells[writeIntoRow, 1] = "№:";
                xlWorkSheet.Cells[writeIntoRow, 2] = "Значение атрибута:";
                xlWorkSheet.Cells[writeIntoRow, 3] = "Количество вхождений:";
                writeIntoRow++;

                int countAtAll = resList1.Count;
                string countAtAllString = countAtAll.ToString();
                List<elementOfList> resSortList1 = new List<elementOfList>();

                string constantlyButton3Name = button3.Text;    //Имя кнопки: сохраняем до изменений, возвращаем после изменений.
                progressBar1.Maximum = resList1.Count;

                foreach (CrmIndeterminatedObject r in resList1)
                {
                    resNumber1++;
                    progressBar1.Value = resNumber1;
                    button3.Text = "Обрабатывается " + resNumber1.ToString() + " из " + countAtAllString + " ...";
                    button3.Refresh(); //Перерисовать немедленно, иначе изменение отобразится только после того как весь код кнопки отработает.

                    //Выбрать значение нужного параметра в отдельную переменную в зависимости от того, какой именно параметр нужен:
                    string neededParameter = "";
                    neededParameter = r.soughtAttributeName;    //"Бизнес-партнер" - "Категория", "accountcategorycode".

                    //Посчитать количество вхождений каждого значения
                    bool isInList = false;   //Уже в списке; изначально - нет.
                    resNumber2 = 0; //Счетчик.
                    foreach (elementOfList rr in resSortList1)
                    {
                        resNumber2++;
                        if (neededParameter == rr.name) //Если рассматриваемое значение == resNumber2-му элементу массива найденных.
                        {
                            rr.quantity++;  //Увеличить на единицу соответствующее количество - i-й элемент массива количеств.
                            isInList = true;    //Уже в списке.
                            break;
                        }
                    }
                    if (!isInList)  //Если еще не в списке, т. е. такое значение найдено впервые.
                    {
                        elementOfList newElem = new elementOfList(neededParameter, 1);
                        resSortList1.Add(newElem);
                    }
                }
                button3.Text = "Дополнительная обработка...";

                //Вывод окончательного результата - списка с подсчитанными количествами элементов:
                writeIntoRow = 5;   //Строка, в которую писать.
                resNumber2 = 0; //Счетчик.
                foreach (elementOfList rr in resSortList1)
                {
                    resNumber2++;

                    if (rr.name == "")
                    {
                        rr.name = "<Пустое значение>";
                    }
                    xlWorkSheet.Cells[writeIntoRow, 1] = resNumber2;
                    xlWorkSheet.Cells[writeIntoRow, 2] = rr.name;
                    xlWorkSheet.Cells[writeIntoRow, 3] = rr.quantity;

                    writeIntoRow++;
                }

                #endregion



                xlWorkSheet.Columns.AutoFit();  //Ширина столбцов - автоматически по содержимому.
                xlWorkSheet.Rows.AutoFit();
                //xlWorkSheet.get_Range("D1", "D3").EntireColumn.ColumnWidth = 40; //Ширина столбцов - задать для указанных столбцов, определяемых указанной областью.

                //Отсортировать строки: сначала по полю "Количество вхождений" (столбец C), затем по полю "Значение атрибута" (столбец B):
                Range sortingRange = xlWorkSheet.get_Range("B5", "D" + (writeIntoRow - 1).ToString());
                sortingRange.Sort(
                    sortingRange.Columns[2, Type.Missing],
                    XlSortOrder.xlDescending,
                    sortingRange.Columns[1, Type.Missing],
                    Type.Missing,
                    XlSortOrder.xlAscending,
                    sortingRange.Columns[3, Type.Missing],
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

                //Для того чтобы не было запроса на сохранение (по просьбе пользователей), отключаем Close и Quit:
                //Закрытие Excel:
                //xlWorkBook.Close(true, misValue, misValue);
                //xlApp.Quit();

                //Освобождение ресурсов:
                FunctionsToAnyDataWorking.releaseObject(xlWorkSheet);
                FunctionsToAnyDataWorking.releaseObject(xlWorkBook);
                FunctionsToAnyDataWorking.releaseObject(xlApp);

                FunctionsToAnyDataWorking.releaseObject(resList1);
                FunctionsToAnyDataWorking.releaseObject(xmlDoc1);
                FunctionsToAnyDataWorking.releaseObject(result1);

                progressBar1.Visible = false;
                this.Cursor = Cursors.Default;
                button3.Text = constantlyButton3Name;    //Имя кнопки: возвращаем после изменений.
            }
            else
            {
                MessageBox.Show("Не удалось получить статистику значений для атрибута.");
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.garsar.ru");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:a.bobov@garsar.ru");
            //System.Diagnostics.Process.Start("http://www.facebook.com/av.bobov");
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

        //Класс для получения объектов "Объект неопределенного заранее типа" из CRM.
        //Используется, когда в ходе выполнения программы определяется, какого типа объект из нескольких заранее предусмотренных следует выбрать из базы в данную переменную.
        //После того, как будет определен тип объекта, фетч-запрос вернет список объектов этого типа, при этом в наличии будут атрибуты только относящиеся к выбранному типу объекта.
        public class CrmIndeterminatedObject
        {
            public string soughtAttribute { get; set; }     //Искомый атрибут.
            public string soughtAttributeName { get; set; } //Имя искомого атрибута.
        }

        //Функция разбирает объекты в Xml-документе в список объектов "Объект неопределенного заранее типа".
        //Входные параметры:
        //XmlDocument xmlDoc - имя Xml-документа.
        //string attribute - имя атрибута, который искать при разборе Xml-документа.
        //string type - имя типа атрибута, в зависимости от которого брать параметр при разборе Xml-документа.
        //Выходные параметры:
        //List<CrmIndeterminatedObject> ParsedItems - список объектов CrmIndeterminatedObject, представляющих собой объекты "Объект неопределенного заранее типа" из CRM.
        public static List<CrmIndeterminatedObject> getParsedItemsIndeterminatedObject(XmlDocument xmlDoc, string attribute, string type)
        {
            try
            {
                XmlNodeList Itemslist = xmlDoc.SelectNodes("resultset/result");

                List<CrmIndeterminatedObject> DataCollection = new List<CrmIndeterminatedObject>();
                foreach (XmlNode node in Itemslist)
                {
                    CrmIndeterminatedObject currentNode = new CrmIndeterminatedObject();
                    currentNode.soughtAttribute = FunctionsToXmlDataWorking.ParseNodeValue(node, attribute);                        //Искомый атрибут.
                    try
                    {
                        switch (type)
                        {
                            case ("Picklist"):  //"Picklist".
                            case ("Status"):    //"Status".
                            case ("State"):     //"State".
                            case ("Boolean"):   //"Boolean".
                                currentNode.soughtAttributeName = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, attribute, "name");   //Имя искомого атрибута.
                                break;
                            case ("Lookup"):    //"Lookup".
                                try   //У атрибутов этих типов должно быть имя; попробовать взять его.
                                {
                                    currentNode.soughtAttributeName = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, attribute, "name");   //Имя искомого атрибута.
                                }
                                catch   //Если нет имени (пример - (системный?) атрибут ) - взять его непосредственное значение.
                                {
                                    currentNode.soughtAttributeName = FunctionsToXmlDataWorking.ParseNodeValue(node, attribute);                    //Непосредственное значение искомого атрибута.
                                }
                                break;
                            case ("DateTime"):  //"DateTime".
                                currentNode.soughtAttributeName = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, attribute, "date");   //Имя искомого атрибута (там еще есть параметр time).
                                //Привести дату к виду 2013.07.25:
                                try
                                {
                                    currentNode.soughtAttributeName = currentNode.soughtAttributeName.Substring(6, 4) + "." + currentNode.soughtAttributeName.Substring(3, 2) + "." + currentNode.soughtAttributeName.Substring(0, 2);
                                }
                                catch { }
                                break;
                            default:
                                //currentNode.soughtAttributeName = "<Другой тип данных!>";
                                currentNode.soughtAttributeName = FunctionsToXmlDataWorking.ParseNodeValue(node, attribute);                    //Непосредственное значение искомого атрибута.
                                break;
                        }
                    }
                    catch
                    {
                        currentNode.soughtAttributeName = "<Неизвестный тип данных!>";
                    }
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }
    }

    class FunctionsToStringWorking
    {
        //Функция возвращает ответ на вопрос, является ли символ буквой английского алфавита (строчной или прописной) или нет.
        //Входные параметры:
        //char letter - исходный символ.
        //Выходные параметры:
        //bool isEnglishLetter - true, если исходный символ является буквой английского алфавита (строчной или прописной), иначе - false.
        public static bool isEnglishLetter(char letter)
        {
            bool answer = false;
            if
                (
                ((letter>='A')&&(letter<='Z'))
                ||
                ((letter>='a')&&(letter<='z'))
                )
            {
                answer = true;
            }
            return answer;
        }
    }

    class FunctionsToAnyDataWorking
    {
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

    class FunctionsToXmlDataWorking
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

    class FunctionsToCrmDataWorking
    {
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
    }
}
