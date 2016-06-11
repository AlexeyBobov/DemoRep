using System;
using System.Collections.Generic;
using Microsoft.Win32;

//Microsoft Dynamics CRM namespaces:
using Microsoft.Crm.Sdk;
using Microsoft.Crm.SdkTypeProxy;
using Microsoft.Crm.SdkTypeProxy.Metadata;
using Microsoft.Crm.Sdk.Query;
using Microsoft.Crm.Sdk.Metadata;
using System.Xml;                               //Добавлен для разбора Xml-файла.

namespace Crm.Plugins
{
    public class ContractdetailOwnerChangingForAccountWithInstalledComplectOnOwnerChange : IPlugin
    {
        /*
        Разработано по заявке "Возврат актов" А. Домбровского от 15.10.2014 (http://portal/service_applications/sdesk/Lists/Posts/Post.aspx?ID=127).
        В рамках заявки необходимо, чтобы в поле Строки контракта "Ответственный" (gar_ownerid) (вкладка "Администрирование", секция "Администрирование", скрытое поле)
        был указан Пользователь, который является Ответственным за БП, указанный в поле Строки "БП, где установлен комплект" (gar_account).
        Смена Ответственного в Строке при смене "БП, где установлен комплект" реализована скриптами на форме.
        Смена Ответственного при смене такового в самом "БП, где установлен комплект" реализуется данным плагином.
        */
        //
        /*
        Регистрировать плагин со следующими параметрами:
        Assign
        account
        none
        ownerid
        (имя плагина, подставится само)
        Assign of account in Parent Pipeline
        Calling User
        1
        Pre, Synchronous, Server, Parent
        */
        //
        public void Execute(IPluginExecutionContext context)
        {
            if (context.Depth > 3) return;

            DynamicEntity entity = null;
            if (context.InputParameters.Properties.Contains(ParameterName.Target))
            {
                if (context.InputParameters.Properties["Target"] is DynamicEntity)
                {
                    entity = (DynamicEntity)context.InputParameters.Properties[ParameterName.Target];
                    if (context.MessageName == "Update")
                    {
                        //
                        //Код на событие "Update".
                        //
                    }   //Если это == "Update".
                }   //Если это DynamicEntity.
                else    //Если это не DynamicEntity.
                {
                    if (context.InputParameters.Properties["Target"] is Moniker)
                    {
                        if (context.InputParameters.Contains(ParameterName.Assignee))
                        {
                            Moniker e = (Moniker)context.InputParameters.Properties["Target"];
                            Guid accountid = e.Id;
                            
                            ICrmService service = context.CreateCrmService(false);

                            //Получить Ответственного после изменения:
                            Guid ownerid = ((SecurityPrincipal)context.InputParameters[ParameterName.Assignee]).PrincipalId;
                            systemuser powner = getSystemuserO(service, ownerid);

                            //Получить Ответственного до изменения из БП:
                            ColumnSet cs = new ColumnSet(new string[] { "ownerid", "statuscode", "tickersymbol", "gar_marketingovogom_list", "gar_resultjob", "gar_date_closed", "gar_verification_date" });
                            DynamicEntity e_pre = getInstanceOfDynamicEntityById(service, "account", accountid, cs);
                            systemuser powner_pre = getSystemuser(service, e_pre, "ownerid");
                            
                            if (powner.systemuserid.Value != powner_pre.systemuserid.Value) //Если сменился Ответственный, записать Историю работы.
                            {
                                string fetch1 = @"
                                    <fetch mapping=""logical"">
                                        <entity name=""contractdetail"">
                                            <attribute name=""contractdetailid""/>
                                            <filter type='and'>
                                                <condition attribute = 'gar_account' operator='eq' value='" + accountid.ToString() + @"'/>
                                                <condition attribute = 'statecode' operator='eq' value='0'/>
                                            </filter>
                                        </entity>
                                    </fetch>";
                                string result1 = service.Fetch(fetch1);
                                XmlDocument xmlDoc1 = new XmlDocument();
                                xmlDoc1.LoadXml(result1);
                                List<CrmContractdetail> resList1 = getParsedItemsContractdetail(xmlDoc1);

                                foreach (CrmContractdetail r1 in resList1)
                                {
                                    //Возвращение + изменение: Возвращает Строку контракта по заданному id + изменяет возвращенную Строку контракта:
                                    #region Retrieve + Update Contractdetail Dynamically

                                    ColumnSet cs1 = new ColumnSet();
                                    DynamicEntity entity1 = getInstanceOfDynamicEntityById(service, "contractdetail", r1.contractdetailid, cs1);

                                    //Изменяет возвращенную Строку контракта:
                                    LookupProperty gar_ownerid = new LookupProperty();
                                    gar_ownerid.Name = "gar_ownerid";
                                    gar_ownerid.Value = new Lookup();
                                    gar_ownerid.Value.type = "systemuser";
                                    gar_ownerid.Value.Value = ownerid;
                                    entity1.Properties.Add(gar_ownerid);
                                    TargetUpdateDynamic updateDynamic = new TargetUpdateDynamic();
                                    updateDynamic.Entity = entity1;
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
                                }
                            }
                        }   //Если это == "Assign".
                    }   //Если это Moniker.
                }   //Если это не DynamicEntity.
            }   //Если содержит "Target".
        }

        //Получить systemuser'а из поля prop сущности e:
        private systemuser getSystemuser(ICrmService service, DynamicEntity e, string prop)
        {
            systemuser s = null;
            if (e.Properties.Contains(prop))
            {
                Guid gp = getGuidOfLookupOrOwner(e.Properties[prop]);
                if (gp == new Guid("00000000-0000-0000-0000-000000000000")) return null;

                s = (systemuser)service.Retrieve("systemuser", gp, new ColumnSet(new string[] { "systemuserid", "fullname", "businessunitid" }));
            }
            return s;
        }

        //Получить systemuser'а по ИД:
        private systemuser getSystemuserO(ICrmService service, Guid systemuserid)
        {
            return (systemuser)service.Retrieve("systemuser", systemuserid, new ColumnSet(new string[] { "systemuserid", "fullname", "businessunitid" }));
        }

        //Получить Guid записи в объекте p, где (p is Lookup) или (p is Owner):
        private Guid getGuidOfLookupOrOwner(object p)
        {
            Guid gp = new Guid("00000000-0000-0000-0000-000000000000");
            if (p is Lookup)
            {
                Lookup lp = (Lookup)p;
                if (!lp.IsNull)
                {
                    gp = lp.Value;
                }
            }
            if (p is Owner)
            {
                Owner lp = (Owner)p;
                if (!lp.IsNull)
                {
                    gp = lp.Value;
                }
            }
            return gp;
        }

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

        //Класс для получения объектов "Роль" из CRM.
        public class CrmRole
        {
            public string roleid { get; set; }              //GUID.
            public string name { get; set; }                //Имя.
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
                    currentNode.contractdetailid = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractdetailid");//GUID.
                    currentNode.title = FunctionsToXmlDataWorking.ParseNodeValue(node, "title");//Наименование.
                    currentNode.contractid = FunctionsToXmlDataWorking.ParseNodeValue(node, "contractid");//Контракт.
                    currentNode.statuscode = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.statuscodename = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "statuscode", "name");
                    currentNode.statecode = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.customerid = FunctionsToXmlDataWorking.ParseNodeValue(node, "customerid");//Клиент, на кого заключен контракт.
                    currentNode.customeridname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "customerid", "name");
                    currentNode.gar_gar_stomost_without_discounts_abs = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_gar_stomost_without_discounts_abs");//Стоимость без скидок в АБС.
                    currentNode.gar_discount_from_contract_percentage = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_discount_from_contract_percentage");//Скидка из контракта (%).
                    currentNode.gar_discount_from_contract = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_discount_from_contract");//Скидка из контракта.
                    currentNode.gar_additional_discount_percentage = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_additional_discount_percentage");//Скидка вручную (%).
                    currentNode.gar_more_discount = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_more_discount");//Скидка вручную.
                    currentNode.gar_contract_number = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_contract_number");//Номер контракта.
                    currentNode.activeon = FunctionsToXmlDataWorking.ParseNodeValue(node, "activeon");//Дата начала.
                    currentNode.gar_activeon_fact = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_activeon_fact");//Дата фактического начала.
                    currentNode.gar_end_date_renovation_fact = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_end_date_renovation_fact");//Дата окончания.
                    currentNode.gar_end_date_renovation = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_end_date_renovation");//Дата фактического окончания.
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
                    currentNode.roleid = FunctionsToXmlDataWorking.ParseNodeValue(node, "roleid");
                    currentNode.name = FunctionsToXmlDataWorking.ParseNodeValue(node, "name");
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
                    currentNode.listid = FunctionsToXmlDataWorking.ParseNodeValue(node, "listid");//GUID.
                    currentNode.listname = FunctionsToXmlDataWorking.ParseNodeValue(node, "listname");//Имя.
                    currentNode.statecode = FunctionsToXmlDataWorking.ParseNodeValue(node, "statecode");//Статус.
                    currentNode.statuscode = FunctionsToXmlDataWorking.ParseNodeValue(node, "statuscode");//Состояние.
                    currentNode.createdon = FunctionsToXmlDataWorking.ParseNodeValue(node, "createdon");//Дата создания.
                    currentNode.gar_project = FunctionsToXmlDataWorking.ParseNodeValue(node, "gar_project");//Проект.
                    currentNode.gar_projectname = FunctionsToXmlDataWorking.ParseNodeAttributeValue(node, "gar_project", "name");
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
                    currentNode.listmemberid = FunctionsToXmlDataWorking.ParseNodeValue(node, "listmemberid");
                    currentNode.listid = FunctionsToXmlDataWorking.ParseNodeValue(node, "listid");
                    currentNode.entityid = FunctionsToXmlDataWorking.ParseNodeValue(node, "entityid");

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
                    currentNode.systemuserroleid = FunctionsToXmlDataWorking.ParseNodeValue(node, "systemuserroleid");
                    currentNode.systemuserid = FunctionsToXmlDataWorking.ParseNodeValue(node, "systemuserid");
                    currentNode.roleid = FunctionsToXmlDataWorking.ParseNodeValue(node, "roleid");
                    DataCollection.Add(currentNode);
                }
                return DataCollection;
            }
            catch { return null; }
        }

        #region Private methods

        /// <summary>
        /// Creates a CrmService proxy for plug-ins that execute in the child pipeline.
        /// </summary>
        /// <param name="context">The execution context that was passed to the plug-ins Execute method.</param>
        /// <param name="flag">Set to True to use impersontation.</param>
        /// <returns>A CrmService instance.</returns>
        private CrmService CreateCrmService(IPluginExecutionContext context, Boolean flag)
        {
            CrmAuthenticationToken authToken = new CrmAuthenticationToken();
            authToken.AuthenticationType = 0;
            authToken.OrganizationName = context.OrganizationName;

            // Include support for impersonation.
            if (flag)
                authToken.CallerId = context.UserId;
            else
                authToken.CallerId = context.InitiatingUserId;

            CrmService service = new CrmService();
            service.CrmAuthenticationTokenValue = authToken;
            service.UseDefaultCredentials = true;

            // Include support for infinite loop detection.
            CorrelationToken corToken = new CorrelationToken();
            corToken.CorrelationId = context.CorrelationId;
            corToken.CorrelationUpdatedTime = context.CorrelationUpdatedTime;
            corToken.Depth = context.Depth;

            RegistryKey regkey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\MSCRM");

            service.Url = String.Concat(regkey.GetValue("ServerUrl").ToString(), "/2007/crmservice.asmx");
            service.CorrelationTokenValue = corToken;

            return service;
        }

        /// <summary>
        /// Creates a MetadataService proxy for plug-ins that execute in the child pipeline.
        /// </summary>
        /// <param name="context">The execution context that was passed to the plug-ins Execute method.</param>
        /// <param name="flag">Set to True to use impersontation.</param>
        /// <returns>A MetadataService instance.</returns>
        private MetadataService CreateMetadataService(IPluginExecutionContext context, Boolean flag)
        {
            CrmAuthenticationToken authToken = new CrmAuthenticationToken();
            authToken.AuthenticationType = 0;
            authToken.OrganizationName = context.OrganizationName;

            // Include support for impersonation.
            if (flag)
                authToken.CallerId = context.UserId;
            else
                authToken.CallerId = context.InitiatingUserId;

            MetadataService service = new MetadataService();
            service.CrmAuthenticationTokenValue = authToken;
            service.UseDefaultCredentials = true;

            RegistryKey regkey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\MSCRM");

            service.Url = String.Concat(regkey.GetValue("ServerUrl").ToString(), "/2007/metadataservice.asmx");

            return service;
        }

        #endregion Private Methods
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
}
