//Определение CrmService и создание объекта сервиса:
loadScript('/ISV/ascentium/CrmService.js');
function loadScript(oScriptURL)
{
	var xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
	xmlHttp.open("GET", oScriptURL, false);
	xmlHttp.send();
	window.execScript(xmlHttp.responseText);
}
var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
//Конец определения CrmService и создания объекта сервиса.



//Изменение по заявке Зои Бунькиной от 01.03.2013 "Дублирование окон при назначении задачи" (http://wss30/sd/Lists/Posts/Post.aspx?ID=8200).
//
//Текст заявки:
//При переназначении ответственного в Задаче отображается несколько окон. При этом окна НЕ появляются, если фамилию в поле Ответственный набирать с клавиатуры. ...
//...Данная ситуация отмечена у всех пользователей, в рабочей и тестовой базе. По другим действиям все нормально.
//
//Решение заявки:
//Причина отображения нескольких окон - в использовании вызова обработчика события поля "В отношении" (crmForm.all.regardingobjectid.FireOnChange()), который вызывается при смене Ответственного...
//...(когда форма закрывается и открывается вновь), видимо конфликтуя с незакрытым в это время окном смены Ответственного.
//
//Внешний файл скрипта с вызовом и из OnLoad, и из OnChange не используется, т. к. его использование порождало непонятный глюк.
//
//Поэтому вместо вызова обработчика regardingobjectid.FireOnChange из OnLoad необходимые действия будут записаны в коде OnLoad. ...
//...Для избежания дублирования кода код будет оформлен в виде функции add_filters_and_fill_fields(deliveredOService), которая будет записана в OnLoad,...
//...а выполняться как из OnLoad, так и из regardingobjectid.OnChange, для чего она будет присвоена обработчику regardingobjectid.onchange.
//
//Скрипт на regardingobjectid.OnChange должен быть пустым, поскольку его код будет затираться присвоением regardingobjectid.onchange в OnLoad.

//Функция по установке/снятию фильтров и заполнению/обнулению полей, выполняемая из OnLoad и из regardingobjectid.OnChange:
function add_filters_and_fill_fields(deliveredOService)
{
	//Если на форме есть искомое поле и какой-нибудь объект в нем, то: 1. установить фильтры на поля, 2. заполнить необходимые поля соответствующими значениями.
	//Если нет, то снять фильтры и обнулить соответствующие поля.
	if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == 'account')
	//...и если Задача открыта в отношении Бизнес-партнера, т. е. в поле В отношении стоит Бизнес-партнер.
	{
		//Установка фильтра на поле LookUp.
		//Устанавливаем фильтр на поле Контакт в зависимости от значения поля В отношении.
		var field = crmForm.all.gar_contact;
		//Отключаем поле поиска в диалоговом окне лукапа:
		field.lookupbrowse = 1;
		//Передаем fetch xml через параметр поиска лукапа:
		field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
		+"<condition attribute='parentcustomerid' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
		+"</filter></entity></fetch>");
		
		//Установка фильтра на поле LookUp.
		//Устанавливаем фильтр на поле История работы в зависимости от значения поля В отношении.
		var field = crmForm.all.gar_history_task;
		//Отключаем поле поиска в диалоговом окне лукапа:
		field.lookupbrowse = 1;
		//Передаем fetch xml через параметр поиска лукапа:
		field.AddParam("search","<fetch mapping='logical'><entity name='gar_history'><filter>"
		+"<condition attribute='gar_history_bp' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
		+"</filter></entity></fetch>");
		
		//Получение списка атрибутов из объекта, выведенного в поле LookUp, и их вывод в поля открытой формы.
		//Получаем из Клиента (В отношении) атрибуты: Сотрудник, Руководитель (через обращение к Сотруднику); Ответственный ФО.
		var iDataSourceObject = new Array;
		iDataSourceObject = crmForm.all.regardingobjectid.DataValue;
		var iDataSourceObjectId = iDataSourceObject[0].id;
		var iCols = ["ownerid", "gar_systemuser_fo"];
		var beRetrievedDataSourceObject = deliveredOService.Retrieve("account", iDataSourceObjectId, iCols);
		
		if ("ownerid" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
		{
			//Если надо передать значение в поле LookUp:
			//Выводим значение путем создания объекта и записи его в нужное поле:
			var o1 = new Object();
			o1.id =  beRetrievedDataSourceObject.attributes["ownerid"].value;
			o1.typename = 'systemuser';
			o1.name =  beRetrievedDataSourceObject.attributes["ownerid"].name;
			crmForm.gar_worker.DataValue = [o1];
			
			//Если Сотрудник выведен в поле, то находим его Руководителя и выводим в другое поле:
			if (crmForm.gar_worker.DataValue != null)
			{
				var iCols2 = ["parentsystemuserid"];
				var beRetrievedDataSourceObject2 = deliveredOService.Retrieve("systemuser", o1.id, iCols2);
				
				if ("parentsystemuserid" in beRetrievedDataSourceObject2.attributes)	//Если атрибут есть и != null.
				{
					//Если надо передать значение в поле LookUp:
					//Выводим значение путем создания объекта и записи его в нужное поле:
					var o2 = new Object();
					o2.id =  beRetrievedDataSourceObject2.attributes["parentsystemuserid"].value;
					o2.typename = 'systemuser';
					o2.name =  beRetrievedDataSourceObject2.attributes["parentsystemuserid"].name;
					crmForm.gar_head.DataValue = [o2];
				}else
				{
					crmForm.all.gar_head.DataValue = null;
				}
			}else
			{
				crmForm.all.gar_head.DataValue = null;
			}
		}else
		{
			crmForm.gar_worker.DataValue = null;
		}
		
		if ("gar_systemuser_fo" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
		{
			//Если надо передать значение в поле LookUp:
			//Выводим значение путем создания объекта и записи его в нужное поле:
			var o3 = new Object();
			o3.id =  beRetrievedDataSourceObject.attributes["gar_systemuser_fo"].value;
			o3.typename = 'systemuser';
			o3.name =  beRetrievedDataSourceObject.attributes["gar_systemuser_fo"].name;
			crmForm.gar_systemuser_fo_task.DataValue = [o3];
		}else
		{
			crmForm.gar_systemuser_fo_task.DataValue = null;
		}
		//Конец получения списка атрибутов и их вывода в поля открытой формы.
	}else
	{
		crmForm.all.gar_contact.DataValue = null;	//При выборе не БП обнулять Контакт.
		crmForm.all.gar_fax_number.DataValue = null;	//При выборе не БП обнулять Номер факса.
		crmForm.all.gar_worker.DataValue = null;	//При выборе не БП обнулять Сотрудника.
		crmForm.all.gar_head.DataValue = null;	//При выборе не БП обнулять Руководителя.
		crmForm.all.gar_systemuser_fo_task.DataValue = null;	//При выборе не БП обнулять Ответственного ФО.
		
		//При выборе не БП снимать фильтр с Контакта:
		var field = crmForm.all.gar_contact;
		//Отключаем поле поиска в диалоговом окне лукапа:
		field.lookupbrowse = 1;
		//Передаем fetch xml через параметр поиска лукапа:
		field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
		+"</filter></entity></fetch>");
		
		//При выборе не БП снимать фильтр с Истории работ:
		var field = crmForm.all.gar_history_task;
		//Отключаем поле поиска в диалоговом окне лукапа:
		field.lookupbrowse = 1;
		//Передаем fetch xml через параметр поиска лукапа:
		field.AddParam("search","<fetch mapping='logical'><entity name='gar_history'><filter>"
		+"</filter></entity></fetch>");
	}
	//Конец проверки, есть ли на форме искомое поле и объект в нем, и заполнения необходимых полей соответствующими значениями.
}

//Выполнить функцию по установке/снятию фильтров и заполнению/обнулению полей отсюда (из OnLoad):
add_filters_and_fill_fields(oService);

//Функция, которую следует повесить на обработчик поля "В отношении" (regardingobjectid.OnChange). Содержит функцию по установке/снятию фильтров и заполнению/обнулению полей, + возможные дополнительные действия:
function add_filters_and_fill_fields_forOnChange()
{
	//Выполнить функцию по установке/снятию фильтров и заполнению/обнулению полей из regardingobjectid.OnChange:	
	add_filters_and_fill_fields(oService);
	
	//Выполнить какие-либо дополнительные действия:
	//alert('add_filters_and_fill_fields_forOnChange');
}

//Повесить функцию на обработчик поля "В отношении" (regardingobjectid.OnChange):
crmForm.all.regardingobjectid.onchange = add_filters_and_fill_fields_forOnChange;

//Конец изменения по заявке Зои Бунькиной от 01.03.2013 "Дублирование окон при назначении задачи" (http://wss30/sd/Lists/Posts/Post.aspx?ID=8200).



//Фильтрация маркетинговых списков по ответственному
var cb = crmForm.all.createdby.DataValue;
var o = crmForm.all.ownerid.DataValue;
var field = crmForm.all.gar_tasks;
var fs = '';
var cbid = null;
var oid = null;
if (field != null)
{
    field.lookupbrowse = 1; 
    if (cb != null)
    {
        var cbid = cb[0].id;
    }
    if (o != null)
    {
        var oid = o[0].id;
    }
    if (cbid != null || oid != null)
    {
        fs = "<fetch mapping='logical'><entity name='list'>" 
           + "<filter type='and'><condition attribute='statecode' operator='eq' value='0' />"
           + "<condition attribute='ownerid' operator='in'>";
        if (cbid != null)
        {
           fs = fs + "<value>" + cbid + "</value>";
        }
        if (oid != null)
        {
           fs = fs + "<value>" + oid + "</value>";
        }
        fs = fs + "</condition></filter></entity></fetch>";
        field.AddParam("search", fs);
    }
    else
    {
	field.AddParam("search", "<fetch mapping='logical'><entity name='list'>" 
           + "<filter type='and'><condition attribute='statecode' operator='eq' value='0' />"
           + "<condition attribute='ownerid' operator='eq-userid' "
           + "/></filter></entity></fetch>");
    }
}



//Код, проверяющий и подставляющий правильное название БП взамен ошибочного:
if (crmForm.FormType == 2)	//Если обрабатывается существующая, не завершенная задача.
{
	if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == "account")
	{
		var bp = crmForm.all.regardingobjectid.DataValue;
		
		var iCols = ["name"];
		var beRetrievedDataSourceObject = oService.Retrieve("account", bp[0].id, iCols);	//По текущему id БП берем его же из базы.
		if ("name" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
		{
			var o = new Object();
			o.id = bp[0].id;
			o.typename = bp[0].typename;
			o.name = beRetrievedDataSourceObject.attributes["name"].value;
		}
		
		if (bp[0].name != o.name)
		{
			crmForm.all.regardingobjectid.DataValue = null;
			crmForm.all.regardingobjectid.DataValue = [o];
			
			//var beTaskToUpdate = new BusinessEntity("task");
			//beTaskToUpdate.attributes["activityid"] = crmForm.ObjectId;
			//beTaskToUpdate.attributes["regardingobjectid"] = o.id;
			//oService.Update(beTaskToUpdate);
			
			/*
			alert
			(
				'Наименование БП в поле "В отношении" не обновилось ранее.\n'+
				'Старое наименование: '+bp[0].name+'\n'+
				'Новое наименование: '+o.name+'\n'+
				'Сохраните форму перед закрытием.\n'+
				''
			);
			*/
		}
	}
}



//Доработка по заявке "Поля в задаче" (http://portal/service_applications/sdesk/_layouts/15/start.aspx#/Lists/Posts/Post.aspx?ID=1077) от 28.07.2015 (Яна Бизюкова).
//Если поле "Необходимо сделать" (gar_do) пустое – убрать. Если поле непустое - запретить редактирование (сделано в настройках формы, код не нужен).
loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
var IAm = WhoAmI(oService);	//Для вызова этой функции необходимо: loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');

var departments = {
	'ОРК': '',
	'ОРК VIP': '',
	'ОРК ДГ': '',
	'ОРК ИТП': '',
	'ОРК КК': '',
	'ОРК КОРК': '',
	'ОРК ОК': '',
	'ОРК РК': '',
	'ОРК СНК': ''
};

var businessunitName = IAm.attributes["businessunitid"].name;	//Получить имя подразделения пользователя из функции WhoAmI (выше по коду).
if (!(businessunitName in departments))	//Если подразделение текущего пользователя - не ОРК:
{
	if (crmForm.all.gar_do != null)	//Если поле "Необходимо сделать" на форме есть.
	{
		if (crmForm.all.gar_do.DataValue == null)	//Если поле на форме есть, но оно пустое.
		{
			//Делаем невидимым поле на форме:
			crmForm.all.gar_do_c.style.display = "none";
			crmForm.all.gar_do_d.style.display = "none";
		}
		else	//Если поле непустое.
		{
			//Делаем недоступным для редактирования поле на форме:
			crmForm.all.gar_do.disabled = true;
			crmForm.all.gar_do.readOnly = true;
		}
	}
}
//Конец доработки по заявке "Поля в задаче" (http://portal/service_applications/sdesk/_layouts/15/start.aspx#/Lists/Posts/Post.aspx?ID=1077) от 28.07.2015 (Яна Бизюкова).