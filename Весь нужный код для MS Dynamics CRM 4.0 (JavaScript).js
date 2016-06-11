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



//Проверка, есть ли на форме искомое поле и есть ли в искомом поле какой-нибудь объект.
//Если есть, заполнить необходимые поля соответствующими значениями.
//Если нет, обнулить соответствующие поля.
if (crmForm.all.gar_contact != null && crmForm.all.gar_contact.DataValue != null)
{
	
	//Получение списка атрибутов из объекта, выведенного в поле LookUp, и их вывод в поля открытой формы.
	//Получаем из Клиента атрибуты: Режим налогообложения, Классификация, Форма собственности, Условное объединение с иными БП.
	var iDataSourceObject = new Array;
	iDataSourceObject = crmForm.all.customer.DataValue;
	var iDataSourceObjectId = iDataSourceObject[0].id;
	var iCols = ["businesstypecode", "accountclassificationcode", "ownershipcode", "gar_conditional_association"];
	var beRetrievedDataSourceObject = oService.Retrieve("account", iDataSourceObjectId, iCols);
	
	if ("businesstypecode" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		crmForm.all.gar_businesstypecode.DataValue = beRetrievedDataSourceObject.attributes["businesstypecode"].name;
	}else
	{
		crmForm.all.gar_businesstypecode.DataValue = null;
	}
	
	if ("accountclassificationcode" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		crmForm.all.gar_accountclassificationcode_bp.DataValue = beRetrievedDataSourceObject.attributes["accountclassificationcode"].name;
	}else
	{
		crmForm.all.gar_accountclassificationcode_bp.DataValue = null;
	}
	
	if ("ownershipcode" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		crmForm.all.gar_ownershipcode_bp.DataValue = beRetrievedDataSourceObject.attributes["ownershipcode"].name;
	}else
	{
		crmForm.all.gar_ownershipcode_bp.DataValue = null;
	}
	
	if ("gar_conditional_association" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		//Передаем значение булевого типа через условие, потому что value в DataValue напрямую не присваивается (не тот тип):
		switch (beRetrievedDataSourceObject.attributes["gar_conditional_association"].value)
		{
			case '0':
				crmForm.all.gar_conditional_association_bp.DataValue = false;
				break;
			case '1':
				crmForm.all.gar_conditional_association_bp.DataValue = true;
				break;
		}
	}else
	{
		crmForm.all.gar_conditional_association_bp.DataValue = null;
	}
	
	if ("gar_list_phonecalls" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		//Если надо передать значение в поле LookUp:
		//Записываем значение в поле LookUp путем создания объекта и записи его в нужное поле:
		var o = new Object();
		o.id = beRetrievedDataSourceObject.attributes["gar_list_phonecalls"].value;
		o.typename = 'list';
		o.name = beRetrievedDataSourceObject.attributes["gar_list_phonecalls"].name;
		crmForm.gar_list.DataValue = [o];
	}else
	{
		crmForm.gar_list.DataValue = null;
	}
	
	if ("gar_stoimost_with_discounts" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		//Передаем значение типа Money, отняв 0 для приведения к числовому типу, потому что в DataValue поля типа Money значение string (из value) напрямую не присваивается (не тот тип):
		crmForm.all.gar_amount.DataValue = beRetrievedDataSourceObject.attributes["gar_stoimost_with_discounts"].value - 0;
	}
	
	if ("gar_activeon_fact" in beRetrievedDataSourceObject.attributes)	//Если атрибут есть и != null.
	{
		//Если надо передать значение в поле DateTime:
		//Передаем значение типа Date, приведя его к требуемому формату, потому что в DataValue поля типа Date значение string (из value) напрямую не присваивается (не тот тип):
		crmForm.all.gar_start_of.DataValue = new Date(beRetrievedDataSourceObject.attributes["gar_activeon_fact"].value.replace(/(\d+)-(\d+)-(\d+)/, '$2/$3/$1').substr(0, 10));
	}
	//Конец получения списка атрибутов и их вывода в поля открытой формы.
	
}else
{
	crmForm.all.gar_post.DataValue = null;
	crmForm.all.emailaddress.DataValue = null;
}
//Конец проверки, есть ли на форме искомое поле и объект в нем, и заполнения необходимых полей соответствующими значениями.



//Если не создается вновь, а редактируется:
if (crmForm.FormType != 1)



//Привести даты к целочисленному формату для сравнения:
var today = new Date();
var choosenDayInt = parseInt(crmForm.all.activeon.DataValue.toFormattedString("yyyymmdd"));
var todayInt = parseInt(today.toFormattedString("yyyymmdd"));



//Проверка, есть ли на форме искомое поле и есть ли в искомом поле какой-нибудь объект.
//Если есть, то в общем случае: 1. установить фильтр на поле, 2. заполнить необходимые поля соответствующими значениями.
//Если нет, обнулить соответствующие поля.
if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null)


	
	//Установка фильтра на поле LookUp.
	//Устанавливаем фильтр на поле Контакт в зависимости от значения поля В отношении.
	var field = crmForm.all.gar_contact;
	//Отключаем поле поиска в диалоговом окне лукапа:
	field.lookupbrowse = 1;
	//Передаем fetch xml через параметр поиска лукапа:
	field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
	+"<condition attribute='parentcustomerid' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
	+"</filter></entity></fetch>");
	
	
	
	//Снятие фильтра с поля LookUp.
	var field = crmForm.all.gar_contact_assistant;
	field.lookupbrowse = 1;
	field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
	+"</filter></entity></fetch>");



//Еще условия фильтрования бывают, например, такими:
field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
+"<condition attribute='statecode' operator='eq' value='0' />"
+"<condition attribute='customertypecode' operator='ne' value='200000' />"
+"<condition attribute='parentcustomerid' operator='eq' value='" + receiver[0].id + 
"' /></filter></entity></fetch>");



//Получение значения поля типа "пиклист" по его цифровому коду (attributename - имя поля, objecttypecode - тип сущности, attributevalue - код требуемого значения):
//Пример 1 - старый кусок:
	var sFetchXml = "<fetch mapping=\"logical\"><entity name=\"stringmap\"><attribute name=\"value\"/><filter type=\"and\">"
		+"<condition attribute='attributename' operator='eq' value='gar_function' />"
		+"<condition attribute='objecttypecode' operator='eq' value='2' />"
		+"<condition attribute='attributevalue' operator='eq' value='"+id+"' /></filter></entity></fetch>";
//Пример 2 - красивый отлаженный кусок:
	//Получение значения поля типа "пиклист" по его цифровому коду (attributename - имя поля, objecttypecode - тип сущности, attributevalue - код требуемого значения):
	if (crmForm.all.gar_priority_training.DataValue != null)
	{
		var pickFetch = ""+
			"<fetch mapping='logical'>"+
				"<entity name='stringmap'>"+
					"<attribute name='value'/>"+
					"<filter type='and'>"+
						"<condition attribute='attributename' operator='eq' value='gar_priority_training' />"+
						"<condition attribute='objecttypecode' operator='eq' value='4214' />"+
						"<condition attribute='attributevalue' operator='eq' value='" + crmForm.all.gar_priority_training.DataValue + "' />"+
					"</filter>"+
				"</entity>"+
			"</fetch>";
		var result = oService.Fetch(pickFetch);
		crmForm.all.subject.DataValue = result[0].attributes["value"].value;
	}
	else
	{
		crmForm.all.subject.DataValue = null;
	}
//Здесь objecttypecode (тип сущности):
//1 - Бизнес-партнер (account).
//2 - Контакт (contact).
//3 - Возможная сделка (opportunity).
//4 - Интерес (lead).
//1010 - Контракт (contract).
//1011 - Строка контракта (contractdetail).
//4214 - Действие сервиса (serviceappointment).
//10004 - Разнесение оплаты (gar_paysplit).
//...Прочие можно брать из полной выгрузки - Export all customizations.



//Сделать невидимыми (запретить для выбора) следующие значения поля Результат работы (Г) (gar_resultjob) типа "пиклист":
//Не обзванивается (значение 16), Не хотят смотреть либо чужая СПС (9), Совсем не отвечают (7), Стоплист (13), Клиенты по И-В (1), Судебное разбирательство (10).
var exceptionsList = {'1': '', '7': '', '9': '', '10': '', '13': '', '16': ''};	//Список исключений (значения, которые не показывать в выпадающем пиклисте на форме).
var rsj = crmForm.all.gar_resultjob;	//Сокращаем имя поля для краткости записей.
var dv = rsj.DataValue;					//Сохраняем текущее значение поля.
var resultArray = new Array();			//Список для отображения (массив значений, которые будут отображены).
var j = 0;
for (i = 0; i < rsj.Options.length; i++)	//Рассматриваем общий список значений.
{
	//Из общего списка добавить в список для отображения те значения, которых нет в списке исключений;...
	//...кроме того, добавить текущее значение, даже если оно в списке исключений (это нужно потому, что иначе, если текущее значение в списке исключений,...
	//...то оно не будет выведено в поле при открытии карточки - поле останется пустым).
	if (!(rsj.Options[i].DataValue in exceptionsList) || (rsj.Options[i].DataValue == dv))
	{
		resultArray[j] = rsj.Options[i];
		j++;
	}
}
rsj.Options = resultArray;				//Возвращаем полученный список значений.
rsj.DataValue = dv;						//Восстанавливаем первоначальное значение.



//Примечание для значений полей денежного формата:
//Запись crmForm.all.gar_amount.value		дает значение вида 13 699,80 - т. е. с учетом форматирования и разделителей для денежного формата.
//Запись crmForm.all.gar_amount.DataValue	дает значение вида 13699.8 - т. е. в виде числа, без учета форматирования и разделителей для денежного формата.



//Получение списка атрибутов из объекта, чей id известен, можно обернуть в такую функцию:
//(привожу Функцию, затем Кусок кода с ее использованием).

//Функция получения списка атрибутов "cols" из объекта типа "typename" с известным id:
function getFields(typename, id, cols)
{
	//Требует var oService = new Ascentium_CrmService, который уже определен в блоке "Определение CrmService и создание объекта сервиса"
	//var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
	var c = oService.Retrieve(typename, id, cols);
	return c.attributes;	//["name"].value
}

//Кусок кода с ее использованием:
var field = crmForm.all.customerid;
if (!IsNull(field) && !IsNull(field.DataValue))
{
	if (field.DataValue[0].typename == 'account')
	{
		var id = field.DataValue[0].id;
		var ac = getFields('account', id, ['gar_competitor']);
		if ('gar_competitor' in ac)
		{
			idc = ac['gar_competitor'].value;
			c = getFields('competitor', idc, ['name']);
			var o = new Object();
			o.id = idc;
			o.typename = 'competitor';
			o.name = c['name'].value;
			crmForm.all.gar_competitor.DataValue = [o];
		}else
		{
			crmForm.all.gar_competitor.DataValue = null;
		}
	}
}



Строка
	if ("businesstypecode" in beRetrievedDataSourceObject.attributes)
заменяет строку
	if (beRetrievedDataSourceObject.attributes["businesstypecode"] != null && beRetrievedDataSourceObject.attributes["businesstypecode"].value != null)



//Если заполнено поле "Строка заказа", запретить его изменение:
if (crmForm.all.gar_demontratsy != null && crmForm.all.gar_demontratsy.DataValue != null)
{
    crmForm.all.gar_demontratsy.Disabled = true;
}



//switch...case...default:
	switch (t)
	{
		case '{63220E4D-69CF-DE11-8E20-00155D4E1B14}':	//Если Тема==ГЛ.
			//Какой-то код.
			break;
		case '{8B0ED757-69CF-DE11-8E20-00155D4E1B14}':	//Если Тема==Обучение.
		case '{0854A205-59F1-DE11-8853-00155D4E1B14}':	//Если Тема==Обучение УЦ.
			//Какой-то код.
			break;
		default:	//Если Тема - ни одна из указанных.
			//Какой-то код.
			break;
	}



//Делаем невидимым поле на форме:
crmForm.all.gar_account_c.style.display = "none";
crmForm.all.gar_account_d.style.display = "none";



//Делаем недоступным для редактирования поле на форме:
crmForm.all.gar_name.disabled = true;
crmForm.all.gar_name.readOnly = true;



//Сохранение предыдущих значений:
crmForm.all.activeon.prevDataValue = CloneDate(crmForm.all.activeon.DataValue);			//Висело на системном OnLoad в файле customizations.
crmForm.all.parentcustomerid.prevDataValue = crmForm.all.parentcustomerid.DataValue;	//Моя обработка (Контакт.OnLoad и OnSave).



//Взятие текста, выбранного в пиклистовом поле на текущей форме:
crmForm.all.gar_name.DataValue = crmForm.all.gar_type_agreement.SelectedText;



//Скрытие некоторых значений поля типа "пиклист":
cert = crmForm.all.gar_certificate;
certValue = cert.DataValue;
switch (certValue)
{
	case '3':	//Если Сертификат=="Серебряный АЭРО".
	case '4':	//Если Сертификат=="Золотой АЭРО".
		//Отобразить только строки со значениями: '' (пустая строка - первоначальное значение), '3' ("Серебряный АЭРО"), '4' ("Золотой АЭРО"):
		var oTempArray = new Array();
		var iIndex = 0;
		for (i=0; i<cert.Options.length; i++)
		{
			if (cert.Options[i].DataValue in {'': '', '3': '', '4': ''})
			{
				oTempArray[iIndex] = cert.Options[i];
				iIndex++;
			}
		}
		cert.Options = oTempArray;
		//Вернуть в поле Сертификат нужное значение:
		crmForm.all.gar_certificate.DataValue = certValue;
		break;
	case null:	//Если Сертификат==null (не выбран).
		//Отобразить только строки со значениями: '' (пустая строка - первоначальное значение), '3' ("Серебряный АЭРО"), '4' ("Золотой АЭРО"):
		var oTempArray = new Array();
		var iIndex = 0;
		for (i=0; i<cert.Options.length; i++)
		{
			if (cert.Options[i].DataValue in {'': '', '3': '', '4': ''})
			{
				oTempArray[iIndex] = cert.Options[i];
				iIndex++;
			}
		}
		cert.Options = oTempArray;
		break;
	default:	//Если в поле Сертификат неизвестное значение.
		alert('Неизвестное значение в поле Сертификат!');
		break;
}



//Записать в поле лукап текущего пользователя:
loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
var IAm = WhoAmI(oService);	//Для вызова этой функции необходимо: loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
//Передать значение в поле LookUp. Выводим значение путем создания объекта и записи его в нужное поле:
var o = new Object();
o.id = IAm.attributes["systemuserid"].value;
o.typename = 'systemuser';
o.name = IAm.attributes["fullname"].value;
crmForm.all.resources.DataValue = [o];



//GUID текущего объекта (открытого в открытом окне):
var it = crmForm.ObjectId;



//В зависимости от того, в тестовой или в рабочей базе:
if (ORG_UNIQUE_NAME == "CmpnyLab")	{ var lookupViewId = "{6FF24503-5D6F-E211-91EA-00155D001308}"; }
if (ORG_UNIQUE_NAME == "Cmpny")		{ var lookupViewId = "{E9E0ED78-BA75-E211-9FAF-00155D001308}"; }



//Некоторые параметры полей:
crmForm.all.gar_contact.outerHTML
crmForm.all.gar_contact.lookupstyle='multi';	//Бывает 'multi' или 'single'.
Параметры лукапа по умолчанию (хранятся в параметре outerHTML):	//<IMG style="IME-MODE: auto" id=gar_contact class=ms-crm-Lookup title="Click to select a value for Контакты." alt="Click to select a value for Контакты." src="/_imgs/btn_off_lookup.gif" req="0" resolveemailaddress="0" showproperty="1" autoresolve="1" defaulttype="0" lookupstyle="single" lookupbrowse="0" lookuptypeIcons="/_imgs/ico_16_2.gif" lookuptypenames="contact:2" lookuptypes="2">



//Перенос секции из одной вкладки в другую:
var t1 = document.getElementById('tab' + 1);	//Вкладка 1 (считая с нуля), откуда забрать элемент.
var t2 = document.getElementById('tab' + 2);	//Вкладка 2 (считая с нуля), куда вставить элемент.
//Вариант 1. Заменить секцию "Эта секция в коде заменяется..." из вкладки "Запрос в ТО" (t2.childNodes[0].rows[2]) на секцию "Примечания" из вкладки "Запрос на ГЛ" (t1.childNodes[0].rows[3]):
t2.childNodes[0].rows[2].parentNode.replaceChild(t1.childNodes[0].rows[3], t2.childNodes[0].rows[2]);
//Вариант 2. Взять произвольную имеющуюся секцию из вкладки "Запрос в ТО" (t2.childNodes[0].rows[1]), взять ее родительский элемент,...
//...и добавить к нему секцию "Примечания" из вкладки "Запрос на ГЛ" (t1.childNodes[0].rows[3]) последним дочерним элементом:
t2.childNodes[0].rows[1].parentNode.appendChild(t1.childNodes[0].rows[3]);



//Принудительное закрытие формы:
window.close();



//Работа со сторонним окном. Открытие, вывод текста, закрытие, закрытие с паузой:
var newWindow = window.open('kml.htm', 'displayWindow', 'width=500,height=400,status=yes,toolbar=yes,menubar=yes');
newWindow.document.write('Это окно будет закрыто через 2 секунды после окончания работы кода (после закрытия всех окон alert и т. д.).');
setTimeout(function(){ newWindow.close(); }, 2000);

newWindow.close();			//Закрыть документ и окно.
newWindow.document.close();	//Закрыть документ, но оставить открытым окно.



//Работа с окном подтверждения:
if (confirm("Загрузить в гуглкарты или сохранить в файл?"))
{
	alert("\"Загрузить в гуглкарты\" в разработке!");
}
else
{
	alert("\"Сохранить в файл\" в разработке!");
}



//Проверка на ошибки:
try
{
	var a = 5;
	var res = a(1); // ошибка!
}
catch(err)
{
	alert("name: " + err.name + "\nmessage: " + err.message + "\nstack: " + err.stack);
}



//Перебор всех имеющихся свойств объекта:
var ooo = {a:5, b:true}
for (var key in ooo)
{
	alert(key+' : '+ooo[key]);
}



//
//
//Далее - не мои скрипты, а из разных источников:
//
//



//Скрытие некоторых секций:
function HSSection(tabIndex, sectionIndex, displayType)
{
	var s = document.getElementById('tab' + tabIndex);
	s.childNodes[0].rows[sectionIndex].style.display = displayType;
}
HSSection(0, 2, 'none');
crmForm.all.tab1Tab.style.display = 'none';



//Проверка типа значения в поле LookUp:
var receiver = crmForm.all.to.DataValue;
if (receiver[0].type==1) //т. е. в получателе стоит бизнес-партнер.



//Проверка принадлежности объекта к одному из нескольких заданных типов:
var field = crmForm.all.regardingobjectid;
if (!IsNull(field) && !IsNull(field.DataValue))
{
    var dv = field.DataValue[0];
    if (dv.typename in {'account': '', 'contact': ''})
    {
		//Код.
    }
}



//Вызов обработчика события одного поля из другого поля или из OnLoad, OnSave:
crmForm.all.businesstypecode.FireOnChange();



//фильтрация полей возможной сделки по БП из поля Потенциальный клиент:

function setSearch(searchField)
{
	var field = searchField;
	field.lookupbrowse = 1;
	var f="<fetch mapping='logical'><entity name='contact'><filter>" 
	+"<condition attribute='parentcustomerid' operator='eq' value='"+id+"'/>"
	+"</filter></entity></fetch>";
	field.AddParam("search",f);
}

var customer=crmForm.all.customerid;
if(!IsNull(customer) && !IsNull(customer.DataValue))
{
	if (customer.DataValue[0].typename == 'account')
	{
		setSearch(crmForm.all.gar_main_contact);
		setSearch(crmForm.all.gar_who_to_call);
		setSearch(crmForm.all.gar_interested);
		setSearch(crmForm.all.gar_uninterested);
		setSearch(crmForm.all.gar_decided);
		setSearch(crmForm.all.gar_decided2);
		setSearch(crmForm.all.gar_decided3);
		setSearch(crmForm.all.gar_decided4);
		setSearch(crmForm.all.gar_decided5);
		setSearch(crmForm.all.gar_who_looked);
		setSearch(crmForm.all.gar_otvetstvenny_signing_1);
		setSearch(crmForm.all.gar_otvetstvenny_signing_2);
		setSearch(crmForm.all.gar_otvetstvenny_signing_3);
	}
}



//С форума - про закрытие возможной сделки (event.Mode == 5 - сохранение со статусом "Закрыта"):
if (event.Mode == 5)
{
	var status = crmFormSubmit.crNewState.value;
	var xml = crmFormSubmit.crActivityXml.value;
	var XmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	XmlDoc.async = false;
	XmlDoc.loadXML(xml);
	//If status is won
	if (status == 1)
	{
		var descriptionnode = XmlDoc.selectSingleNode("//opportunityclose/description");
		if (descriptionnode == null || descriptionnode.nodeTypedValue == "")
		{
			alert("Opportunity can't be closed!\nFill description field in Close Opportunity Dialogue");
			event.returnValue = false;
			return false;
		}
	}else if (status == 2)
	{
		var competitornode = XmlDoc.selectSingleNode("//opportunityclose/competitorid");
		if (competitornode == null || competitornode.nodeTypedValue == "")
		{
			alert("Opportunity can't be closed!\nFill competitor field in Close Opportunity Dialogue");
			event.returnValue = false;
			return false;
		}
	}
}



//Обновление параметров объекта в базе:
if (crmForm.all.customer != null && crmForm.all.customer.DataValue != null)
{
	var BP = crmForm.all.customer.DataValue;
	//Обновляем объект:
	var currentDate = new Date();
	var currentDatePlus6 = new Date(currentDate.getYear(),currentDate.getMonth()+6,currentDate.getDate(),currentDate.getHours(),currentDate.getMinutes());
	var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
	var beAccountToUpdate = new BusinessEntity("account");
	beAccountToUpdate.attributes["accountid"] = BP[0].id;
	beAccountToUpdate.attributes["gar_date_closed"] = currentDate.toFormattedString("mm/dd/yyyy hh:MM APM");
	beAccountToUpdate.attributes["gar_verification_date"] = currentDatePlus6.toFormattedString("mm/dd/yyyy hh:MM APM");
	beAccountToUpdate.attributes["gar_resultjob"] = 13;
	beAccountToUpdate.attributes["gar_account"] = BP[0].id;	//Если надо обнулить лукап-поле, то = [];
	oService.Update(beAccountToUpdate);
}



//Как программно менять видимость и уровень необходимости поля:
crmForm.all.price.RequiredLevel = '0';
crmForm.all.price_c.style.display = "none";
crmForm.all.price_d.style.display = "none";
//Правда RequiredLevel (судя по описанием некоторых людей) работать не должно, скорее это просто скрытие * с формы.
//_c, _d - это соответственно лэйбл и само поле.



//Скрытие строки формы CRM:
var field = crmForm.all.name;	//Поле, которое надо скрыть.
while ((field.parentNode != null) && (field.tagName.toLowerCase() != "tr"))	//Поиск строки таблицы HTML, в которую входит искомое поле.
{
	field = field.parentNode;
}
if (field.tagName.toLowerCase() == "tr")	//Если строка найдена, то скрываем ее.
{
	field.style.display = "none";
}



//Установка бизнес-требования для поля:
//Вы можете обратиться к уровню требований к заполнению поля при использовании свойства RequiredLevel:
var reqLevel = crmForm.all.your_Field.RequiredLevel;
//Однако, поскольку свойство только для чтения, Вы не можете использовать его, чтобы изменить необходимый уровень заполнения.
//Есть недокументированный (и поэтому неподдерживаемый) метод объекта crmForm, позволяющий сделать поле обязательным или не обязательным для заполнения:
function SetFieldReqLevel(sField, bRequired);
//Если bRequired установлен в ноль (ложь), то поле свободно для заполнения. Если во что-нибудь другое (т.е. истина), то поле обязательно для заполнения.
//Например так:
crmForm.SetFieldReqLevel("new_partnerid", 0);
//Заметьте: Вы не можете использовать этот метод, чтобы сделать поле рекомендуемым для заполнения.

//Вот код, чтобы установить поле в любое из возможных состояний:
//Не обязательно:
crmForm.all.<имя_поля>.setAttribute("req", 0);
crmForm.all.<имя_поля>_c.className = "n";
//Рекомендуемо:
crmForm.all.<имя_поля>.setAttribute("req", 1);
crmForm.all.<имя_поля>_c.className = "rec";
//Обязательно:
crmForm.all.<имя_поля>.setAttribute("req", 2);
crmForm.all.<имя_поля>_c.className = "req";



//В файле Customizations:
//Системный и пользовательский скрипт на обработчике событий OnChange:
                        <row>
                          <cell auto="false" showlabel="true" locklevel="0" rowspan="1" colspan="2" id="{483c7a9b-6544-4ac7-8225-4889527e5132}">
                            <labels>
                              <label description="Бизнес-партнер" languagecode="1033" />
                              <label description="Бизнес-партнер" languagecode="1049" />
                              <label description="Kunde" languagecode="1031" />
                            </labels>
                            <events>
                              <event name="onchange" application="true" active="true">
                                <script><![CDATA[
																	/* customerid has changed clear both contract and contractdetail */
																	crmForm.all.contractid.Clear();
																	crmForm.all.contractdetailid.Clear();
																	/* set contractdetailid to not required because contract is null*/
																	crmForm.SetFieldReqLevel("contractdetailid", false);
																]]></script>
                                <dependencies>
                                  <dependency id="contractid" />
                                  <dependency id="contractdetailid" />
                                </dependencies>
                              </event>
                              <event name="onchange" application="false" active="true">
                                <script><![CDATA[//Пользовательский скрипт.
									]]></script>
                                <dependencies />
                              </event>
                            </events>



//Про getElementById и про то, как скрыть элемент навигации (неработающие примеры).
var table = document.getElementById('idWalkTable');
document.getElementById('Layer1').scrollTop = 50;
document.getElementById("elementId").style;

//Работающий пример:
//Скрываем неправильный пункт "Обращение" из навигации на форме Контракта:
document.all.navCases.style.display ="none";
//Скрываем ненужный пункт "Копировать контракт" из меню Actions на форме Контракта:
document.all._MIclone.style.display ="none";



//Пример, как посмотреть все содержимое документа как объекта JavaScript:
//Вывести все свойства указанного объекта (здесь document.all, можно document или иное) окнами по 25 штук:
var num = 0;
var col = 0;
var resMessage = "";
for (var key in document.all)
{
	num++;
	col++;
	resMessage = resMessage + key+':'+document.all[key] + "\n";
	if (col == 25)
	{
		alert("num = "+num+"\n"+resMessage);
		col = 0;
		resMessage = "";
	}
	//alert("num = "+num+"\n"+key+':'+document.all[key]);	//Вывод по одному.
}
alert("num = "+num+"\n"+resMessage);	//Вывести те, которые не уложились в 25.



/*
Объект style элемента (document.getElementById("elementId").style), содержит только те значения, 
которые были заданны явно в атрибуте style в тэге элемента или были предварительно назначены через скрипт. 
Если Вы задаёте CSS свойства через тэг <STYLE></STYLE> или внешние листы стилей, то они не будут присутствовать в объекте style элемента.
*/

//Формы на странице:
alert(document.forms.length) //Получаем общее количество форм на странице
alert(document.forms[0].name) //Узнаем имя первой формы через массив forms
alert(document.forms.data.length) //Определяем количество элементов в форме с именем data
alert(document.forms["data"].length) //То же самое

//Узнаем имя первой формы через массив forms (рабочий пример):
alert('document.forms[0].name = '+document.forms[0].name+'\n'+
	'document.forms[0].length = '+document.forms[0].length);

alert('Всего элементов в коллекции all: '+document.all.length);
for (var i = 0; i < document.all.length; i++)
{
	if (i < 2)
	{
		alert(
			i+'\n'+
			'document.all[i].value = '+document.all[i].value+'\n'+
			'document.all[i].name = '+document.all[i].name+'\n'+
			'document.all[i].type = '+document.all[i].type+'\n'+
			'document.all[i].style = '+document.all[i].style+'\n'+
			'document.all[i].id = '+document.all[i].id+'\n'+
			//'document.all[i].class = '+document.all[i].class+'\n'+
			'document.all[i].onclick = '+document.all[i].onclick+'\n'+
			'document.all[i].title = '+document.all[i].title+'\n'+
			//'document.all[i].style = '+document.all[i].style+'\n'+
			//'document.all[i].style = '+document.all[i].style+'\n'+
			''
		);
	}
}

//Удаление узла DOM (удалить себя из родительского):
list.removeChild(elem);
//или, если неизвестен родитель:
elem.parentNode.removeChild(elem);



//Условный оператор:
return (dd < 10)? "0" + dd : dd;
access = (age > 14) ? true : false;



//Код из примера в файле про создание кнопки:

//Контракт.Клиент - customerid
//Контракт.Наша компания - gar_account
//Контракт.Обновить Нашу компанию в Клиенте (КНОПКА) - gar_account_updatebtn
//Клиент.Наша компания - gar_account

//Задаём свойства будущей кнопки:
crmForm.all.gar_account_updatebtn.DataValue = "Обновить Нашу компанию в Клиенте";
crmForm.all.gar_account_updatebtn.style.textAlign = "center";
crmForm.all.gar_account_updatebtn.style.vAlign = "middle";
crmForm.all.gar_account_updatebtn.style.cursor = "hand";
crmForm.all.gar_account_updatebtn.style.borderColor = "#330066";
crmForm.all.gar_account_updatebtn.style.backgroundColor = "#CADFFC";
crmForm.all.gar_account_updatebtn.style.color = "#000000";
crmForm.all.gar_account_updatebtn.contentEditable = false;
crmForm.all.gar_account_updatebtn.attachEvent("onmousedown",changeB1);
crmForm.all.gar_account_updatebtn.attachEvent("onmouseup",changeB2);
crmForm.all.gar_account_updatebtn.attachEvent("onmouseover",changeB3);
crmForm.all.gar_account_updatebtn.attachEvent("onmouseleave",changeB4);
crmForm.all.gar_account_updatebtn.attachEvent("onclick",doChange);
//Задаем функции обработки клиентских событий:
function changeB1() {
	crmForm.all.gar_account_updatebtn.style.color = "#000099";
}
function changeB2() {
	crmForm.all.gar_account_updatebtn.style.color = "#000000";
}
function changeB3() {
	crmForm.all.gar_account_updatebtn.style.backgroundColor = "#6699FF";
}
function changeB4() {
	crmForm.all.gar_account_updatebtn.style.backgroundColor = "#CADFFC";
}
function doChange() {
	//Проверка, есть ли значение в поле Клиент.
	if (crmForm.all.customerid != null && crmForm.all.customerid.DataValue != null)
	{
		//Если есть значение в поле Клиент, то:
		//Получаем ID Нашей компании из текущей формы Контракта:
		var ourCompany = crmForm.all.gar_account.DataValue;
		var ourCompanyID = ourCompany[0].id;
		//Получаем ID Клиента из текущей формы Контракта:
		var client = crmForm.all.customerid.DataValue;
		var clientID = client[0].id;
		//Изменяем Нашу компанию в Клиенте:
		var beUpdatedAccount = new BusinessEntity("account");
		beUpdatedAccount.attributes["accountid"] = clientID;
		beUpdatedAccount.attributes["gar_account"] = ourCompanyID;	//Если надо обнулить лукап-поле, то = [];
		oService.Update(beUpdatedAccount);
		//Убираем с формы сработавшую кнопку:
		crmForm.all.gar_account_updatebtn_c.style.display ="none";
		crmForm.all.gar_account_updatebtn_d.style.display ="none";
		//Фокус на поле Наша компания:
		crmForm.all.gar_account.focus();
	}else
	{
		alert("Нет значения в поле Клиент!");
		//Фокус на поле Клиент:
		crmForm.all.customerid.focus();
	}
}



//Фокус на поле и переместить в конец поля:
//Так изменяет данные в поле, в результате при закрытии форма запрашивает подтверждение:
crmForm.all.name.focus();
crmForm.all.name.value = crmForm.all.name.value;
//А так не изменяет данные, поэтому подтверждение не запрашивается:
var r = crmForm.all.name.createTextRange();
r.collapse(false);
r.select();



//От А. Короткова: Как повесить обработчик на битовое поле "Скидка" в Контракте, чтобы он работал:
var el = document.getElementById("usediscountaspercentage");
el.removeAttribute("disabled");
el.removeAttribute("DoNotSubmit");
function onChangeUsediscount()
{
	alert ("hello");
}
var ch1 = document.getElementById("rad_usediscountaspercentage");
var ch2 = document.getElementById("rad_usediscountaspercentage2");
ch1.onclick = onChangeUsediscount;
ch2.onclick = onChangeUsediscount;



//Получение списка отношений из объекта, id которого указан в fetch-строке, и его обработка:
//Выбрать все СПС из раздела Типы СПС в БП, перенести в текстовое поле Типы СПС через точку с запятой.
var linkFetch = ""+
	"<fetch mapping='logical'>"+
		"<entity name='gar_sps'>"+
			"<attribute name='gar_name'/>"+
			"<order attribute='gar_name' descending='false'/>"+
			"<link-entity name='gar_account_gar_sps' from='gar_spsid' to='gar_spsid' visible='false' intersect='true'>"+
				"<link-entity name='account' from='accountid' to='accountid' alias='aa'>"+
					"<filter type='and'>"+
						"<condition attribute='accountid' operator='eq' value='"+customer.DataValue[0].id+"'/>"+
					"</filter>"+
				"</link-entity>"+
			"</link-entity>"+
		"</entity>"+
	"</fetch>";
var result = oService.Fetch(linkFetch);
var inputingString = "";
for (var i = 0; i < result.length; i++)
{
	if (i > 0)
	{
		inputingString = inputingString + "; ";
	}
	var newPartOfString = result[i].attributes["gar_name"].value;	//Собственно значение нужного элемента fetch-строки результата запроса.
	inputingString = inputingString + newPartOfString;
}
crmForm.all.gar_sps.DataValue = inputingString;



//Создание связей между новым Составом комплекта и Блоками из старого Состава комплекта:
//Выбрать все Блоки из раздела Блоки в старом Составе комплекта, и связать их отношением с новым Составом комплекта.
var productFetch = ""+
	"<fetch mapping='logical'>"+
		"<entity name='product'>"+
			"<attribute name='productid'/>"+
			"<attribute name='name'/>"+
			"<attribute name='stockweight'/>"+
			"<order attribute='name' descending='false'/>"+
			"<link-entity name='gar_gar_kit_product' from='productid' to='productid' visible='false' intersect='true'>"+
				"<link-entity name='gar_kit' from='gar_kitid' to='gar_kitid'>"+
					"<filter type='and'>"+
						"<condition attribute='gar_kitid' operator='eq' value='"+idGar_kitFetch+"'/>"+
					"</filter>"+
				"</link-entity>"+
			"</link-entity>"+
		"</entity>"+
	"</fetch>";
var resultProductFetch = oService.Fetch(productFetch);
for (var j = 0; j < resultProductFetch.length; j++)
{
	oService.Associate("gar_kit", newGar_kitId, "product", resultProductFetch[j].attributes["productid"].value, "gar_gar_kit_product");
}
//Конец создания связей между новым Составом комплекта и Блоками из старого Состава комплекта.



//Если нет значения в поле Звонка "Направление" - показать диалоговое окно:
f = crmForm.all.directioncode;
if (!f.DataValue)
{
	showModelessDialog('/CallTimer/timer.htm', window, "dialogHeight:240px;dialogWidth:340px;dialogTop:0px;dialogLeft:0px;");
}



//В зависимости от значения поля включить другое поле:
var bIsPriceOverride = crmForm.all.ispriceoverridden.DataValue;
crmForm.all.priceperunit.Disabled = !bIsPriceOverride;
crmForm.SetFieldReqLevel("priceperunit",bIsPriceOverride);



//Замена окна выбора Контактов:

//Для отображения Контактов только текущего Бизнес-партнера передаем ID текущего Бизнес-партнера как дополнительный параметр "addParam" в строке запроса:
var addParam = "";
//Если в поле В отношении есть значение и это Бизнес-партнер - сформировать дополнительный параметр, взяв ID текущего Бизнес-партнера:
if (crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == 'account')
{
	//Строка параметра содержит имя и собственно ID. Имя "accountID" не изменять - оно используется в плагине FetchChangingForExecuteEventWithParameterAccountID,...
	//...который добавляет фильтрацию по заданному Бизнес-партнеру в фетч-запрос:
	addParam = "&accountID="+crmForm.all.regardingobjectid.DataValue[0].id;
}

//Код замены окна выбора Контактов на поле "Клиенты":
var nnId = "customers"; //Поле Клиенты.
var lookupTypeCode = 2; //Код сущности Контакт.
var lookupSrc = "/" + ORG_UNIQUE_NAME + "/ISV/lookup/lookupmultiCopyWithoutPropAndNew.aspx";
var lookupArg = "/" + ORG_UNIQUE_NAME + "/_root/homepage.aspx?etc=" + lookupTypeCode +"&viewid=" + lookupViewId;
var navId = document.getElementById(nnId);
if (navId != null)
{
	navId.onclick = CustomLookup;
}

function CustomLookup()
{
	var lookupItems = window.showModalDialog(lookupSrc, lookupArg, "dialogWidth:800px; dialogHeight:600px;");
	if (lookupItems)  //This is the CRM internal JS funciton on \_static\_grid\action.js
	{
		crmForm.all.customers.DataValue = lookupItems.items;
	}
}



//Работа с файлами:
//Create the FileSystemObject:
var objFSO = new ActiveXObject("Scripting.FileSystemObject");
//Create, open, close and delete the files:
var objCreatedFile = objFSO.CreateTextFile("c:\\88888888\\HowToDemoFile.txt", true);
var ForReading = 1, ForWriting = 2, ForAppending = 8;
var objOpenedFile = objFSO.OpenTextFile("c:\\88888888\\HowToDemoFile2.txt", ForWriting, true);
objCreatedFile.Close();
objOpenedFile.Close();
objFSO.DeleteFile("c:\\88888888\\HowToDemoFile.txt");
objFSO.DeleteFile("c:\\88888888\\HowToDemoFile2.txt");
//Write and read:
var objTextFile = objFSO.CreateTextFile("c:\\HowToDemoFile.txt", true);
objTextFile.WriteLine("This line is written using WriteLine().");
objTextFile.WriteBlankLines(3);
objTextFile.Write("This line is written using Write().");
objTextFile.Close();
//Use different methods to read contents of file:
objTextFile = objFSO.OpenTextFile("c:\\HowToDemoFile.txt", ForReading);
var sReadLine = objTextFile.ReadLine();
var sRead = objTextFile.Read(4);
var sReadAll = objTextFile.ReadAll();
objTextFile.Close();
//Перенос и копирование:
objFSO.MoveFile("c:\\HowToDemoFile.txt", "c:\\Temp\\");
objFSO.CopyFile("c:\\Temp\\HowToDemoFile.txt", "c:\\");
