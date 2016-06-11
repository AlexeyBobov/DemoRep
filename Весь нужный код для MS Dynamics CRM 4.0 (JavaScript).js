//����������� CrmService � �������� ������� �������:
loadScript('/ISV/ascentium/CrmService.js');
function loadScript(oScriptURL)
{
	var xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
	xmlHttp.open("GET", oScriptURL, false);
	xmlHttp.send();
	window.execScript(xmlHttp.responseText);
}
var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
//����� ����������� CrmService � �������� ������� �������.



//��������, ���� �� �� ����� ������� ���� � ���� �� � ������� ���� �����-������ ������.
//���� ����, ��������� ����������� ���� ���������������� ����������.
//���� ���, �������� ��������������� ����.
if (crmForm.all.gar_contact != null && crmForm.all.gar_contact.DataValue != null)
{
	
	//��������� ������ ��������� �� �������, ����������� � ���� LookUp, � �� ����� � ���� �������� �����.
	//�������� �� ������� ��������: ����� ���������������, �������������, ����� �������������, �������� ����������� � ����� ��.
	var iDataSourceObject = new Array;
	iDataSourceObject = crmForm.all.customer.DataValue;
	var iDataSourceObjectId = iDataSourceObject[0].id;
	var iCols = ["businesstypecode", "accountclassificationcode", "ownershipcode", "gar_conditional_association"];
	var beRetrievedDataSourceObject = oService.Retrieve("account", iDataSourceObjectId, iCols);
	
	if ("businesstypecode" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		crmForm.all.gar_businesstypecode.DataValue = beRetrievedDataSourceObject.attributes["businesstypecode"].name;
	}else
	{
		crmForm.all.gar_businesstypecode.DataValue = null;
	}
	
	if ("accountclassificationcode" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		crmForm.all.gar_accountclassificationcode_bp.DataValue = beRetrievedDataSourceObject.attributes["accountclassificationcode"].name;
	}else
	{
		crmForm.all.gar_accountclassificationcode_bp.DataValue = null;
	}
	
	if ("ownershipcode" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		crmForm.all.gar_ownershipcode_bp.DataValue = beRetrievedDataSourceObject.attributes["ownershipcode"].name;
	}else
	{
		crmForm.all.gar_ownershipcode_bp.DataValue = null;
	}
	
	if ("gar_conditional_association" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		//�������� �������� �������� ���� ����� �������, ������ ��� value � DataValue �������� �� ������������� (�� ��� ���):
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
	
	if ("gar_list_phonecalls" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		//���� ���� �������� �������� � ���� LookUp:
		//���������� �������� � ���� LookUp ����� �������� ������� � ������ ��� � ������ ����:
		var o = new Object();
		o.id = beRetrievedDataSourceObject.attributes["gar_list_phonecalls"].value;
		o.typename = 'list';
		o.name = beRetrievedDataSourceObject.attributes["gar_list_phonecalls"].name;
		crmForm.gar_list.DataValue = [o];
	}else
	{
		crmForm.gar_list.DataValue = null;
	}
	
	if ("gar_stoimost_with_discounts" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		//�������� �������� ���� Money, ����� 0 ��� ���������� � ��������� ����, ������ ��� � DataValue ���� ���� Money �������� string (�� value) �������� �� ������������� (�� ��� ���):
		crmForm.all.gar_amount.DataValue = beRetrievedDataSourceObject.attributes["gar_stoimost_with_discounts"].value - 0;
	}
	
	if ("gar_activeon_fact" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
	{
		//���� ���� �������� �������� � ���� DateTime:
		//�������� �������� ���� Date, ������� ��� � ���������� �������, ������ ��� � DataValue ���� ���� Date �������� string (�� value) �������� �� ������������� (�� ��� ���):
		crmForm.all.gar_start_of.DataValue = new Date(beRetrievedDataSourceObject.attributes["gar_activeon_fact"].value.replace(/(\d+)-(\d+)-(\d+)/, '$2/$3/$1').substr(0, 10));
	}
	//����� ��������� ������ ��������� � �� ������ � ���� �������� �����.
	
}else
{
	crmForm.all.gar_post.DataValue = null;
	crmForm.all.emailaddress.DataValue = null;
}
//����� ��������, ���� �� �� ����� ������� ���� � ������ � ���, � ���������� ����������� ����� ���������������� ����������.



//���� �� ��������� �����, � �������������:
if (crmForm.FormType != 1)



//�������� ���� � �������������� ������� ��� ���������:
var today = new Date();
var choosenDayInt = parseInt(crmForm.all.activeon.DataValue.toFormattedString("yyyymmdd"));
var todayInt = parseInt(today.toFormattedString("yyyymmdd"));



//��������, ���� �� �� ����� ������� ���� � ���� �� � ������� ���� �����-������ ������.
//���� ����, �� � ����� ������: 1. ���������� ������ �� ����, 2. ��������� ����������� ���� ���������������� ����������.
//���� ���, �������� ��������������� ����.
if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null)


	
	//��������� ������� �� ���� LookUp.
	//������������� ������ �� ���� ������� � ����������� �� �������� ���� � ���������.
	var field = crmForm.all.gar_contact;
	//��������� ���� ������ � ���������� ���� ������:
	field.lookupbrowse = 1;
	//�������� fetch xml ����� �������� ������ ������:
	field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
	+"<condition attribute='parentcustomerid' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
	+"</filter></entity></fetch>");
	
	
	
	//������ ������� � ���� LookUp.
	var field = crmForm.all.gar_contact_assistant;
	field.lookupbrowse = 1;
	field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
	+"</filter></entity></fetch>");



//��� ������� ������������ ������, ��������, ������:
field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
+"<condition attribute='statecode' operator='eq' value='0' />"
+"<condition attribute='customertypecode' operator='ne' value='200000' />"
+"<condition attribute='parentcustomerid' operator='eq' value='" + receiver[0].id + 
"' /></filter></entity></fetch>");



//��������� �������� ���� ���� "�������" �� ��� ��������� ���� (attributename - ��� ����, objecttypecode - ��� ��������, attributevalue - ��� ���������� ��������):
//������ 1 - ������ �����:
	var sFetchXml = "<fetch mapping=\"logical\"><entity name=\"stringmap\"><attribute name=\"value\"/><filter type=\"and\">"
		+"<condition attribute='attributename' operator='eq' value='gar_function' />"
		+"<condition attribute='objecttypecode' operator='eq' value='2' />"
		+"<condition attribute='attributevalue' operator='eq' value='"+id+"' /></filter></entity></fetch>";
//������ 2 - �������� ���������� �����:
	//��������� �������� ���� ���� "�������" �� ��� ��������� ���� (attributename - ��� ����, objecttypecode - ��� ��������, attributevalue - ��� ���������� ��������):
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
//����� objecttypecode (��� ��������):
//1 - ������-������� (account).
//2 - ������� (contact).
//3 - ��������� ������ (opportunity).
//4 - ������� (lead).
//1010 - �������� (contract).
//1011 - ������ ��������� (contractdetail).
//4214 - �������� ������� (serviceappointment).
//10004 - ���������� ������ (gar_paysplit).
//...������ ����� ����� �� ������ �������� - Export all customizations.



//������� ���������� (��������� ��� ������) ��������� �������� ���� ��������� ������ (�) (gar_resultjob) ���� "�������":
//�� ������������� (�������� 16), �� ����� �������� ���� ����� ��� (9), ������ �� �������� (7), �������� (13), ������� �� �-� (1), �������� ��������������� (10).
var exceptionsList = {'1': '', '7': '', '9': '', '10': '', '13': '', '16': ''};	//������ ���������� (��������, ������� �� ���������� � ���������� �������� �� �����).
var rsj = crmForm.all.gar_resultjob;	//��������� ��� ���� ��� ��������� �������.
var dv = rsj.DataValue;					//��������� ������� �������� ����.
var resultArray = new Array();			//������ ��� ����������� (������ ��������, ������� ����� ����������).
var j = 0;
for (i = 0; i < rsj.Options.length; i++)	//������������� ����� ������ ��������.
{
	//�� ������ ������ �������� � ������ ��� ����������� �� ��������, ������� ��� � ������ ����������;...
	//...����� ����, �������� ������� ��������, ���� ���� ��� � ������ ���������� (��� ����� ������, ��� �����, ���� ������� �������� � ������ ����������,...
	//...�� ��� �� ����� �������� � ���� ��� �������� �������� - ���� ��������� ������).
	if (!(rsj.Options[i].DataValue in exceptionsList) || (rsj.Options[i].DataValue == dv))
	{
		resultArray[j] = rsj.Options[i];
		j++;
	}
}
rsj.Options = resultArray;				//���������� ���������� ������ ��������.
rsj.DataValue = dv;						//��������������� �������������� ��������.



//���������� ��� �������� ����� ��������� �������:
//������ crmForm.all.gar_amount.value		���� �������� ���� 13 699,80 - �. �. � ������ �������������� � ������������ ��� ��������� �������.
//������ crmForm.all.gar_amount.DataValue	���� �������� ���� 13699.8 - �. �. � ���� �����, ��� ����� �������������� � ������������ ��� ��������� �������.



//��������� ������ ��������� �� �������, ��� id ��������, ����� �������� � ����� �������:
//(������� �������, ����� ����� ���� � �� ��������������).

//������� ��������� ������ ��������� "cols" �� ������� ���� "typename" � ��������� id:
function getFields(typename, id, cols)
{
	//������� var oService = new Ascentium_CrmService, ������� ��� ��������� � ����� "����������� CrmService � �������� ������� �������"
	//var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
	var c = oService.Retrieve(typename, id, cols);
	return c.attributes;	//["name"].value
}

//����� ���� � �� ��������������:
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



������
	if ("businesstypecode" in beRetrievedDataSourceObject.attributes)
�������� ������
	if (beRetrievedDataSourceObject.attributes["businesstypecode"] != null && beRetrievedDataSourceObject.attributes["businesstypecode"].value != null)



//���� ��������� ���� "������ ������", ��������� ��� ���������:
if (crmForm.all.gar_demontratsy != null && crmForm.all.gar_demontratsy.DataValue != null)
{
    crmForm.all.gar_demontratsy.Disabled = true;
}



//switch...case...default:
	switch (t)
	{
		case '{63220E4D-69CF-DE11-8E20-00155D4E1B14}':	//���� ����==��.
			//�����-�� ���.
			break;
		case '{8B0ED757-69CF-DE11-8E20-00155D4E1B14}':	//���� ����==��������.
		case '{0854A205-59F1-DE11-8853-00155D4E1B14}':	//���� ����==�������� ��.
			//�����-�� ���.
			break;
		default:	//���� ���� - �� ���� �� ���������.
			//�����-�� ���.
			break;
	}



//������ ��������� ���� �� �����:
crmForm.all.gar_account_c.style.display = "none";
crmForm.all.gar_account_d.style.display = "none";



//������ ����������� ��� �������������� ���� �� �����:
crmForm.all.gar_name.disabled = true;
crmForm.all.gar_name.readOnly = true;



//���������� ���������� ��������:
crmForm.all.activeon.prevDataValue = CloneDate(crmForm.all.activeon.DataValue);			//������ �� ��������� OnLoad � ����� customizations.
crmForm.all.parentcustomerid.prevDataValue = crmForm.all.parentcustomerid.DataValue;	//��� ��������� (�������.OnLoad � OnSave).



//������ ������, ���������� � ����������� ���� �� ������� �����:
crmForm.all.gar_name.DataValue = crmForm.all.gar_type_agreement.SelectedText;



//������� ��������� �������� ���� ���� "�������":
cert = crmForm.all.gar_certificate;
certValue = cert.DataValue;
switch (certValue)
{
	case '3':	//���� ����������=="���������� ����".
	case '4':	//���� ����������=="������� ����".
		//���������� ������ ������ �� ����������: '' (������ ������ - �������������� ��������), '3' ("���������� ����"), '4' ("������� ����"):
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
		//������� � ���� ���������� ������ ��������:
		crmForm.all.gar_certificate.DataValue = certValue;
		break;
	case null:	//���� ����������==null (�� ������).
		//���������� ������ ������ �� ����������: '' (������ ������ - �������������� ��������), '3' ("���������� ����"), '4' ("������� ����"):
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
	default:	//���� � ���� ���������� ����������� ��������.
		alert('����������� �������� � ���� ����������!');
		break;
}



//�������� � ���� ����� �������� ������������:
loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
var IAm = WhoAmI(oService);	//��� ������ ���� ������� ����������: loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
//�������� �������� � ���� LookUp. ������� �������� ����� �������� ������� � ������ ��� � ������ ����:
var o = new Object();
o.id = IAm.attributes["systemuserid"].value;
o.typename = 'systemuser';
o.name = IAm.attributes["fullname"].value;
crmForm.all.resources.DataValue = [o];



//GUID �������� ������� (��������� � �������� ����):
var it = crmForm.ObjectId;



//� ����������� �� ����, � �������� ��� � ������� ����:
if (ORG_UNIQUE_NAME == "CmpnyLab")	{ var lookupViewId = "{6FF24503-5D6F-E211-91EA-00155D001308}"; }
if (ORG_UNIQUE_NAME == "Cmpny")		{ var lookupViewId = "{E9E0ED78-BA75-E211-9FAF-00155D001308}"; }



//��������� ��������� �����:
crmForm.all.gar_contact.outerHTML
crmForm.all.gar_contact.lookupstyle='multi';	//������ 'multi' ��� 'single'.
��������� ������ �� ��������� (�������� � ��������� outerHTML):	//<IMG style="IME-MODE: auto" id=gar_contact class=ms-crm-Lookup title="Click to select a value for ��������." alt="Click to select a value for ��������." src="/_imgs/btn_off_lookup.gif" req="0" resolveemailaddress="0" showproperty="1" autoresolve="1" defaulttype="0" lookupstyle="single" lookupbrowse="0" lookuptypeIcons="/_imgs/ico_16_2.gif" lookuptypenames="contact:2" lookuptypes="2">



//������� ������ �� ����� ������� � ������:
var t1 = document.getElementById('tab' + 1);	//������� 1 (������ � ����), ������ ������� �������.
var t2 = document.getElementById('tab' + 2);	//������� 2 (������ � ����), ���� �������� �������.
//������� 1. �������� ������ "��� ������ � ���� ����������..." �� ������� "������ � ��" (t2.childNodes[0].rows[2]) �� ������ "����������" �� ������� "������ �� ��" (t1.childNodes[0].rows[3]):
t2.childNodes[0].rows[2].parentNode.replaceChild(t1.childNodes[0].rows[3], t2.childNodes[0].rows[2]);
//������� 2. ����� ������������ ��������� ������ �� ������� "������ � ��" (t2.childNodes[0].rows[1]), ����� �� ������������ �������,...
//...� �������� � ���� ������ "����������" �� ������� "������ �� ��" (t1.childNodes[0].rows[3]) ��������� �������� ���������:
t2.childNodes[0].rows[1].parentNode.appendChild(t1.childNodes[0].rows[3]);



//�������������� �������� �����:
window.close();



//������ �� ��������� �����. ��������, ����� ������, ��������, �������� � ������:
var newWindow = window.open('kml.htm', 'displayWindow', 'width=500,height=400,status=yes,toolbar=yes,menubar=yes');
newWindow.document.write('��� ���� ����� ������� ����� 2 ������� ����� ��������� ������ ���� (����� �������� ���� ���� alert � �. �.).');
setTimeout(function(){ newWindow.close(); }, 2000);

newWindow.close();			//������� �������� � ����.
newWindow.document.close();	//������� ��������, �� �������� �������� ����.



//������ � ����� �������������:
if (confirm("��������� � ��������� ��� ��������� � ����?"))
{
	alert("\"��������� � ���������\" � ����������!");
}
else
{
	alert("\"��������� � ����\" � ����������!");
}



//�������� �� ������:
try
{
	var a = 5;
	var res = a(1); // ������!
}
catch(err)
{
	alert("name: " + err.name + "\nmessage: " + err.message + "\nstack: " + err.stack);
}



//������� ���� ��������� ������� �������:
var ooo = {a:5, b:true}
for (var key in ooo)
{
	alert(key+' : '+ooo[key]);
}



//
//
//����� - �� ��� �������, � �� ������ ����������:
//
//



//������� ��������� ������:
function HSSection(tabIndex, sectionIndex, displayType)
{
	var s = document.getElementById('tab' + tabIndex);
	s.childNodes[0].rows[sectionIndex].style.display = displayType;
}
HSSection(0, 2, 'none');
crmForm.all.tab1Tab.style.display = 'none';



//�������� ���� �������� � ���� LookUp:
var receiver = crmForm.all.to.DataValue;
if (receiver[0].type==1) //�. �. � ���������� ����� ������-�������.



//�������� �������������� ������� � ������ �� ���������� �������� �����:
var field = crmForm.all.regardingobjectid;
if (!IsNull(field) && !IsNull(field.DataValue))
{
    var dv = field.DataValue[0];
    if (dv.typename in {'account': '', 'contact': ''})
    {
		//���.
    }
}



//����� ����������� ������� ������ ���� �� ������� ���� ��� �� OnLoad, OnSave:
crmForm.all.businesstypecode.FireOnChange();



//���������� ����� ��������� ������ �� �� �� ���� ������������� ������:

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



//� ������ - ��� �������� ��������� ������ (event.Mode == 5 - ���������� �� �������� "�������"):
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



//���������� ���������� ������� � ����:
if (crmForm.all.customer != null && crmForm.all.customer.DataValue != null)
{
	var BP = crmForm.all.customer.DataValue;
	//��������� ������:
	var currentDate = new Date();
	var currentDatePlus6 = new Date(currentDate.getYear(),currentDate.getMonth()+6,currentDate.getDate(),currentDate.getHours(),currentDate.getMinutes());
	var oService = new Ascentium_CrmService(ORG_UNIQUE_NAME);
	var beAccountToUpdate = new BusinessEntity("account");
	beAccountToUpdate.attributes["accountid"] = BP[0].id;
	beAccountToUpdate.attributes["gar_date_closed"] = currentDate.toFormattedString("mm/dd/yyyy hh:MM APM");
	beAccountToUpdate.attributes["gar_verification_date"] = currentDatePlus6.toFormattedString("mm/dd/yyyy hh:MM APM");
	beAccountToUpdate.attributes["gar_resultjob"] = 13;
	beAccountToUpdate.attributes["gar_account"] = BP[0].id;	//���� ���� �������� �����-����, �� = [];
	oService.Update(beAccountToUpdate);
}



//��� ���������� ������ ��������� � ������� ������������� ����:
crmForm.all.price.RequiredLevel = '0';
crmForm.all.price_c.style.display = "none";
crmForm.all.price_d.style.display = "none";
//������ RequiredLevel (���� �� ��������� ��������� �����) �������� �� ������, ������ ��� ������ ������� * � �����.
//_c, _d - ��� �������������� ����� � ���� ����.



//������� ������ ����� CRM:
var field = crmForm.all.name;	//����, ������� ���� ������.
while ((field.parentNode != null) && (field.tagName.toLowerCase() != "tr"))	//����� ������ ������� HTML, � ������� ������ ������� ����.
{
	field = field.parentNode;
}
if (field.tagName.toLowerCase() == "tr")	//���� ������ �������, �� �������� ��.
{
	field.style.display = "none";
}



//��������� ������-���������� ��� ����:
//�� ������ ���������� � ������ ���������� � ���������� ���� ��� ������������� �������� RequiredLevel:
var reqLevel = crmForm.all.your_Field.RequiredLevel;
//������, ��������� �������� ������ ��� ������, �� �� ������ ������������ ���, ����� �������� ����������� ������� ����������.
//���� ������������������� (� ������� ����������������) ����� ������� crmForm, ����������� ������� ���� ������������ ��� �� ������������ ��� ����������:
function SetFieldReqLevel(sField, bRequired);
//���� bRequired ���������� � ���� (����), �� ���� �������� ��� ����������. ���� �� ���-������ ������ (�.�. ������), �� ���� ����������� ��� ����������.
//�������� ���:
crmForm.SetFieldReqLevel("new_partnerid", 0);
//��������: �� �� ������ ������������ ���� �����, ����� ������� ���� ������������� ��� ����������.

//��� ���, ����� ���������� ���� � ����� �� ��������� ���������:
//�� �����������:
crmForm.all.<���_����>.setAttribute("req", 0);
crmForm.all.<���_����>_c.className = "n";
//������������:
crmForm.all.<���_����>.setAttribute("req", 1);
crmForm.all.<���_����>_c.className = "rec";
//�����������:
crmForm.all.<���_����>.setAttribute("req", 2);
crmForm.all.<���_����>_c.className = "req";



//� ����� Customizations:
//��������� � ���������������� ������ �� ����������� ������� OnChange:
                        <row>
                          <cell auto="false" showlabel="true" locklevel="0" rowspan="1" colspan="2" id="{483c7a9b-6544-4ac7-8225-4889527e5132}">
                            <labels>
                              <label description="������-�������" languagecode="1033" />
                              <label description="������-�������" languagecode="1049" />
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
                                <script><![CDATA[//���������������� ������.
									]]></script>
                                <dependencies />
                              </event>
                            </events>



//��� getElementById � ��� ��, ��� ������ ������� ��������� (������������ �������).
var table = document.getElementById('idWalkTable');
document.getElementById('Layer1').scrollTop = 50;
document.getElementById("elementId").style;

//���������� ������:
//�������� ������������ ����� "���������" �� ��������� �� ����� ���������:
document.all.navCases.style.display ="none";
//�������� �������� ����� "���������� ��������" �� ���� Actions �� ����� ���������:
document.all._MIclone.style.display ="none";



//������, ��� ���������� ��� ���������� ��������� ��� ������� JavaScript:
//������� ��� �������� ���������� ������� (����� document.all, ����� document ��� ����) ������ �� 25 ����:
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
	//alert("num = "+num+"\n"+key+':'+document.all[key]);	//����� �� ������.
}
alert("num = "+num+"\n"+resMessage);	//������� ��, ������� �� ��������� � 25.



/*
������ style �������� (document.getElementById("elementId").style), �������� ������ �� ��������, 
������� ���� ������� ���� � �������� style � ���� �������� ��� ���� �������������� ��������� ����� ������. 
���� �� ������ CSS �������� ����� ��� <STYLE></STYLE> ��� ������� ����� ������, �� ��� �� ����� �������������� � ������� style ��������.
*/

//����� �� ��������:
alert(document.forms.length) //�������� ����� ���������� ���� �� ��������
alert(document.forms[0].name) //������ ��� ������ ����� ����� ������ forms
alert(document.forms.data.length) //���������� ���������� ��������� � ����� � ������ data
alert(document.forms["data"].length) //�� �� �����

//������ ��� ������ ����� ����� ������ forms (������� ������):
alert('document.forms[0].name = '+document.forms[0].name+'\n'+
	'document.forms[0].length = '+document.forms[0].length);

alert('����� ��������� � ��������� all: '+document.all.length);
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

//�������� ���� DOM (������� ���� �� �������������):
list.removeChild(elem);
//���, ���� ���������� ��������:
elem.parentNode.removeChild(elem);



//�������� ��������:
return (dd < 10)? "0" + dd : dd;
access = (age > 14) ? true : false;



//��� �� ������� � ����� ��� �������� ������:

//��������.������ - customerid
//��������.���� �������� - gar_account
//��������.�������� ���� �������� � ������� (������) - gar_account_updatebtn
//������.���� �������� - gar_account

//����� �������� ������� ������:
crmForm.all.gar_account_updatebtn.DataValue = "�������� ���� �������� � �������";
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
//������ ������� ��������� ���������� �������:
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
	//��������, ���� �� �������� � ���� ������.
	if (crmForm.all.customerid != null && crmForm.all.customerid.DataValue != null)
	{
		//���� ���� �������� � ���� ������, ��:
		//�������� ID ����� �������� �� ������� ����� ���������:
		var ourCompany = crmForm.all.gar_account.DataValue;
		var ourCompanyID = ourCompany[0].id;
		//�������� ID ������� �� ������� ����� ���������:
		var client = crmForm.all.customerid.DataValue;
		var clientID = client[0].id;
		//�������� ���� �������� � �������:
		var beUpdatedAccount = new BusinessEntity("account");
		beUpdatedAccount.attributes["accountid"] = clientID;
		beUpdatedAccount.attributes["gar_account"] = ourCompanyID;	//���� ���� �������� �����-����, �� = [];
		oService.Update(beUpdatedAccount);
		//������� � ����� ����������� ������:
		crmForm.all.gar_account_updatebtn_c.style.display ="none";
		crmForm.all.gar_account_updatebtn_d.style.display ="none";
		//����� �� ���� ���� ��������:
		crmForm.all.gar_account.focus();
	}else
	{
		alert("��� �������� � ���� ������!");
		//����� �� ���� ������:
		crmForm.all.customerid.focus();
	}
}



//����� �� ���� � ����������� � ����� ����:
//��� �������� ������ � ����, � ���������� ��� �������� ����� ����������� �������������:
crmForm.all.name.focus();
crmForm.all.name.value = crmForm.all.name.value;
//� ��� �� �������� ������, ������� ������������� �� �������������:
var r = crmForm.all.name.createTextRange();
r.collapse(false);
r.select();



//�� �. ���������: ��� �������� ���������� �� ������� ���� "������" � ���������, ����� �� �������:
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



//��������� ������ ��������� �� �������, id �������� ������ � fetch-������, � ��� ���������:
//������� ��� ��� �� ������� ���� ��� � ��, ��������� � ��������� ���� ���� ��� ����� ����� � �������.
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
	var newPartOfString = result[i].attributes["gar_name"].value;	//���������� �������� ������� �������� fetch-������ ���������� �������.
	inputingString = inputingString + newPartOfString;
}
crmForm.all.gar_sps.DataValue = inputingString;



//�������� ������ ����� ����� �������� ��������� � ������� �� ������� ������� ���������:
//������� ��� ����� �� ������� ����� � ������ ������� ���������, � ������� �� ���������� � ����� �������� ���������.
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
//����� �������� ������ ����� ����� �������� ��������� � ������� �� ������� ������� ���������.



//���� ��� �������� � ���� ������ "�����������" - �������� ���������� ����:
f = crmForm.all.directioncode;
if (!f.DataValue)
{
	showModelessDialog('/CallTimer/timer.htm', window, "dialogHeight:240px;dialogWidth:340px;dialogTop:0px;dialogLeft:0px;");
}



//� ����������� �� �������� ���� �������� ������ ����:
var bIsPriceOverride = crmForm.all.ispriceoverridden.DataValue;
crmForm.all.priceperunit.Disabled = !bIsPriceOverride;
crmForm.SetFieldReqLevel("priceperunit",bIsPriceOverride);



//������ ���� ������ ���������:

//��� ����������� ��������� ������ �������� ������-�������� �������� ID �������� ������-�������� ��� �������������� �������� "addParam" � ������ �������:
var addParam = "";
//���� � ���� � ��������� ���� �������� � ��� ������-������� - ������������ �������������� ��������, ���� ID �������� ������-��������:
if (crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == 'account')
{
	//������ ��������� �������� ��� � ���������� ID. ��� "accountID" �� �������� - ��� ������������ � ������� FetchChangingForExecuteEventWithParameterAccountID,...
	//...������� ��������� ���������� �� ��������� ������-�������� � ����-������:
	addParam = "&accountID="+crmForm.all.regardingobjectid.DataValue[0].id;
}

//��� ������ ���� ������ ��������� �� ���� "�������":
var nnId = "customers"; //���� �������.
var lookupTypeCode = 2; //��� �������� �������.
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



//������ � �������:
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
//������� � �����������:
objFSO.MoveFile("c:\\HowToDemoFile.txt", "c:\\Temp\\");
objFSO.CopyFile("c:\\Temp\\HowToDemoFile.txt", "c:\\");
