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



//��������� �� ������ ��� ��������� �� 01.03.2013 "������������ ���� ��� ���������� ������" (http://wss30/sd/Lists/Posts/Post.aspx?ID=8200).
//
//����� ������:
//��� �������������� �������������� � ������ ������������ ��������� ����. ��� ���� ���� �� ����������, ���� ������� � ���� ������������� �������� � ����������. ...
//...������ �������� �������� � ���� �������������, � ������� � �������� ����. �� ������ ��������� ��� ���������.
//
//������� ������:
//������� ����������� ���������� ���� - � ������������� ������ ����������� ������� ���� "� ���������" (crmForm.all.regardingobjectid.FireOnChange()), ������� ���������� ��� ����� ��������������...
//...(����� ����� ����������� � ����������� �����), ������ ���������� � ���������� � ��� ����� ����� ����� ��������������.
//
//������� ���� ������� � ������� � �� OnLoad, � �� OnChange �� ������������, �. �. ��� ������������� ��������� ���������� ����.
//
//������� ������ ������ ����������� regardingobjectid.FireOnChange �� OnLoad ����������� �������� ����� �������� � ���� OnLoad. ...
//...��� ��������� ������������ ���� ��� ����� �������� � ���� ������� add_filters_and_fill_fields(deliveredOService), ������� ����� �������� � OnLoad,...
//...� ����������� ��� �� OnLoad, ��� � �� regardingobjectid.OnChange, ��� ���� ��� ����� ��������� ����������� regardingobjectid.onchange.
//
//������ �� regardingobjectid.OnChange ������ ���� ������, ��������� ��� ��� ����� ���������� ����������� regardingobjectid.onchange � OnLoad.

//������� �� ���������/������ �������� � ����������/��������� �����, ����������� �� OnLoad � �� regardingobjectid.OnChange:
function add_filters_and_fill_fields(deliveredOService)
{
	//���� �� ����� ���� ������� ���� � �����-������ ������ � ���, ��: 1. ���������� ������� �� ����, 2. ��������� ����������� ���� ���������������� ����������.
	//���� ���, �� ����� ������� � �������� ��������������� ����.
	if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == 'account')
	//...� ���� ������ ������� � ��������� ������-��������, �. �. � ���� � ��������� ����� ������-�������.
	{
		//��������� ������� �� ���� LookUp.
		//������������� ������ �� ���� ������� � ����������� �� �������� ���� � ���������.
		var field = crmForm.all.gar_contact;
		//��������� ���� ������ � ���������� ���� ������:
		field.lookupbrowse = 1;
		//�������� fetch xml ����� �������� ������ ������:
		field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
		+"<condition attribute='parentcustomerid' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
		+"</filter></entity></fetch>");
		
		//��������� ������� �� ���� LookUp.
		//������������� ������ �� ���� ������� ������ � ����������� �� �������� ���� � ���������.
		var field = crmForm.all.gar_history_task;
		//��������� ���� ������ � ���������� ���� ������:
		field.lookupbrowse = 1;
		//�������� fetch xml ����� �������� ������ ������:
		field.AddParam("search","<fetch mapping='logical'><entity name='gar_history'><filter>"
		+"<condition attribute='gar_history_bp' operator='eq' value='"+ crmForm.all.regardingobjectid.DataValue[0].id +"' />"
		+"</filter></entity></fetch>");
		
		//��������� ������ ��������� �� �������, ����������� � ���� LookUp, � �� ����� � ���� �������� �����.
		//�������� �� ������� (� ���������) ��������: ���������, ������������ (����� ��������� � ����������); ������������� ��.
		var iDataSourceObject = new Array;
		iDataSourceObject = crmForm.all.regardingobjectid.DataValue;
		var iDataSourceObjectId = iDataSourceObject[0].id;
		var iCols = ["ownerid", "gar_systemuser_fo"];
		var beRetrievedDataSourceObject = deliveredOService.Retrieve("account", iDataSourceObjectId, iCols);
		
		if ("ownerid" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
		{
			//���� ���� �������� �������� � ���� LookUp:
			//������� �������� ����� �������� ������� � ������ ��� � ������ ����:
			var o1 = new Object();
			o1.id =  beRetrievedDataSourceObject.attributes["ownerid"].value;
			o1.typename = 'systemuser';
			o1.name =  beRetrievedDataSourceObject.attributes["ownerid"].name;
			crmForm.gar_worker.DataValue = [o1];
			
			//���� ��������� ������� � ����, �� ������� ��� ������������ � ������� � ������ ����:
			if (crmForm.gar_worker.DataValue != null)
			{
				var iCols2 = ["parentsystemuserid"];
				var beRetrievedDataSourceObject2 = deliveredOService.Retrieve("systemuser", o1.id, iCols2);
				
				if ("parentsystemuserid" in beRetrievedDataSourceObject2.attributes)	//���� ������� ���� � != null.
				{
					//���� ���� �������� �������� � ���� LookUp:
					//������� �������� ����� �������� ������� � ������ ��� � ������ ����:
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
		
		if ("gar_systemuser_fo" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
		{
			//���� ���� �������� �������� � ���� LookUp:
			//������� �������� ����� �������� ������� � ������ ��� � ������ ����:
			var o3 = new Object();
			o3.id =  beRetrievedDataSourceObject.attributes["gar_systemuser_fo"].value;
			o3.typename = 'systemuser';
			o3.name =  beRetrievedDataSourceObject.attributes["gar_systemuser_fo"].name;
			crmForm.gar_systemuser_fo_task.DataValue = [o3];
		}else
		{
			crmForm.gar_systemuser_fo_task.DataValue = null;
		}
		//����� ��������� ������ ��������� � �� ������ � ���� �������� �����.
	}else
	{
		crmForm.all.gar_contact.DataValue = null;	//��� ������ �� �� �������� �������.
		crmForm.all.gar_fax_number.DataValue = null;	//��� ������ �� �� �������� ����� �����.
		crmForm.all.gar_worker.DataValue = null;	//��� ������ �� �� �������� ����������.
		crmForm.all.gar_head.DataValue = null;	//��� ������ �� �� �������� ������������.
		crmForm.all.gar_systemuser_fo_task.DataValue = null;	//��� ������ �� �� �������� �������������� ��.
		
		//��� ������ �� �� ������� ������ � ��������:
		var field = crmForm.all.gar_contact;
		//��������� ���� ������ � ���������� ���� ������:
		field.lookupbrowse = 1;
		//�������� fetch xml ����� �������� ������ ������:
		field.AddParam("search","<fetch mapping='logical'><entity name='contact'><filter>"
		+"</filter></entity></fetch>");
		
		//��� ������ �� �� ������� ������ � ������� �����:
		var field = crmForm.all.gar_history_task;
		//��������� ���� ������ � ���������� ���� ������:
		field.lookupbrowse = 1;
		//�������� fetch xml ����� �������� ������ ������:
		field.AddParam("search","<fetch mapping='logical'><entity name='gar_history'><filter>"
		+"</filter></entity></fetch>");
	}
	//����� ��������, ���� �� �� ����� ������� ���� � ������ � ���, � ���������� ����������� ����� ���������������� ����������.
}

//��������� ������� �� ���������/������ �������� � ����������/��������� ����� ������ (�� OnLoad):
add_filters_and_fill_fields(oService);

//�������, ������� ������� �������� �� ���������� ���� "� ���������" (regardingobjectid.OnChange). �������� ������� �� ���������/������ �������� � ����������/��������� �����, + ��������� �������������� ��������:
function add_filters_and_fill_fields_forOnChange()
{
	//��������� ������� �� ���������/������ �������� � ����������/��������� ����� �� regardingobjectid.OnChange:	
	add_filters_and_fill_fields(oService);
	
	//��������� �����-���� �������������� ��������:
	//alert('add_filters_and_fill_fields_forOnChange');
}

//�������� ������� �� ���������� ���� "� ���������" (regardingobjectid.OnChange):
crmForm.all.regardingobjectid.onchange = add_filters_and_fill_fields_forOnChange;

//����� ��������� �� ������ ��� ��������� �� 01.03.2013 "������������ ���� ��� ���������� ������" (http://wss30/sd/Lists/Posts/Post.aspx?ID=8200).



//���������� ������������� ������� �� ��������������
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



//���, ����������� � ������������� ���������� �������� �� ������ ����������:
if (crmForm.FormType == 2)	//���� �������������� ������������, �� ����������� ������.
{
	if (crmForm.all.regardingobjectid != null && crmForm.all.regardingobjectid.DataValue != null && crmForm.all.regardingobjectid.DataValue[0].typename == "account")
	{
		var bp = crmForm.all.regardingobjectid.DataValue;
		
		var iCols = ["name"];
		var beRetrievedDataSourceObject = oService.Retrieve("account", bp[0].id, iCols);	//�� �������� id �� ����� ��� �� �� ����.
		if ("name" in beRetrievedDataSourceObject.attributes)	//���� ������� ���� � != null.
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
				'������������ �� � ���� "� ���������" �� ���������� �����.\n'+
				'������ ������������: '+bp[0].name+'\n'+
				'����� ������������: '+o.name+'\n'+
				'��������� ����� ����� ���������.\n'+
				''
			);
			*/
		}
	}
}



//��������� �� ������ "���� � ������" (http://portal/service_applications/sdesk/_layouts/15/start.aspx#/Lists/Posts/Post.aspx?ID=1077) �� 28.07.2015 (��� ��������).
//���� ���� "���������� �������" (gar_do) ������ � ������. ���� ���� �������� - ��������� �������������� (������� � ���������� �����, ��� �� �����).
loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');
var IAm = WhoAmI(oService);	//��� ������ ���� ������� ����������: loadScript('/ISV/SomeOtherScripts/GeneralFunctions.js');

var departments = {
	'���': '',
	'��� VIP': '',
	'��� ��': '',
	'��� ���': '',
	'��� ��': '',
	'��� ����': '',
	'��� ��': '',
	'��� ��': '',
	'��� ���': ''
};

var businessunitName = IAm.attributes["businessunitid"].name;	//�������� ��� ������������� ������������ �� ������� WhoAmI (���� �� ����).
if (!(businessunitName in departments))	//���� ������������� �������� ������������ - �� ���:
{
	if (crmForm.all.gar_do != null)	//���� ���� "���������� �������" �� ����� ����.
	{
		if (crmForm.all.gar_do.DataValue == null)	//���� ���� �� ����� ����, �� ��� ������.
		{
			//������ ��������� ���� �� �����:
			crmForm.all.gar_do_c.style.display = "none";
			crmForm.all.gar_do_d.style.display = "none";
		}
		else	//���� ���� ��������.
		{
			//������ ����������� ��� �������������� ���� �� �����:
			crmForm.all.gar_do.disabled = true;
			crmForm.all.gar_do.readOnly = true;
		}
	}
}
//����� ��������� �� ������ "���� � ������" (http://portal/service_applications/sdesk/_layouts/15/start.aspx#/Lists/Posts/Post.aspx?ID=1077) �� 28.07.2015 (��� ��������).