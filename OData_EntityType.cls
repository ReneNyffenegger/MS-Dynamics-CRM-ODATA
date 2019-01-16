'
' vi: ft=basic
'
option explicit

private entityType_ as msxml2.IXMLDOMElement

public  svc_  as OData_Service

sub init(svc as OData_Service, name as string) ' {

    set svc_  = svc

    set entityType_ = svc_.metaData.selectSingleNode("//EntityType[@Name='" & name & "']")

end sub ' }


function properties as collection ' {

    set properties = new collection

    if not entityType_ is nothing then ' {

       dim entityTypeProperties_ as msxml2.IXMLDOMSelection
       set entityTypeProperties_ = entityType_.selectNodes("./Property")

       dim prop as IXMLDOMElement
       for each prop in entityTypeProperties_
           dim odataProperty   as  OData_Property
           set odataProperty = new OData_Property
           call odataProperty.init(me.svc_, prop)
           properties.add odataProperty
       next prop

    end if ' }

end function ' }
