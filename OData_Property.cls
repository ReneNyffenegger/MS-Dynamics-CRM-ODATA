'
' vi: ft=basic
'
option explicit

private property_ as msxml2.IXMLDOMElement
private svc_         as OData_Service

sub init(svc as OData_Service, p as msxml2.IXMLDOMElement) ' {
    set svc_        = svc
    set property_   = p
end sub ' }

function name as string ' {
 '
 '  The edm:Property element MUST include a Name attribute whose value is a SimpleIdentifier
 ' 
    name = property_.getAttribute("Name")
end function ' }

function type_ as string ' {
 '
 '  The edm:Property element MUST include a Type attribute.
 '  The value of the Type attribute MUST be the
 '    o  QualifiedName of a primitive type,
 '    o  complex type, or
 '    o  enumeration type in scope, or 
 '    o  a collection of one of these types.
 '
    type_ = property_.getAttribute("Type")
end function ' }

function description as string ' {

   dim xmlAnnotation as msxml2.IXMLDOMElement
   set xmlAnnotation = property_.selectSingleNode("./Annotation[@Term='Org.OData.Core.V1.Description']")
   

   if xmlAnnotation is nothing then
      description = ""
      exit function
   end if

   description = xmlAnnotation.getAttribute("String")

end function ' }
