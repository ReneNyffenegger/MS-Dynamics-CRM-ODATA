'
' vi: ft=basic
'
option explicit

public metadata as msxml2.DOMDocument

sub class_initialize() ' {
    set metadata = new msxml2.DOMDocument
end sub ' }

sub init(fileName_or_URL as string) ' {

    metadata.async            = false
    metadata.validateOnParse  = true
    metadata.resolveExternals = true   ' ? 
    
    if not metadata.load(fileName_or_URL) then
       debug.print(metadata.parseError.reason)
       debug.print("Line: " & metadata.parseError.line   )
       debug.print("Pos:  " & metadata.parseError.linepos)
       exit sub
    end if

end sub ' }

function EntityType(name as string) as OData_EntityType ' {

    set  EntityType = new OData_EntityType
    call EntityType.init(me, name)

end function ' }
