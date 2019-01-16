option explicit

const ms_dynamics_service_url_root = "https://odata-service/"

sub main() ' {

    dim entityTypeName as string
    entityTypeName = "account"

    dim odataService as new OData_Service
    odataService.init(ms_dynamics_service_url_root & "$metadata")
  
    dim odataEntityType as OData_EntityType
    set odataEntityType = odataService.EntityType(entityTypeName)

    dim f as integer
    f = freeFile

    open environ("USERPROFILE") & "\" & entityTypeName & ".properties.txt" for output as #f ' "
    
    dim prop as OData_Property
    for each prop in odataEntityType.properties ' {
        print# f,   prop.name & " - " & prop.type_
    next prop ' }

    close# f
     
end sub ' }
