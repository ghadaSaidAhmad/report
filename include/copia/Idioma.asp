<%'Este fichero se tiene que #incluir despues de hacer RecoverSession. 
'Si no, no tenemos el Idioma

dim Idioma
Idioma = session("Idioma")
if Idioma="" then
	'Idioma por defecto.
	Idioma = "ES"	'ES --> Español
					'EN --> English
end if


dim IDM_NoAccess
dim IDM_MAINTITLE1, IDM_MAINTITLE2
dim IDM_ConfigParametersTitle
dim IDM_Activity, IDM_GeneralTheme, IDM_GenericTheming, IDM_TemaDeCliente
dim IDM_MenuConfig, IDM_MenuAdminParameters, IDM_MenuAdminUsers, IDM_MenuAdminClientsBrands, IDM_MenuSaveParameters, IDM_MenuGoToReport, IDM_MenuFilterReport, IDM_MenuImprimir, IDM_MenuExportar, IDM_MenuInputData
dim IDM_Tematica, IDM_Nombre, IDM_LastUpdatedBy, IDM_LastUpdatedDate
dim IDM_NuevaTematica, IDM_ModificarTematica, IDM_BorrarTematica, IDM_SeleccioneTematica, IDM_GuardeTemaParaAgregarImg
dim IDM_Quincena, IDM_1aQuincena, IDM_2aQuincena
dim IDM_PrevMonth, IDM_NextMonth, IDM_PrevYear, IDM_NextYear
dim IDM_JS_ListKeepPressedCtrl
dim IDM_Personalizado, IDM_FilterNext, IDM_FilterPrevious
dim IDM_FilterPest02, IDM_FilterPest05, IDM_FilterPest1, IDM_FilterPest2
dim IDM_FilterSaltoCada
dim IDM_FilterTopBarTitle, IDM_FilterTitle, IDM_FilterSubTitle, IDM_FilterStart, IDM_FilterMonths, IDM_FilterLastYear, IDM_FilterReport, IDM_FilterReportType0, IDM_FilterReportType1, IDM_FilterReportType0Todas, IDM_FilterReportType1Todas
dim IDM_FilterSelectAll, IDM_FilterUnselectAll, IDM_SelectClient, IDM_SelectBrand
dim IDM_Client, IDM_Clients, IDM_Brand, IDM_Brands, IDM_ActivityType, IDM_Type
dim IDM_JS_SelectQuickReport, IDM_JS_SelectClient, IDM_JS_SelectSomeBrand, IDM_JS_SelectBrand, IDM_JS_SelectSomeClient, IDM_JS_SelectSomeActivityType
dim IDM_FilterClose, IDM_FilterApply, IDM_ExportExcel
dim IDM_JS_NombreTematica, IDM_JS_TemaYaExiste, IDM_JS_SeleccioneTematica, IDM_SELECT_TemasDe
dim IDM_JS_ClickOKCancel, IDM_Close, IDM_Save, IDM_Search, IDM_Delete, IDM_Cancel, IDM_Copy, IDM_CopyAlt, IDM_Paste, IDM_PasteAlt, IDM_Clear, IDM_ClearAlt
dim IDM_UserGroupListTitle, IDM_User, IDM_Users, IDM_Groups, IDM_Group, IDM_NewUser, IDM_NewGroup, IDM_UserGroupListErr30, IDM_UserGroupListErr40
dim FLD_OpcionLista, FLD_OpcionEdit, FLD_OpcionBorrar, FLD_OpcionCopiar, IDM_FLD_IDEmpleado, IDM_FLD_NombreEmpleado, IDM_FLD_EMail
dim IDM_FLD_IDGroup, IDM_FLD_Description, IDM_FLD_Observations
dim IDM_SelectUser, IDM_UserNewEditTitle, IDM_Idioma, IDM_AddToGroup, IDM_GroupsAssTo, IDM_NoRecordsFound, IDM_SelectGroup
dim IDM_UserNewEditErr10, IDM_UserNewEditErr20
dim IDM_GroupNewEditTitle, IDM_WriteGroupDescr, IDM_GroupDescr, IDM_Observations, IDM_UsersAssTo, IDM_DeleteUserFrom
dim IDM_ListPage, IDM_ListPageOf, IDM_ListPageRecords, IDM_ListFirst, IDM_ListPrevious, IDM_ListNext, IDM_ListLast
dim IDM_indBaja, IDM_Image, IDM_ExtraInfo, IDM_RemoveImage, IDM_JS_NombreObligatorio, IDM_JS_RellenarAlgunCampo
dim IDM_JS_DatosModificadosGuardar, IDM_JS_DatosModificadosGuardarCambiar
dim IDM_JS_RealData_DatosModificadosSalirSinGuardar, IDM_JS_RealData_ErrorEnValor
dim IDM_ActivityChangeClient, IDM_ActivityChangeBrand, IDM_ActivitySelectClient, IDM_ActivitySelectBrand
dim IDM_ActivityNoChange
dim IDM_ClientBrandListTitle, IDM_NewClient, IDM_NewBrand, IDM_ClientNewEditTitle
dim IDM_Name, IDM_ShortName, IDM_PlanTo, IDM_Deleted, IDM_Orden
dim IDM_BrandNewEditTitle, IDM_BrandCode
dim IDM_CalidadExp, IDM_CalidadOf
dim IDM_SOAUpdated

dim IDM_Oferta, IDM_Ratio, IDM_Folleto, IDM_Cabecera, IDM_NTiendas, IDM_NTiendasShort, IDM_NTiendasReal, IDM_NTiendasRealShort, IDM_PercentComplaint, IDM_PercentComplaintShort, IDM_Status, IDM_Adicional

select case Idioma
	case "ES"
		
		IDM_NoAccess = "No tiene acceso para acceder a esta aplicación"
		IDM_MAINTITLE1 = "SOA Reporting"
		IDM_MAINTITLE2 = "Trade Marketing Analysis"
		
		IDM_Activity = "Actividad"
		IDM_GeneralTheme = "Temática General"
		
		'Menu ------------------------------------------------------------------------
        IDM_MenuConfig = "Administración"
        IDM_MenuInputData = "Entrar Núm. tiendas y % cumplimiento"
        IDM_MenuAdminUsers = "Administrar Usuarios"
        IDM_MenuAdminClientsBrands = "Administrar Clientes/Marcas"
        IDM_MenuGoToReport = "Ir al report"
        IDM_MenuAdminParameters = "Parámetros de configuración"
        IDM_MenuSaveParameters = "Guardar parámetros"
        IDM_MenuFilterReport = "Generar nuevo report"
        IDM_MenuImprimir = "Imprimir Report"
        IDM_MenuExportar = "Exportar Report"

        IDM_Tematica = "Temática"
        IDM_Quincena = "Quincena"
        IDM_1aQuincena = "1ª Quincena"
        IDM_2aQuincena = "2ª Quincena"
        
        IDM_PrevMonth = "Mes Anterior"
        IDM_NextMonth = "Mes Siguiente"
        IDM_PrevYear = "Año Anterior"
        IDM_NextYear = "Año Siguiente"
        
        IDM_JS_ListKeepPressedCtrl = "Mantenga presionada la tecla Ctrl (Control) para seleccionar múltiples elementos"
        
		'FORMULARIO
        IDM_Nombre = "Nombre"
        IDM_LastUpdatedBy = "Actualizado por"
        IDM_LastUpdatedDate = "Actualización"
        IDM_NuevaTematica = "Nueva Temática"
        IDM_ModificarTematica = "Modificar Temática"
        IDM_BorrarTematica = "Borrar Temática"
        IDM_SeleccioneTematica = "    Seleccione una Temática"
        IDM_GuardeTemaParaAgregarImg = "Guarde el tema para seleccionar una imagen"
        IDM_Image = "Imagen"
        IDM_indBaja = "Borrado"
        IDM_ExtraInfo = "Más info."
        IDM_RemoveImage = "Quitar Imagen"
        IDM_GenericTheming = "General"
        IDM_SELECT_TemasDe = "Temáticas de: "
        IDM_TemaDeCliente = "Temática de Cliente"

        IDM_Oferta = "Oferta"
        IDM_Ratio = "Impacto en MS%"
        IDM_Folleto = "Gama en folleto"
        IDM_Cabecera = "Cabecera"
        IDM_NTiendas = "Nº centros con exposición"
        IDM_NTiendasShort = "Centros"
        IDM_NTiendasReal = "Nº Centros Real"
        IDM_NTiendasRealShort = "Cen.Real"
        IDM_PercentComplaint = "% cumplimiento exposición"
        IDM_PercentComplaintShort = "%Cumpl"
        IDM_Status = "Estado"
        IDM_Adicional = "Adicional"
        IDM_ActivityChangeClient = "Cambiar Cliente"
        IDM_ActivityChangeBrand = "Cambiar Marca"
        IDM_ActivitySelectClient = "Seleccione un Cliente"
        IDM_ActivitySelectBrand = "Seleccione una Marca"
        IDM_ActivityNoChange = "No Cambiar"
        IDM_CalidadExp = "Calidad Exposición"
        IDM_CalidadOf = "Calidad Oferta"

        IDM_JS_NombreObligatorio = "Debe rellenar el campo Nombre"
        IDM_JS_RellenarAlgunCampo = "Debe rellenar alguno de los campos de descripción"
        IDM_JS_DatosModificadosGuardar = "Los datos se han modificado. \n\rDesea guardarlos antes de cerrar?"
        IDM_JS_DatosModificadosGuardarCambiar = "Los datos se han modificado. \n\rDesea guardarlos antes de cambiar?"
        IDM_JS_NombreTematica = "Por favor, escriba el nombre de la nueva temática"
        IDM_JS_SeleccioneTematica = "Por favor, seleccione una temática o escriba en el cuadro Más info."
        IDM_JS_TemaYaExiste = "Atención, el tema ya existe\n\rSe ha seleccionado en la lista"
        
        IDM_JS_RealData_DatosModificadosSalirSinGuardar = "Ha modificado los datos.\n\rQuiere salir sin guardar?"
        IDM_JS_RealData_ErrorEnValor = "Se ha encontrado un error en un valor.\n\rPor favor, revíselo y guarde de nuevo."

        'FILTRO
        IDM_Personalizado = "Personalizado"
        IDM_FilterNext = "Siguiente >>"
        IDM_FilterPrevious = "<< Anterior"
        IDM_FilterPest02 = "Express"
        IDM_FilterPest05 = "Vista"
        IDM_FilterPest1 = "Organización"
        IDM_FilterPest2 = "Datos"
        IDM_FilterClose = "Cerrar"
        IDM_ExportExcel = "Exportar XL"
        IDM_FilterApply = "Aplicar"
        IDM_FilterTopBarTitle = "Generador de SOA Reports"
        IDM_FilterTitle = "Asistente Generador de Reports"
        IDM_FilterSubTitle = "Siga los pasos de este asistente para obtener los datos con los que desea trabajar"
        IDM_FilterStart = "Empezar en"
        IDM_FilterMonths = "Número de Meses"
        IDM_FilterLastYear = "Año Anterior"
        IDM_FilterSaltoCada = "Elementos por página"
        IDM_FilterReport = "Vista inicial"
        IDM_FilterReportType0 = "1 Cliente - Varias Marcas"
        IDM_FilterReportType1 = "1 Marca - Varios Clientes"
        IDM_FilterReportType0Todas = "1 Cliente - Todas las Marcas"
        IDM_FilterReportType1Todas = "1 Marca - Todos los Clientes"
        IDM_FilterSelectAll = "Marcar Todo"
        IDM_FilterUnselectAll = "Quitar Todo"
        IDM_SelectClient = "Seleccionar Cliente"
        IDM_SelectBrand = "Seleccionar Marca"
        IDM_Client = "Cliente"
        IDM_Clients = "Clientes"
        IDM_Brand = "Marca"
        IDM_Brands = "Marcas"
        IDM_ActivityType = "Tipo Actividad"
        IDM_Type = "Tipo"
        
        IDM_JS_SelectQuickReport = "Por favor, seleccione un tipo de report"
        IDM_JS_SelectClient = "Por favor, seleccione un cliente de la lista"
        IDM_JS_SelectSomeBrand = "Por favor, seleccione una o más marcas de la lista"
        IDM_JS_SelectBrand = "Por favor, seleccione una marca de la lista"
        IDM_JS_SelectSomeClient = "Por favor, seleccione uno o más clientes de la lista"
        IDM_JS_SelectSomeActivityType = "Por favor, seleccione algún tipo de actividad"
        
        
        ' CONFIGURATION
        IDM_ConfigParametersTitle = "Parámetros de configuración"
        
        
        IDM_JS_ClickOKCancel = "Pulse OK para continuar"
        IDM_Close = "Cerrar"
        IDM_Save = "Guardar"
        IDM_Delete = "Borrar"
        IDM_Search = "Buscar"
        IDM_Cancel = "Cancelar"
        IDM_Copy = "Copiar"
        IDM_CopyAlt = "Copiar al Portapapeles"
        IDM_Paste = "Pegar"
        IDM_PasteAlt = "Pegar desde el Portapapeles"
        IDM_Clear = "Limpiar"
        IDM_ClearAlt = "Limpiar Portapapeles"
        
        
        ' LIST
        IDM_ListPage = "Página"
        IDM_ListPageOf = "de"
        IDM_ListPageRecords = "registros"
        IDM_ListFirst = "Primera"
        IDM_ListPrevious = "Anterior"
        IDM_ListNext = "Próxima"
        IDM_ListLast = "Última"
        
        ' USERS GROUPS
        IDM_UserGroupListTitle = "Usuarios y grupos"
        IDM_User = "Usuario"
        IDM_Users = "Usuarios"
        IDM_Group = "Grupo"
        IDM_Groups = "Grupos"
        IDM_NewUser = "Nuevo Usuario"
        IDM_NewGroup = "Nuevo Grupo"
        IDM_UserGroupListErr30 = "No puede borrar el group porque tiene usuarios asignados"
        IDM_UserGroupListErr40 = "No puede borrar este grupo"
        FLD_OpcionLista = "Opciones"
        FLD_OpcionEdit = "Editar"
        FLD_OpcionBorrar = "Borrar"
        FLD_OpcionCopiar = "Copiar"
        IDM_FLD_IDEmpleado = "IDEmpleado"
        IDM_FLD_NombreEmpleado = "Nombre"
        IDM_FLD_EMail = "EMail"
        IDM_FLD_IDGroup = "IDGrupo"
        IDM_FLD_Description = "Descripción"
        IDM_FLD_Observations = "Observaciones"
        
        ' USER NEW/EDIT
        IDM_UserNewEditTitle = "Usuario"
        IDM_Idioma = "Idioma"
        IDM_UserNewEditErr10 = "El usuario ya existe"
        IDM_UserNewEditErr20 = "El empleado no existe en la base de datos de empleados"
        IDM_SelectUser = "Seleccione un usuario"
        IDM_AddToGroup = "Añadir al grupo"
        IDM_GroupsAssTo = "Grupos asignados"
        IDM_SelectGroup = "Seleccione primero un grupo"
        
        
        ' GROUP NEW/EDIT
        IDM_GroupNewEditTitle = "Grupo"
        IDM_WriteGroupDescr = "Escriba una descripción"
        IDM_GroupDescr = "Descripción"
        IDM_Observations = "Observaciones"
        IDM_UsersAssTo = "Usuarios asignados"
        IDM_DeleteUserFrom = "Quitar del grupo al usuario"
        
        
        ' CLIENT / BRAND LIST
        IDM_ClientBrandListTitle = "Clientes / Marcas"
        IDM_NewClient = "Nuevo Cliente"
        IDM_NewBrand = "Nueva Marca"
        
        ' CLIENT EDIT
        IDM_ClientNewEditTitle = "Cliente"
        IDM_Name = "Nombre"
        IDM_ShortName = "Nombre Corto XL"
        IDM_PlanTo = "PlanTo"
        IDM_Orden = "Orden"
        IDM_Deleted = "Borrado"
        
        ' BRAND EDIT
        IDM_BrandNewEditTitle = "Marca"
        IDM_BrandCode = "Nombre JDE"
        
        IDM_SOAUpdated = "Generado desde SOA Online"
        
	case "EN"

		IDM_NoAccess = "Sorry, you have no access to this application"

		IDM_MAINTITLE1 = "SOA Reporting"
		IDM_MAINTITLE2 = "Trade Marketing Analysis"

		IDM_Activity = "Activity"
		IDM_GeneralTheme = "General Theme"

		'Menu ------------------------------------------------------------------------
        IDM_MenuConfig = "Administration"
        IDM_MenuInputData = "Input NShops & %Complaint"
        IDM_MenuAdminUsers = "Admin Users"
        IDM_MenuAdminClientsBrands = "Administrar Clients/Brands"
        IDM_MenuGoToReport = "Go to Report"
        IDM_MenuAdminParameters = "Configuration Parameters"
        IDM_MenuSaveParameters = "Save Parameters"
        IDM_MenuFilterReport = "Report Generator"
        IDM_MenuImprimir = "Print Report"
        IDM_MenuExportar = "Export Report"

        IDM_Tematica = "Theme"
		IDM_Quincena = "Half"
        IDM_1aQuincena = "1st Half"
        IDM_2aQuincena = "2nd Half"
		
        IDM_PrevMonth = "Previous Month"
        IDM_NextMonth = "Next Month"
        IDM_PrevYear = "Previous Year"
        IDM_NextYear = "Next Year"

        IDM_JS_ListKeepPressedCtrl = "Please keep pressed the Ctrl (Control) key to select multiple entries"

		'FORMULARIO
        IDM_Nombre = "Name"
        IDM_LastUpdatedBy = "Last Updated"
        IDM_LastUpdatedDate = "Updated Date"
        IDM_NuevaTematica = "New Theme"
        IDM_ModificarTematica = "Modify Theme"
        IDM_BorrarTematica = "Delete Theme"
        IDM_SeleccioneTematica = "    Select Theme"
        IDM_GuardeTemaParaAgregarImg = "Save theme to attach an image"
        IDM_Image = "Image"
        IDM_indBaja = "Deleted"
        IDM_ExtraInfo = "Extra info."
        IDM_RemoveImage = "Remove Image"
        IDM_GenericTheming = "General"
        IDM_SELECT_TemasDe = "Themes for: "
        IDM_TemaDeCliente = "Client Theme"

        IDM_Oferta = "Offer"
        IDM_Ratio = "MS% Impact"
        IDM_Folleto = "Gama en folleto"
        IDM_Cabecera = "Header"
        IDM_NTiendas = "# Shops"
        IDM_NTiendasShort = "Shops"
        IDM_NTiendasReal = "# Shops Real"
        IDM_NTiendasRealShort = "Sh.Real"
        IDM_PercentComplaint = "% Complaint"
        IDM_PercentComplaintShort = "%Cumpl"
        IDM_Status = "Status"
        IDM_Adicional = "Additional"
        IDM_ActivityChangeClient = "Change Client"
        IDM_ActivityChangeBrand = "Change Brand"
        IDM_ActivitySelectClient = "Select a Client"
        IDM_ActivitySelectBrand = "Select a Brand"
        IDM_ActivityNoChange = "Don't Change"
        IDM_CalidadExp = "Exposition Quality"
        IDM_CalidadOf = "Offer Quality"

        IDM_JS_NombreObligatorio = "The field Name is mandatory"
        IDM_JS_RellenarAlgunCampo = "Please, write in some of the description fields"
        IDM_JS_DatosModificadosGuardar = "Data has been modified. \n\rDo you want to save before closing?"
        IDM_JS_DatosModificadosGuardarCambiar = "Data has been modified. \n\rDo you want to save before changing?"
        IDM_JS_NombreTematica = "Please, write the name of the new theme"
        IDM_JS_SeleccioneTematica = "Please, select a theme or write in Extra Info. box"
        IDM_JS_TemaYaExiste = "Atention, the theme already exists\n\rIt has been selected from the list"

        IDM_JS_RealData_DatosModificadosSalirSinGuardar = "Data has been modified.\n\rContinue without saving?"
        IDM_JS_RealData_ErrorEnValor = ""

        'FILTRO
        IDM_Personalizado = "Personalized"
        IDM_FilterNext = "Next >>"
        IDM_FilterPrevious = "<< Previous"
        IDM_FilterPest02 = "Quick"
        IDM_FilterPest05 = "View"
        IDM_FilterPest1 = "Organization"
        IDM_FilterPest2 = "Data"
        IDM_FilterClose = "Close"
        IDM_ExportExcel = "Excel Export"
        IDM_FilterApply = "Apply"
        IDM_FilterTopBarTitle = "SOA Reports Generator"
        IDM_FilterTitle = "Reports Generator Wizard"
        IDM_FilterSubTitle = "Follow these steps to get the working data"
        IDM_FilterStart = "Start"
        IDM_FilterMonths = "View Months"
        IDM_FilterLastYear = "Last Year"
        IDM_FilterSaltoCada = "Elements in page"
        IDM_FilterReport = "Default View"
        IDM_FilterReportType0 = "1 Client - Many Brands"
        IDM_FilterReportType1 = "1 Brand - Many Clients"
        IDM_FilterReportType0Todas = "1 Client - All Brands"
        IDM_FilterReportType1Todas = "1 Brand - All Clients"
        IDM_FilterSelectAll = "Select All"
        IDM_FilterUnselectAll = "Unselect All"
        IDM_SelectClient = "Select Client"
        IDM_SelectBrand = "Select Brand"
        IDM_Client = "Client"
        IDM_Clients = "Clients"
        IDM_Brand = "Brand"
        IDM_Brands = "Brands"
        IDM_ActivityType = "Activity Type"
        IDM_Type = "Type"

        IDM_JS_SelectQuickReport = "Please, select a report type"
        IDM_JS_SelectClient = "Please, select a client from the list"
        IDM_JS_SelectSomeBrand = "Please, select one or more brands in the list"
        IDM_JS_SelectBrand = "Please, select a brand from the list"
        IDM_JS_SelectSomeClient = "Please, select one or more clients in the list"
        IDM_JS_SelectSomeActivityType = "Please, selecte some activity type"

        ' CONFIGURATION
        IDM_ConfigParametersTitle = "Configuration Parameters"
        
        
        IDM_JS_ClickOKCancel = "Click OK to Continue"
        IDM_Close = "Close"
        IDM_Save = "Save"
        IDM_Delete = "Delete"
        IDM_Search = "Search"
        IDM_Cancel = "Cancel"
        IDM_Copy = "Copy"
        IDM_CopyAlt = "Copy to Clipboard"
        IDM_Paste = "Paste"
        IDM_PasteAlt = "Paste from Clipboard"
        IDM_Clear = "Clear"
        IDM_ClearAlt = "Clear Clipboard"
        

        ' LIST
        IDM_ListPage = "Page"
        IDM_ListPageOf = "of"
        IDM_ListPageRecords = "records"
        IDM_ListFirst = "First"
        IDM_ListPrevious = "Previous"
        IDM_ListNext = "Next"
        IDM_ListLast = "Last"

        ' USERS GROUPS
        IDM_UserGroupListTitle = "Users and groups"
        IDM_User = "User"
        IDM_Users = "Users"
        IDM_Group = "Group"
        IDM_Groups = "Groups"
        IDM_NewUser = "New User"
        IDM_NewGroup = "New Group"
        IDM_UserGroupListErr30 = "Cannot delete this group because there are users assigned"
        IDM_UserGroupListErr40 = "Cannot delete this group"
        FLD_OpcionLista = "Options"
        FLD_OpcionEdit = "Edit"
        FLD_OpcionBorrar = "Delete"
        FLD_OpcionCopiar = "Copy"
        IDM_FLD_IDEmpleado = "IDEmpleado"
        IDM_FLD_NombreEmpleado = "Nombre"
        IDM_FLD_EMail = "EMail"
        IDM_FLD_IDGroup = "IDGroup"
        IDM_FLD_Description = "Description"
        IDM_FLD_Observations = "Observations"

        ' USER NEW/EDIT
        IDM_UserNewEditTitle = "User"
        IDM_Idioma = "Language"
        IDM_UserNewEditErr10 = "User already exists"
        IDM_UserNewEditErr20 = "Not found in Employers database"
        IDM_SelectUser = "Select a user"
        IDM_AddToGroup = "Add to group"
        IDM_GroupsAssTo = "Assigned groups"
        IDM_SelectGroup = "Select a group and click again"

        ' GROUP NEW/EDIT
        IDM_GroupNewEditTitle = "Group"
        IDM_WriteGroupDescr = "Please, write a description"
        IDM_GroupDescr = "Description"
        IDM_Observations = "Observaciones"
        IDM_UsersAssTo = "Assigned users"
        IDM_DeleteUserFrom = "Remove user"

        ' CLIENT / BRAND LIST
        IDM_ClientBrandListTitle = "Clients / Brands"
        IDM_NewClient = "New Client"
        IDM_NewBrand = "New Brand"

        ' CLIENT EDIT
        IDM_ClientNewEditTitle = "Client"
        IDM_Name = "Name"
        IDM_ShortName = "Short Name XL"
        IDM_PlanTo = "PlanTo"
        IDM_Orden = "Order"
        IDM_Deleted = "Borrado"

        ' BRAND EDIT
        IDM_BrandNewEditTitle = "Brand"
        IDM_BrandCode = "JDE Name"

        IDM_SOAUpdated = "Generated from SOA Online"
        
end select

%>

