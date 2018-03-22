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
dim IDM_AdminFormsTitle, IDM_FormName, IDM_Form, IDM_QualityForm, IDM_AdminFormTitle, IDM_MenuSaveForm, IDM_NewForm, IDM_AddQuestion, IDM_Question, IDM_Weight, IDM_RespType, IDM_TotalWeight, IDM_CurrentTotalW, IDM_FromDate, IDM_FormReadOnly
dim IDM_DeleteQuestion, IDM_ModifyQuestion, IDM_CreateResponse, IDM_ResponseValue, IDM_ResponseText, IDM_NoResponseYet, IDM_DeleteResponse, IDM_SaveResponse, IDM_PestQuestions, IDM_PestAssignBrands, IDM_AssignedBrands, IDM_RemoveFromForm, IDM_NoBrandsAssigned, IDM_PromotionsAssigned, IDM_PromotionsAssignedEx, IDM_RemoveHistory, IDM_NoPromotionsAssigned, IDM_BrandsWithoutForm, IDM_AddBrandToForm, IDM_NoPendingBrands, IDM_BrandsAssOtherForm, IDM_NoBrandsAssOtherForm
dim IDM_JS_QuestTextWeight, IDM_JS_QuestWeight, IDM_JS_RespTextValue, IDM_JS_DeleteFormMessage, IDM_JS_RemoveFormFromBrand, IDM_JS_DeleteHistory, IDM_JS_ReassignForm, IDM_JS_DeleteHistoryBrand, IDM_JS_ChangeForm
dim IDM_Activity, IDM_GeneralTheme, IDM_GenericTheming, IDM_TemaDeCliente, IDM_KPIQuality
dim IDM_MenuConfig, IDM_MenuAdminParameters, IDM_MenuAdminUsers, IDM_MenuAdminClientsBrands, IDM_MenuAdminForms, IDM_MenuSaveParameters, IDM_MenuGoToReport, IDM_MenuFilterReport, IDM_MenuImprimir, IDM_MenuExportar, IDM_MenuInputData
dim IDM_Tematica, IDM_Nombre, IDM_LastUpdatedBy, IDM_LastUpdatedDate
dim IDM_NuevaTematica, IDM_ModificarTematica, IDM_BorrarTematica, IDM_SeleccioneTematica, IDM_GuardeTemaParaAgregarImg
dim IDM_Quincena, IDM_1aQuincena, IDM_2aQuincena
dim IDM_PrevMonth, IDM_NextMonth, IDM_PrevYear, IDM_NextYear
dim IDM_JS_ListKeepPressedCtrl
dim IDM_Personalizado, IDM_FilterNext, IDM_FilterPrevious
dim IDM_FilterPest02, IDM_FilterPest05, IDM_FilterPest1, IDM_FilterPest2, IDM_FilterPest30
dim IDM_FilterSaltoCada
dim IDM_FilterTopBarTitle, IDM_FilterTitle, IDM_FilterSubTitle, IDM_FilterStart, IDM_FilterMonths, IDM_FilterLastYear, IDM_FilterReport, IDM_FilterReportType0, IDM_FilterReportType1, IDM_FilterReportType0Todas, IDM_FilterReportType1Todas
dim IDM_FilterSelectAll, IDM_FilterUnselectAll, IDM_SelectClient, IDM_SelectBrand
dim IDM_Client, IDM_Clients, IDM_Brand, IDM_Brands, IDM_ActivityType, IDM_Type
dim IDM_BtnReportQuery, IDM_BtnReportQueryTxt, IDM_BtnReportQuery_Option, IDM_BtnReportQuery_AlertJS
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
dim IDM_indBaja, IDM_Image, IDM_ExtraInfo, IDM_RemoveImage, IDM_JS_NombreObligatorio, IDM_JS_RellenarAlgunCampo, IDM_JS_RellenarFormulario
dim IDM_JS_DatosModificadosGuardar, IDM_JS_DatosModificadosGuardarCambiar
dim IDM_JS_RealData_DatosModificadosSalirSinGuardar, IDM_JS_RealData_ErrorEnValor
dim IDM_ActivityChangeClient, IDM_ActivityChangeBrand, IDM_ActivitySelectClient, IDM_ActivitySelectBrand
dim IDM_ActivityNoChange
dim IDM_ClientBrandListTitle, IDM_NewClient, IDM_NewBrand, IDM_ClientNewEditTitle
dim IDM_Name, IDM_ShortName, IDM_PlanTo, IDM_Deleted, IDM_Orden, IDM_ClientFormsActivated
dim IDM_BrandNewEditTitle, IDM_BrandCode
dim IDM_CalidadExp, IDM_CalidadOf
dim IDM_SOAUpdated

dim IDM_Oferta, IDM_Ratio, IDM_Folleto, IDM_Cabecera, IDM_NTiendas, IDM_NTiendasGeneral, IDM_NTiendasShort, IDM_NTiendasReal, IDM_NTiendasTOTAL, IDM_NTiendasRealShort, IDM_Status, IDM_Adicional
dim IDM_Subcategory, IDM_JS_ConfirmDeleteSubcat, IDM_AutoFillSubcategories


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
        IDM_MenuAdminForms = "Administrar Formularios"
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
        IDM_KPIQuality = "KPI Calidad"

        IDM_Oferta = "Oferta"
        IDM_Ratio = "Impacto en MS%"
        IDM_Folleto = "Gama en folleto"
        IDM_Cabecera = "Cabecera"
        IDM_NTiendasGeneral = "GENERAL"
        IDM_NTiendas = "NºCentros HQ"
        IDM_NTiendasShort = "HQ"
        IDM_NTiendasReal = "NºCentros GPV"
        IDM_NTiendasTOTAL = "Total Centros"
        IDM_NTiendasRealShort = "GPV"
        IDM_Status = "Estado"
        IDM_Adicional = "Adicional"
        IDM_ActivityChangeClient = "Cambiar Cliente"
        IDM_ActivityChangeBrand = "Cambiar Marca"
        IDM_ActivitySelectClient = "Seleccione un Cliente"
        IDM_ActivitySelectBrand = "Seleccione una Marca"
        IDM_ActivityNoChange = "No Cambiar"
        IDM_CalidadExp = "Calidad Exposición"
        IDM_CalidadOf = "Calidad Oferta"
        IDM_AutoFillSubcategories = "Autorellenar subcategorías"

        IDM_JS_NombreObligatorio = "Debe rellenar el campo Nombre"
        IDM_JS_RellenarAlgunCampo = "Debe rellenar alguno de los campos de descripción"
        IDM_JS_RellenarFormulario = "Por favor, rellene el formulario de calidad"
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
        IDM_FilterPest30 = "Extras"
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
        IDM_BtnReportQuery = "Lanzar Query"
        IDM_BtnReportQueryTxt = ""
        IDM_BtnReportQuery_Option = "Seleccione un elemento "
        IDM_BtnReportQuery_AlertJS = "Seleccione un report de la lista"
        
        IDM_JS_SelectQuickReport = "Por favor, seleccione un tipo de report"
        IDM_JS_SelectClient = "Por favor, seleccione un cliente de la lista"
        IDM_JS_SelectSomeBrand = "Por favor, seleccione una o más marcas de la lista"
        IDM_JS_SelectBrand = "Por favor, seleccione una marca de la lista"
        IDM_JS_SelectSomeClient = "Por favor, seleccione uno o más clientes de la lista"
        IDM_JS_SelectSomeActivityType = "Por favor, seleccione algún tipo de actividad"
        
        
        ' CONFIGURATION
        IDM_ConfigParametersTitle = "Parámetros de configuración"
        
        
        ' FORMS
        IDM_AdminFormsTitle = "Administración de Formularios"
        IDM_FormName = "Nombre"
        IDM_Form = "Formulario"
        IDM_QualityForm = "Formulario de calidad"
        IDM_AdminFormTitle = "Formulario"
        IDM_FromDate = "Aplicado desde"
        IDM_MenuSaveForm = "Guardar Formulario"
        IDM_NewForm = "Nuevo Formulario"
        IDM_FormReadOnly = "No pueden añadir o quitar preguntas y respuestas al formulario porque ya ha sido usado"
        IDM_JS_DeleteFormMessage = "Esta acción borrará permanentemente del sistema el formulario y su historial.\n\rPulse OK para continuar o Cancelar para abortar."
        IDM_AddQuestion = "Añadir pregunta"
        IDM_Question = "Pregunta"
        IDM_Weight = "Peso"
        IDM_RespType = "Tipo"
        IDM_TotalWeight = "Peso total"
        IDM_CurrentTotalW = "Peso total actual"
        IDM_DeleteQuestion = "Borrar pregunta"
        IDM_ModifyQuestion = "Modificar pregunta"
        IDM_CreateResponse = "Crear respuesta"
        IDM_NoResponseYet = "Todavía no hay respuestas"
        IDM_DeleteResponse = "Borrar respuesta"
        IDM_SaveResponse = "Guardar respuesta"
        IDM_ResponseValue = "Valor"
        IDM_ResponseText = "Texto"
        IDM_PestQuestions = "Preguntas"
        IDM_PestAssignBrands = "Asignación a marcas"
        IDM_AssignedBrands = "Marcas asignadas al formulario"
        IDM_RemoveFromForm = "Eliminar esta marca del formulario"
        IDM_NoBrandsAssigned = "No hay marcas asignadas"
        IDM_PromotionsAssigned = "Promociones asignadas al formulario (histórico)"
        IDM_PromotionsAssignedEx = "Promociones que fueron asignadas a este formulario y posteriormente la marca cambió de formulario o se desasignó"
        IDM_RemoveHistory = "Eliminar la información histórica del formulario en esta marca"
        IDM_NoPromotionsAssigned = "No hay promociones asignadas al formulario para otras marcas"
        IDM_BrandsWithoutForm = "Marcas sin formulario"
        IDM_AddBrandToForm = "Añadir esta marca al formulario"
        IDM_NoPendingBrands = "No hay marcas pendientes de asignar"
        IDM_BrandsAssOtherForm = "Marcas asignadas a otro formulario"
        IDM_NoBrandsAssOtherForm = "No hay marcas asignadas a otros formularios"
        
        IDM_JS_QuestTextWeight = "Debe especificar el texto de la pregunta y su peso numérico"
        IDM_JS_QuestWeight = "Peso no válido"
        IDM_JS_RespTextValue = "Texto o valor no válido"
        IDM_JS_RemoveFormFromBrand = "Desea desasignar el formulario de la marca?"
        IDM_JS_DeleteHistory = "Desea borrar el histórico?"
        IDM_JS_ReassignForm = "Desea reasignar el formulario?"
        IDM_JS_DeleteHistoryBrand = "Desea borrar el histórico del formulario en la marca?"
        IDM_JS_ChangeForm = "La marca está asignada a otro formulario.\n\rDesea cambiarle el formulario?"
        
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
        IDM_ClientFormsActivated = "Formularios activos"
        
        ' BRAND EDIT
        IDM_BrandNewEditTitle = "Marca"
        IDM_BrandCode = "Nombre JDE"
        IDM_Subcategory = "Subcategoría"
        IDM_JS_ConfirmDeleteSubcat = "¿Quieres borrar la Subcategoría y sus datos históricos?"
        
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
        IDM_MenuAdminClientsBrands = "Manage Clients/Brands"
        IDM_MenuAdminForms = "Manage Forms"
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
        IDM_KPIQuality = "KPI Quality"

        IDM_Oferta = "Offer"
        IDM_Ratio = "MS% Impact"
        IDM_Folleto = "Gama en folleto"
        IDM_Cabecera = "Header"
        IDM_NTiendasGeneral = "GENERAL"
        IDM_NTiendas = "#Shops HQ"
        IDM_NTiendasShort = "Shops"
        IDM_NTiendasReal = "#Shops GPV"
        IDM_NTiendasTOTAL = "Total #Shops"
        IDM_NTiendasRealShort = "Sh.GPV"
        IDM_Status = "Status"
        IDM_Adicional = "Additional"
        IDM_ActivityChangeClient = "Change Client"
        IDM_ActivityChangeBrand = "Change Brand"
        IDM_ActivitySelectClient = "Select a Client"
        IDM_ActivitySelectBrand = "Select a Brand"
        IDM_ActivityNoChange = "Don't Change"
        IDM_CalidadExp = "Exposition Quality"
        IDM_CalidadOf = "Offer Quality"
        IDM_AutoFillSubcategories = "Autofill Subcategories"

        IDM_JS_NombreObligatorio = "The field Name is mandatory"
        IDM_JS_RellenarAlgunCampo = "Please, write in some of the description fields"
        IDM_JS_RellenarFormulario = "Please, fill in the quality form"
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
        IDM_FilterPest30 = "Extras"
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
        
        ' FORMS
        IDM_AdminFormsTitle = "Forms Administration"
        IDM_FormName = "Name"
        IDM_Form = "Form"
        IDM_QualityForm = "Quality Form"
        IDM_AdminFormTitle = "Form"
        IDM_FromDate = "Enabled from"
        IDM_MenuSaveForm = "Save Form"
        IDM_NewForm = "New Form"
        IDM_FormReadOnly = "You cannot add or remove questions and responses from the form because it has been used"
        IDM_JS_DeleteFormMessage = "This action will permanently remove the form and all its historical information from the system.\n\rClick OK to continue. Click Cancel to abort."
        IDM_AddQuestion = "Add Question"
        IDM_Question = "Question"
        IDM_Weight = "Weight"
        IDM_RespType = "Type"
        IDM_TotalWeight = "Total Weight"
        IDM_CurrentTotalW = "Current total weight"
        IDM_DeleteQuestion = "Delete Question"
        IDM_ModifyQuestion = "Modify question"
        IDM_CreateResponse = "Create response"
        IDM_NoResponseYet = "No responses yet"
        IDM_DeleteResponse = "Delete response"
        IDM_SaveResponse = "Save response"
        IDM_ResponseText = "Text"
        IDM_ResponseValue = "Value"
        IDM_PestQuestions = "Questions"
        IDM_PestAssignBrands = "Brands assignation"
        IDM_AssignedBrands = "Brands assigned to the form"
        IDM_RemoveFromForm = "Remove the brand from the form"
        IDM_NoBrandsAssigned = "No brands assigned"
        IDM_PromotionsAssigned = "Promotions assigned to the form (history)"
        IDM_PromotionsAssignedEx = "Promotions that were assigned to this form and after that the brand changed of form or was unassigned"
        IDM_RemoveHistory = "Remove the history of the form in this brand"
        IDM_NoPromotionsAssigned = "No promotinos assigned to the form for other brands"
        IDM_BrandsWithoutForm = "Brands without form"
        IDM_AddBrandToForm = "Add this brand to the form"
        IDM_NoPendingBrands = "No brands pending to assign"
        IDM_BrandsAssOtherForm = "Brands assigned to an other form"
        IDM_NoBrandsAssOtherForm = "No brands assigned to other forms"

        IDM_JS_QuestTextWeight = "You must write the question text and its numeric weight"
        IDM_JS_QuestWeight = "Weight not valid"
        IDM_JS_RespTextValue = "Text or value not valid"
        IDM_JS_RemoveFormFromBrand = "Do you want to remove the form from the brand?"
        IDM_JS_DeleteHistory = "Do you want to delete the history?"
        IDM_JS_ReassignForm = "Do you want to reassign the form?"
        IDM_JS_DeleteHistoryBrand = "Do you want to delete the form history in the brand?"
        IDM_JS_ChangeForm = "The brand is assigned to another form.\n\rDo you want to change the form?"
        
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
        IDM_Deleted = "Deleted"
        IDM_ClientFormsActivated = "Forms activated"

        ' BRAND EDIT
        IDM_BrandNewEditTitle = "Brand"
        IDM_BrandCode = "JDE Name"
        IDM_Subcategory = "Subcategory"
        IDM_JS_ConfirmDeleteSubcat = "Do you want to delete the Subcategory and its history data?"

        IDM_SOAUpdated = "Generated from SOA Online"
        
end select

%>

