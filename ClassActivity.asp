<script runat=server language="vbscript">

    ' ############################################################################
    class Activity
        
        public ID
        public IDBrand
        public IDClient
        public WYear
        public WMonth
        public WHalf
        public Name
        public LastUpdatedBy
        public LastUpdatedDate
        public LastUpdatedDateDate
        
        public Oferta
        public IDRatio
        public Folleto
        public Cabecera
        public NShops
        public IDStatus
        public Adicional
        public RatioBackground
        public IDCalidadExp
        public DesCalidadExp
        public IDCalidadOf
        public DesCalidadOf
        
        public RD_NShops
        public TotalNShops
        
        public idForm
        public KPIQuality
        
        public arrNShops(9)
        public arrRD_NShops(9)
        public arrTOTALNShops(9)
        
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            IDBrand = 0
            IDClient = 0
            WYear = 1900
            WMonth = 1
            WHalf = 1
            Name = ""
            Oferta = ""
            IDRatio = 0
            Folleto = ""
            Cabecera = ""
            'NShops = 0
            IDStatus = 1
            Adicional = ""
            RatioBackground = ""
            idForm = -1
            KPIQuality = -1
            
        end sub

        
        public property get GridText
            GridText = Name
        end property
        
        ' Devuelve el nombre del BGColor
        public property get BGColor
            dim ret
            ret = ""

            if GridText<>"" then
                ret = Application("ColorActivity")
            else
                ret = ""
            end if

            BGColor = ret
        end property
        

        ' Devuelve el nombre del FGColor
        public property get FGColor
            dim ret
            ret = ""

            if GridText<>"" then
                ret = "black"
            else
                ret = ""
            end if

            FGColor = ret
        end property
        
    end class
    
    
    ' ############################################################################
    ' READ Activity data
    public function getActivity(id)
        dim SQL
        dim rst
        
        dim act
        set act = new Activity
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT act.*, " & _
        " rat.BGColor AS [RatioBG], " & _
        " act.LastUpdatedDate AS LastUpdatedDateDate, " & _
        " CONVERT(varchar, act.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, act.LastUpdatedDate, 108) AS StringUpdatedDate, " & _
        " em.ApellidosNombre AS NombreUpdated, " & _
        " rd.NShops AS RD_NShops, " & _
        " cexp.Descripcion AS DesCalidadExp, cof.Descripcion AS DesCalidadOf, " & _
        " CASE WHEN act.NShops IS NULL THEN 0 ELSE act.NShops END + CASE WHEN rd.NShops IS NULL THEN 0 ELSE rd.NShops END AS TotalNShops, " & _
        " rd.NShops0 AS RD_NShops0, rd.NShops1 AS RD_NShops1, rd.NShops2 AS RD_NShops2, rd.NShops3 AS RD_NShops3, rd.NShops4 AS RD_NShops4, rd.NShops5 AS RD_NShops5, rd.NShops6 AS RD_NShops6, rd.NShops7 AS RD_NShops7, rd.NShops8 AS RD_NShops8, rd.NShops9 AS RD_NShops9, " & _
        " CASE WHEN act.NShops0 IS NULL THEN 0 ELSE act.NShops0 END + CASE WHEN rd.NShops0 IS NULL THEN 0 ELSE rd.NShops0 END AS TotalNShops0, " & _
        " CASE WHEN act.NShops1 IS NULL THEN 0 ELSE act.NShops1 END + CASE WHEN rd.NShops1 IS NULL THEN 0 ELSE rd.NShops1 END AS TotalNShops1, " & _
        " CASE WHEN act.NShops2 IS NULL THEN 0 ELSE act.NShops2 END + CASE WHEN rd.NShops2 IS NULL THEN 0 ELSE rd.NShops2 END AS TotalNShops2, " & _
        " CASE WHEN act.NShops3 IS NULL THEN 0 ELSE act.NShops3 END + CASE WHEN rd.NShops3 IS NULL THEN 0 ELSE rd.NShops3 END AS TotalNShops3, " & _
        " CASE WHEN act.NShops4 IS NULL THEN 0 ELSE act.NShops4 END + CASE WHEN rd.NShops4 IS NULL THEN 0 ELSE rd.NShops4 END AS TotalNShops4, " & _
        " CASE WHEN act.NShops5 IS NULL THEN 0 ELSE act.NShops5 END + CASE WHEN rd.NShops5 IS NULL THEN 0 ELSE rd.NShops5 END AS TotalNShops5, " & _
        " CASE WHEN act.NShops6 IS NULL THEN 0 ELSE act.NShops6 END + CASE WHEN rd.NShops6 IS NULL THEN 0 ELSE rd.NShops6 END AS TotalNShops6, " & _
        " CASE WHEN act.NShops7 IS NULL THEN 0 ELSE act.NShops7 END + CASE WHEN rd.NShops7 IS NULL THEN 0 ELSE rd.NShops7 END AS TotalNShops7, " & _
        " CASE WHEN act.NShops8 IS NULL THEN 0 ELSE act.NShops8 END + CASE WHEN rd.NShops8 IS NULL THEN 0 ELSE rd.NShops8 END AS TotalNShops8, " & _
        " CASE WHEN act.NShops9 IS NULL THEN 0 ELSE act.NShops9 END + CASE WHEN rd.NShops9 IS NULL THEN 0 ELSE rd.NShops9 END AS TotalNShops9 " & _
        " FROM Activity act " & _
        " LEFT JOIN EmpleadosGlobal em ON act.LastUpdatedBy = em.IDEmpleado " & _
        " LEFT JOIN ActivityRatio rat ON act.IDRatio = rat.id " & _
        " LEFT JOIN RealData rd ON act.WYear = rd.WYear AND act.WMonth = rd.WMonth AND act.WHalf = rd.WHalf AND act.IDClient = rd.IDClient AND act.IDBrand = rd.IDBrand " & _
        " LEFT JOIN CalidadExp cexp ON act.IDCalidadExp = cexp.ID " & _
        " LEFT JOIN CalidadOf cof ON act.IDCalidadOf = cof.ID " & _
        " WHERE act.ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.IDBrand = rst("IDBrand")
            act.IDClient = rst("IDClient")
            act.WYear = rst("WYear")
            act.WMonth = rst("WMonth")
            act.WHalf = rst("WHalf")
            act.Name = rst("Name")
            act.LastUpdatedBy = rst("NombreUpdated")
            act.LastUpdatedDate = rst("StringUpdatedDate")
            act.LastUpdatedDateDate = rst("LastUpdatedDateDate")
            
            act.Oferta = rst("Oferta")
            act.IDRatio = rst("IDRatio")
            act.Folleto = rst("Folleto")
            act.Cabecera = rst("Cabecera")
            act.NShops = rst("NShops")
            
            act.RD_NShops = rst("RD_NShops")
            
            act.IDStatus = rst("IDStatus")
            act.Adicional = rst("Adicional")
            
            act.RatioBackground = rst("RatioBG")
            
            act.IDCalidadExp = rst("IDCalidadExp")
            act.DesCalidadExp = rst("DesCalidadExp")
            act.IDCalidadOf = rst("IDCalidadOf")
            act.DesCalidadOf = rst("DesCalidadOf")
            
            if isNull(rst("idForm")) then
                act.idForm = -1
            else
                act.idForm = rst("idForm")
            end if
            act.KPIQuality = rst("KPIQuality")
            
            dim iSubcat
            for iSubcat = 0 to 9
                act.arrNShops(iSubcat) = rst("NShops" & iSubcat)
            next
            for iSubcat = 0 to 9
                act.arrRD_NShops(iSubcat) = rst("RD_NShops" & iSubcat)
            next
            for iSubcat = 0 to 9
                act.arrTOTALNShops(iSubcat) = rst("TOTALNShops" & iSubcat)
            next
            
            
            act.TotalNShops = rst("TotalNShops")
            
        else
            Err.Raise 555, "ClassActivity", "Activity not found"
        end if
        rst.Close
        set rst = nothing

        set getActivity = act
    end function
    
    
    ' ############################################################################
    ' READ Activity data from a given date
    public function getActivityFromDate(IDClient, IDBrand, WYear, WMonth, WHalf)
        dim SQL, rst
        dim act
        dim iSubcat, idActivity
        
        idActivity = -1
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        
        SQL = "SELECT act.ID " & _
        " FROM Activity act " & _
        " WHERE act.IDClient = " & IDClient & " AND act.IDBrand = " & IDBrand & _
        " AND act.WYear = " & WYear & " AND act.WMonth = " & WMonth & " AND act.WHalf = " & WHalf & " "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            idActivity = rst("ID")
        else
            idActivity = -1
        end if
        rst.Close
        
        
        if idActivity > 0 then
            set act = getActivity(idActivity)
            
        else
            set act = new Activity
            act.IDBrand = IDBrand
            act.IDClient = IDClient
            act.WYear = WYear
            act.WMonth = WMonth
            act.WHalf = WHalf
            
            
            ' Carga Real Data
            SQL = "SELECT * FROM RealData " & _
            " WHERE IDClient = " & IDClient & " AND IDBrand = " & IDBrand & _
            " AND WYear = " & WYear & " AND WMonth = " & WMonth & " AND WHalf = " & WHalf & " "
            rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
            if not rst.EOF then
                act.RD_NShops = rst("NShops")
                
                for iSubcat = 0 to 9
                    act.arrRD_NShops(iSubcat) = rst("NShops" & iSubcat)
                next
            end if
            rst.Close
            
        end if


        set rst = nothing

        
        set getActivityFromDate = act
    end function
        
    
    ' ############################################################################
    public function getActivities(WYear, WMonth, IDClient, IDBrand)
        dim SQL
        dim rst
        dim act1, act2
        dim arrActivity(1) ' Array con una actividad por quincena
        
        
        set act1 = new Activity
        set act2 = new Activity
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, WHalf FROM Activity " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                set act1 = getActivity(rst("id"))
            elseif rst("WHalf") = 2 then
                set act2 = getActivity(rst("id"))
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrActivity(0) = act1
        set arrActivity(1) = act2
        
        getActivities = arrActivity
    end function
    

    

    ' ############################################################################
    'SAVE new or edit
    public sub saveActivity(act)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        dim sArrNShops(9)
        
        isNew = false
        
        dim sIDRatio: sIDRatio= "NULL"
        if act.IDRatio <> "" then sIDRatio = act.IDRatio
        
        act.NShops = Replace(act.NShops, ",", ".")
        if act.NShops<>"" then if NOT isNumeric(act.NShops) then act.NShops = 0
        dim sNShops: sNShops = "NULL"
        if act.NShops <> "" then sNShops = act.NShops
        
        dim iSubcat
        for iSubcat = 0 to 9
            sArrNShops(iSubcat) = "NULL"
            if act.arrNShops(iSubcat)<>"" then 
                act.arrNShops(iSubcat) = Replace(act.arrNShops(iSubcat), ",", ".")
                if NOT isNumeric(act.arrNShops(iSubcat)) then 
                    act.arrNShops(iSubcat) = 0
                end if
                sArrNShops(iSubcat) = act.arrNShops(iSubcat)
            end if
            
        next

        if act.ID > -1 then
            ' Not a new Activity
            SQL = "UPDATE Activity " & _
            " SET  " & _
            " Name = '" & replace(act.Name, "'", "''") & "'" & _
            " , Oferta = '" & replace(act.Oferta, "'", "''") & "' " & _
            " , IDRatio = " & sIDRatio & " " & _
            " , Folleto = '" & replace(act.Folleto, "'", "''") & "' " & _
            " , Cabecera = '" & replace(act.Cabecera, "'", "''") & "' " & _
            " , NShops = " & sNShops & " " & _
            " , IDStatus = '" & replace(act.IDStatus, "'", "") & "' " & _
            " , Adicional = '" & replace(act.Adicional, "'", "''") & "' " & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() " & _
            " , idForm = '" & act.idForm & "' " & _
            " , NShops0 = " & sArrNShops(0) & " " & _
            " , NShops1 = " & sArrNShops(1) & " " & _
            " , NShops2 = " & sArrNShops(2) & " " & _
            " , NShops3 = " & sArrNShops(3) & " " & _
            " , NShops4 = " & sArrNShops(4) & " " & _
            " , NShops5 = " & sArrNShops(5) & " " & _
            " , NShops6 = " & sArrNShops(6) & " " & _
            " , NShops7 = " & sArrNShops(7) & " " & _
            " , NShops8 = " & sArrNShops(8) & " " & _
            " , NShops9 = " & sArrNShops(9) & " " & _
            " WHERE id = " & act.ID
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Activity (IDBrand, IDClient, WYear " & _
            " , WMonth, WHalf, Name " & _
            " , Oferta, IDRatio " & _
            " , Folleto, Cabecera " & _
            " , NShops " & _
            " , IDStatus, Adicional " & _
            " , LastUpdatedBy, LastUpdatedDate " & _
            " , idForm " & _
            " , NShops0 " & _
            " , NShops1 " & _
            " , NShops2 " & _
            " , NShops3 " & _
            " , NShops4 " & _
            " , NShops5 " & _
            " , NShops6 " & _
            " , NShops7 " & _
            " , NShops8 " & _
            " , NShops9 " & _
            " ) " & _
            " VALUES (" & act.IDBrand & ", " & act.IDClient & ", " & act.WYear & " " & _
            " , " & act.WMonth & ", " & act.WHalf & ", '" & Replace(act.Name, "'", "''") & "' " & _
            " , '" & replace(act.Oferta, "'", "''") & "', " & sIDRatio & " " & _
            " , '" & replace(act.Folleto, "'", "''") & "', '" & replace(act.Cabecera, "'", "''") & "' " & _
            " , " & sNShops & " " & _
            " , '" & replace(act.IDStatus, "'", "") & "', '" & replace(act.Adicional, "'", "''") & "' " & _
            " , " & session("IDEmpleado") & ", GETDATE() " & _
            " , '" & act.idForm & "' " & _
            " , " & sArrNShops(0) & " " & _
            " , " & sArrNShops(1) & " " & _
            " , " & sArrNShops(2) & " " & _
            " , " & sArrNShops(3) & " " & _
            " , " & sArrNShops(4) & " " & _
            " , " & sArrNShops(5) & " " & _
            " , " & sArrNShops(6) & " " & _
            " , " & sArrNShops(7) & " " & _
            " , " & sArrNShops(8) & " " & _
            " , " & sArrNShops(9) & " " & _
            " )"
        
        end if
        
        ObjConnectionSQL.Execute SQL
        
        if act.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassActivity", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            act.ID = NewID
        end if
        
        
        ' Guardamos el campo REAL DATA (NShops GPV)
        ' Por seguridad, sólo guardaremos si el usuario tiene permisos
        if isInputData() then
            
            set rst = Server.CreateObject("ADODB.RecordSet")
            
            SQL = "SELECT NShops " & _
            " FROM RealData " & _
            " WHERE WYear = " & act.WYear & " AND WMonth = " & act.WMonth & " AND WHalf = " & act.WHalf & _
            " AND IDBrand = " & act.IDBrand & " AND IDClient = " & act.IDClient
            rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
            if rst.EOF then
                SQL = "INSERT INTO RealData (WYear, WMonth, WHalf, IDClient, IDBrand, NShops"
                for iSubcat = 0 to 9
                    SQL = SQL & ", NShops" & iSubcat
                next
                SQL = SQL & ") VALUES (" & act.WYear & ", " & act.WMonth & ", " & act.WHalf & ", " & act.IDClient & ", " & IDBrand
                if isNull(act.RD_NShops) OR act.RD_NShops = "" then
                    SQL = SQL & ", NULL"
                else
                    SQL = SQL & ", " & act.RD_NShops
                end if
                
                for iSubcat = 0 to 9
                    if isNull(act.arrRD_NShops(iSubcat)) OR act.arrRD_NShops(iSubcat) = "" then
                        SQL = SQL & " , NULL "
                    else
                        SQL = SQL & " , " & act.arrRD_NShops(iSubcat)
                    end if
                next
                SQL = SQL & ")"
            else
                SQL = "UPDATE RealData " & _
                " SET NShops = "
                if isNull(act.RD_NShops) OR act.RD_NShops = "" then
                    SQL = SQL & " NULL"
                else
                    SQL = SQL & act.RD_NShops
                end if
                
                for iSubcat = 0 to 9
                    SQL = SQL & " , NShops" & iSubcat & " = "
                    if isNull(act.arrRD_NShops(iSubcat)) OR act.arrRD_NShops(iSubcat) = "" then
                        SQL = SQL & "NULL"
                    else
                        SQL = SQL & act.arrRD_NShops(iSubcat)
                    end if
                next
                SQL = SQL & " WHERE WYear = " & act.WYear & " AND WMonth = " & act.WMonth & " AND WHalf = " & act.WHalf & _
                " AND IDBrand = " & act.IDBrand & " AND IDClient = " & act.IDClient
            end if
            rst.Close
            
            ObjConnectionSQL.Execute SQL
            
            set rst = nothing
        end if
        
        
        ' check and arrange the form applied to the activity
        'dim aCli, aBra
        'set aCli = getClient(act.IDClient)
        'set aBra = getBrand(act.IDBrand)
        '''''' activityFormCheckAndArrange act, aBra, aCli
        
        
    end sub
    
    
    ' ############################################################################
    sub deleteActivity(id)
        
        dim SQL
        SQL = "DELETE FROM Activity WHERE ID = " & ID
        ObjConnectionSQL.Execute SQL
        
        ' Deletes the form responses if any
        SQL = "DELETE FROM ActivityForm WHERE IDActivity = " & ID
        ObjConnectionSQL.Execute SQL
        
    end sub
    
    
    ' ############################################################################
    function lastUpdateClient(idClient)
        dim rst, SQL 
        set rst = Server.CreateObject("ADODB.RecordSet")

        SQL = "SELECT TOP 1 id " & _
        " FROM Activity act " & _
        " WHERE act.IDClient = '" & IDClient & "' " & _
        " ORDER BY act.LastUpdatedDate DESC "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if NOT rst.EOF then
            lastUpdateClient = rst("id")
        else
            lastUpdateClient = 0
        end if
        rst.Close
        set rst = nothing
        
    end function
    
</script>