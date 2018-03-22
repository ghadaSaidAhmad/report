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
        
        public Oferta, BGColorOferta, FGColorOferta
        public IDRatio
        public Folleto, BGColorFolleto, FGColorFolleto
        public Cabecera, BGColorCabecera, FGColorCabecera
        public NShops
        public PercentComplaint
        public IDStatus
        public Adicional
        
        ' de Real Data
        public RD_PercentComplaint
        public RD_NShops
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            IDBrand = 0
            IDClient = 0
            WYear = 1900
            WMonth = 1
            WHalf = 1
            Name = ""
            Oferta = "": BGColorOferta = "": FGColorOferta = ""
            IDRatio = 0
            Folleto = "": BGColorFolleto = "": FGColorFolleto = ""
            Cabecera = "": BGColorCabecera = "": FGColorCabecera = ""
            'NShops = 0
            'PercentComplaint = 0
            IDStatus = 1
            Adicional = ""
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
        " rat.Name AS [RatioName], rat.BGColor AS [RatioBG], rat.FGColor AS [RatioFG],  " & _
        " st.Name AS [StatusName], st.BGColor AS [StatusBG], st.FGColor AS [StatusFG],  " & _
        " CONVERT(varchar, act.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, act.LastUpdatedDate, 108) AS UpdatedDate, " & _
        " em.ApellidosNombre AS NombreUpdated, " & _
        " rd.PercentComplaint AS RD_PercentComplaint, rd.NShops AS RD_NShops " & _
        " FROM Activity act " & _
        " LEFT JOIN EmpleadosGlobal em ON act.LastUpdatedBy = em.IDEmpleado " & _
        " LEFT JOIN ActivityRatio rat ON act.IDRatio = rat.id " & _
        " LEFT JOIN ActivityStatus st ON act.IDStatus = st.id " & _
        " LEFT JOIN RealData rd ON act.WYear = rd.WYear AND act.WMonth = rd.WMonth AND act.WHalf = rd.WHalf AND act.IDClient = rd.IDClient AND act.IDBrand = rd.IDBrand " & _
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
            act.LastUpdatedDate = rst("UpdatedDate")
            
            act.Oferta = rst("Oferta"): act.BGColorOferta = rst("RatioBG"): act.FGColorOferta = rst("RatioFG")
            act.IDRatio = rst("IDRatio")
            act.Folleto = rst("Folleto"): act.BGColorFolleto = rst("StatusBG"): act.FGColorFolleto = rst("StatusFG")
            act.Cabecera = rst("Cabecera"): act.BGColorCabecera = rst("StatusBG"): act.FGColorCabecera = rst("StatusFG")
            act.NShops = rst("NShops")
            
            ' Percent Complaint viene de REAL DATA
            act.RD_PercentComplaint = rst("RD_PercentComplaint")
            act.RD_NShops = rst("RD_NShops")
            
            act.IDStatus = rst("IDStatus")
            act.Adicional = rst("Adicional")

        else
            Err.Raise 555, "ClassActivity", "Activity not found"
        end if
        rst.Close
        set rst = nothing

        set getActivity = act
    end function
    
    
    ' ############################################################################
    ' READ Activity data
    public function getActivityFromDate(IDClient, IDBrand, WYear, WMonth, WHalf)
        dim SQL
        dim rst
        
        dim act
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT act.ID " & _
        " FROM Activity act " & _
        " WHERE act.IDClient = " & IDClient & " AND act.IDBrand = " & IDBrand & _
        " AND act.WYear = " & WYear & " AND act.WMonth = " & WMonth & " AND act.WHalf = " & WHalf & " "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            set act = getActivity(rst("ID"))
        else
            set act = new Activity
        end if
        rst.Close
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
        isNew = false
        
        dim sIDRatio: sIDRatio= "NULL"
        if act.IDRatio <> "" then sIDRatio = act.IDRatio
        act.NShops = Replace(act.NShops, ",", ".")
        act.PercentComplaint = Replace(act.PercentComplaint, ",", ".")
        act.PercentComplaint = Replace(act.PercentComplaint, "%", "")
        if act.NShops<>"" then if NOT isNumeric(act.NShops) then act.NShops = 0
        if act.PercentComplaint<>"" then if NOT isNumeric(act.PercentComplaint) then act.PercentComplaint = 0
        
        dim sNShops: sNShops = "NULL"
        if act.NShops <> "" then sNShops = act.NShops
        dim sPercentComplaint: sPercentComplaint = "NULL"
        if act.PercentComplaint <> "" then sPercentComplaint = act.PercentComplaint

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
            " , PercentComplaint = " & sPercentComplaint & " " & _
            " , IDStatus = '" & replace(act.IDStatus, "'", "") & "' " & _
            " , Adicional = '" & replace(act.Adicional, "'", "''") & "' " & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() " & _
            " WHERE id = " & act.ID
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Activity (IDBrand, IDClient, WYear " & _
            " , WMonth, WHalf, Name " & _
            " , Oferta, IDRatio " & _
            " , Folleto, Cabecera " & _
            " , NShops, PercentComplaint " & _
            " , IDStatus, Adicional " & _
            " , LastUpdatedBy, LastUpdatedDate " & _
            " ) " & _
            " VALUES (" & act.IDBrand & ", " & act.IDClient & ", " & act.WYear & " " & _
            " , " & act.WMonth & ", " & act.WHalf & ", '" & Replace(act.Name, "'", "''") & "' " & _
            " , '" & replace(act.Oferta, "'", "''") & "', " & sIDRatio & " " & _
            " , '" & replace(act.Folleto, "'", "''") & "', '" & replace(act.Cabecera, "'", "''") & "' " & _
            " , " & sNShops & ", " & sPercentComplaint & " " & _
            " , '" & replace(act.IDStatus, "'", "") & "', '" & replace(act.Adicional, "'", "''") & "' " & _
            " , " & session("IDEmpleado") & ", GETDATE() " & _
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
        
        
        
        
    end sub
    
    
    ' ############################################################################
    sub deleteActivity(id)
        
        dim SQL
        SQL = "DELETE FROM Activity WHERE ID = " & ID
        
        ObjConnectionSQL.Execute SQL
        
    end sub
    
</script>