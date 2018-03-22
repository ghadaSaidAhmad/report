<script runat="server" language="vbscript">

    ' ############################################################################
    class Activity00
        
        public ID
        public IDBrand
        public IDClient
        public WYear
        public WMonth
        public WHalf
        public IDType
        public Name
        public LastUpdatedBy
        public LastUpdatedDate
        public IDTheme
        public ThemeName
        public ThemeImageFileName
        public IDRatio
        public BColor
        public FColor
        
                'Init method
        public sub Class_Initialize()
            ID = -1
            IDBrand = 0
            IDClient = 0
            WYear = 1900
            WMonth = 1
            WHalf = 1
            IDType = 0
            Name = ""
            IDTheme = -1
        end sub
        
        
        ' Texto que se muestra en el GRID
        public property get GridText
            
            GridText = ThemeName & " " & Name
            
        end property
        
        ' Devuelve el nombre del color según el Status
        public property get BGColor
            BGColor = BColor
        end property
        

        ' Devuelve el nombre del color según el Status
        public property get FGColor
            FGColor = FColor
        end property
        
    end class
    
    
    ' ############################################################################
    ' READ Activity data
    public function getActivity00(id)
        dim SQL
        dim rst
        
        dim act
        set act = new Activity00
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT act.*, " & _
        " CONVERT(varchar, act.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, act.LastUpdatedDate, 108) AS UpdatedDate, em.ApellidosNombre AS NombreUpdated, " & _
        " act0.IDTheme , act0.IDRatio, '' AS BGColor, '' AS FGColor, thm.Name AS [ThemeName], thm.ImageFileName AS [ThemeImageFileName] " & _
        " FROM Activity act " & _
        " INNER JOIN Activity00 act0 ON act.id = act0.id " & _
        " LEFT JOIN Theme thm ON act0.IDTheme = thm.id " & _
        " LEFT JOIN EmpleadosGlobal em ON act.LastUpdatedBy = em.IDEmpleado " & _
        " WHERE act.ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.IDBrand = rst("IDBrand")
            act.IDClient = rst("IDClient")
            act.WYear = rst("WYear")
            act.WMonth = rst("WMonth")
            act.WHalf = rst("WHalf")
            act.IDType = rst("IDType")
            act.Name = rst("Name")
            act.LastUpdatedBy = rst("NombreUpdated")
            act.LastUpdatedDate = rst("UpdatedDate")
            
            act.IDTheme = rst("IDTheme")
            act.ThemeName = rst("ThemeName")
            act.ThemeImageFileName = rst("ThemeImageFileName")
            act.IDRatio = rst("IDRatio")
            act.FColor = rst("FGColor")
            act.BColor = rst("BGColor")
            
        else
            Err.Raise 555, "ClassActivity00", "Activity not found " & SQL
        end if
        rst.Close
        set rst = nothing
        
        set getActivity00 = act
    end function
    
    
    
    ' ############################################################################
    public function getActivities00(WYear, WMonth, IDClient, IDBrand, IDType)
        dim SQL
        dim rst
        dim act1, act2
        dim arrActivity(1) ' Array con una actividad por quincena
        
        
        set act1 = new Activity00
        set act2 = new Activity00
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, WHalf FROM Activity " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand & _
        " AND IDType = " & IDType
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                set act1 = getActivity00(rst("id"))
            elseif rst("WHalf") = 2 then
                set act2 = getActivity00(rst("id"))
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrActivity(0) = act1
        set arrActivity(1) = act2
        
        getActivities00 = arrActivity
    end function
    
    

    ' ############################################################################
    'SAVE new or edit
    public sub saveActivity00(act)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        if act.ID > -1 then
            ' Not a new Activity
            SQL = "UPDATE Activity " & _
            " SET Name = '" & replace(act.Name, "'", "''") & "' " & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() " & _
            " WHERE id = " & act.ID
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Activity (IDBrand, IDClient, WYear " & _
            " , Name, WMonth, WHalf, IDType " & _
            " , LastUpdatedBy, LastUpdatedDate " & _
            " ) " & _
            " VALUES (" & act.IDBrand & ", " & act.IDClient & ", " & act.WYear & " " & _
            " , '" & replace(act.Name, "'", "''") & "', " & act.WMonth & ", " & act.WHalf & ", " & act.IDType & " " & _
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
                Err.Raise 1, "ClassActivity00", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            act.ID = NewID
        end if
        
        
        dim sIDTheme: sIDTheme= "NULL"
        if act.IDTheme <> "" then sIDTheme = act.IDTheme

        dim sIDRatio: sIDRatio= "NULL"
        if act.IDRatio <> "" then sIDRatio = act.IDRatio
        
        if NOT isNew then
            
            SQL = "UPDATE Activity00 " & _
            " SET " & _
            " IDTheme = " & sIDTheme & " " & _
            " , IDRatio = " & sIDRatio & " " & _
            " WHERE id = " & act.ID
            
        else
            
            SQL = "INSERT INTO Activity00 (" & _
            " ID " & _
            " , IDTheme " & _
            " , IDRatio " & _
            " ) " & _
            " VALUES (" & _
            " " & act.ID & " " & _
            " , " & sIDTheme & " " & _
            " , " & sIDRatio & " " & _
            " )"
            
        end if
        
        ObjConnectionSQL.Execute SQL
        
    end sub
    


    ' ############################################################################
    sub deleteActivity00(id)
        
        dim SQL
        SQL = "DELETE FROM Activity WHERE ID = " & ID
        ObjConnectionSQL.Execute SQL
        
        SQL = "DELETE FROM Activity00 WHERE ID = " & ID
        ObjConnectionSQL.Execute SQL

    end sub

</script>