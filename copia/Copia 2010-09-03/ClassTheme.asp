<script runat=server language="vbscript">

    ' ############################################################################
    class Theme
        
        public ID
        public Name
        public IDClient
        public ImageFileName
        public indBaja
        public LastUpdatedBy
        public LastUpdatedDate
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            Name = ""
            IDClient = -1
            ImageFileName = ""
            indBaja = 0
        end sub

    end class
    
    
    ' ############################################################################
    ' READ Theme data
    public function getTheme(id)
        dim SQL
        dim rst
        
        dim thm
        set thm = new Theme
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT thm.*, " & _
        " em.ApellidosNombre AS NombreUpdated, " & _
        " CONVERT(varchar, thm.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, thm.LastUpdatedDate, 108) AS UpdatedDate " & _
        " FROM Theme thm " & _
        " LEFT JOIN EmpleadosGlobal em ON thm.LastUpdatedBy = em.IDEmpleado " & _
        " WHERE thm.ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            thm.ID = id
            thm.Name = rst("Name")
            thm.IDClient = rst("IDClient")
            thm.ImageFileName = rst("ImageFileName")
            thm.indBaja = rst("indBaja")
            thm.LastUpdatedBy = rst("NombreUpdated")
            thm.LastUpdatedDate = rst("UpdatedDate")
            
        else
            Err.Raise 555, "ClassTheme", "Theme not found"
        end if
        rst.Close
        set rst = nothing
        
        set getTheme = thm
    end function
    
    
    ' ############################################################################
    ' READ Theme data by Name
    public function getThemeByName(Name, IDClient, CrearSiNoExiste)
        dim SQL
        dim rst
        
        dim thm
        set thm = new Theme
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT thm.* " & _
        " FROM Theme thm " & _
        " WHERE thm.Name = '" & replace(Name, "'", "''") & "' AND IDClient = " & IDClient
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            thm.ID = id
            thm.Name = rst("Name")
            thm.IDClient = rst("IDClient")
            thm.ImageFileName = rst("ImageFileName")
            thm.indBaja = rst("indBaja")
        else
            if CrearSiNoExiste then
                thm.ID = -1
                thm.Name = Name
                thm.IDClient = IDClient
                saveTheme(thm)
            end if
        end if
        rst.Close
        set rst = nothing
        
        set getThemeByName = thm
    end function
    

    ' ############################################################################
    public function getThemes(IDClient)
        dim SQL
        dim rst
        dim thm
        dim arrThemes()
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID FROM Theme " & _
        " WHERE IDClient = " & IDClient & " AND indBaja=0 " & _
        " ORDER BY Name "
        
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        dim nThemes
        nThemes = 0
        while not rst.EOF
            
            redim preserve arrThemes(nThemes)
            set thm = getTheme(rst("ID"))
            set arrThemes(nThemes) = thm
            
            nThemes = nThemes + 1
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getThemes = arrThemes
    end function
    

    ' ############################################################################
    public function getThemesIncludeCurrent(IDClient, IDTheme)
        dim SQL
        dim rst
        dim thm
        dim arrThemes()
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, Name FROM Theme " & _
        " WHERE IDClient = " & IDClient & " AND indBaja=0 " & _
        " UNION SELECT ID, Name FROM Theme WHERE ID = " & IDTheme & _
        " ORDER BY Name "
        
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        dim nThemes
        nThemes = 0
        while not rst.EOF
            
            redim preserve arrThemes(nThemes)
            set thm = getTheme(rst("ID"))
            if thm.indBaja<>0 then
                thm.Name = thm.Name & " (Deleted)"
            end if
            set arrThemes(nThemes) = thm
            
            nThemes = nThemes + 1
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getThemesIncludeCurrent = arrThemes
    end function    

    ' ############################################################################
    'SAVE new or edit
    public sub saveTheme(thm)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        if thm.ID > -1 then
            ' Not a new Theme
            SQL = "UPDATE Theme " & _
            " SET  " & _
            " Name = '" & replace(thm.Name, "'", "''") & "'" & _
            " , IDClient = '" & replace(thm.IDClient, "'", "") & "'" & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() "
            if thm.ImageFileName<>"" then
                SQL = SQL & " , ImageFileName = '" & replace(thm.ImageFileName, "'", "''") & "'"
            end if
            SQL = SQL & " , indBaja = '" & replace(thm.indBaja, "'", "") & "'" & _
            " WHERE id = " & thm.ID
            
        else
            isNew = true
            ' New Theme
            SQL = "INSERT INTO Theme (Name, IDClient, " & _
            " indBaja, " & _
            " LastUpdatedBy, LastUpdatedDate "
            if thm.ImageFileName<>"" then
                SQL = SQL & ", ImageFileName"
            end if
            SQL = SQL & " ) " & _
            " VALUES ('" & Replace(thm.Name, "'", "''") & "', '" & Replace(thm.IDClient, "'", "") & "', " & _
            " '" & Replace(thm.indBaja, "'", "") & "', " & _
            " " & session("IDEmpleado") & ", GETDATE() "
            if thm.ImageFileName<>"" then
                SQL = SQL & ", '" & Replace(thm.ImageFileName, "'", "''") & "'"
            end if
            SQL = SQL & " )"
        
        end if

        ObjConnectionSQL.Execute SQL
        
        if thm.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT MAX(ID) FROM Theme"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassTheme", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            thm.ID = NewID
        end if
        
    end sub
    
    
    public sub removeThemeImage(id)
        dim SQL
        
        SQL = "UPDATE Theme " & _
        " SET ImageFileName = '' " & _
        " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
        " , LastUpdatedDate = GETDATE() " & _
        " WHERE ID = " & id
        ObjConnectionSQL.Execute SQL
        
    end sub
    
    ' ############################################################################
    public sub deleteTheme(id)
        
        dim SQL
        'SQL = "DELETE FROM Theme WHERE ID = " & ID
        SQL = "UPDATE Theme " & _
        " SET indBaja = 1 " & _
        " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
        " , LastUpdatedDate = GETDATE() " & _
        " WHERE ID = " & ID
        
        ObjConnectionSQL.Execute SQL
        
    end sub
    
</script>