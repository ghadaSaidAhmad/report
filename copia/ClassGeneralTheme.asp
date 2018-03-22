<script runat=server language="vbscript">

    ' ############################################################################
    class GeneralTheme
        
        public ID
        public IDClient
        public WYear
        public WMonth
        public WHalf
        public Name
        public LastUpdatedBy
        public LastUpdatedDate
        public IDTheme
        public ThemeName
        public ThemeImageFileName
        public BColor
        public FColor
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            WYear = 1900
            WMonth = 1
            WHalf = 1
            Name = ""
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
    ' READ GeneralTheme data
    public function getGeneralTheme(id)
        dim SQL
        dim rst
        
        dim thm
        set thm = new GeneralTheme
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT gthm.*, CONVERT(varchar, gthm.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, gthm.LastUpdatedDate, 108) AS UpdatedDate, em.ApellidosNombre AS NombreUpdated, " & _
        " gthm.IDTheme, thm.Name AS [ThemeName], thm.ImageFileName AS [ThemeImageFileName] " & _
        " FROM GeneralTheme gthm " & _
        " LEFT JOIN EmpleadosGlobal em ON gthm.LastUpdatedBy = em.IDEmpleado " & _
        " LEFT JOIN Theme thm ON thm.ID = gthm.IDTheme " & _
        " WHERE gthm.ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            thm.ID = id
            thm.IDClient = rst("IDClient")
            thm.WYear = rst("WYear")
            thm.WMonth = rst("WMonth")
            thm.WHalf = rst("WHalf")
            thm.Name = rst("Name")
            thm.LastUpdatedBy = rst("NombreUpdated")
            thm.LastUpdatedDate = rst("UpdatedDate")
            
            thm.IDTheme = rst("IDTheme")
            thm.ThemeName = rst("ThemeName")
            thm.ThemeImageFileName = rst("ThemeImageFileName")
            thm.BColor = Application("ColorActivity")
            thm.FColor = "Black"
        else
            Err.Raise 555, "ClassGeneralTheme", "General Theme not found"
        end if
        rst.Close
        set rst = nothing
        
        set getGeneralTheme = thm
    end function
    
    
    
    ' ############################################################################
    public function getGeneralThemes(IDClient, WYear, WMonth)
        dim SQL
        dim rst
        dim thm1, thm2
        dim arrGeneralTheme(1) ' Array con un tema por quincena
        
        set thm1 = new GeneralTheme
        set thm2 = new GeneralTheme
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, WHalf FROM GeneralTheme " & _
        " WHERE IDClient = " & IDClient & " AND WYear = " & WYear & " AND WMonth = " & WMonth
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                set thm1 = getGeneralTheme(rst("id"))
            elseif rst("WHalf") = 2 then
                set thm2 = getGeneralTheme(rst("id"))
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrGeneralTheme(0) = thm1
        set arrGeneralTheme(1) = thm2
        
        getGeneralThemes = arrGeneralTheme
    end function
    
    
    ' ############################################################################
    ' READ Activity data
    public function getGeneralThemeFromDate(IDClient, WYear, WMonth, WHalf)
        dim SQL
        dim rst
        
        dim thm
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT thm.ID " & _
        " FROM GeneralTheme thm " & _
        " WHERE thm.IDClient = " & IDClient & _
        " AND thm.WYear = " & WYear & " AND thm.WMonth = " & WMonth & " AND thm.WHalf = " & WHalf & " "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            set thm = getGeneralTheme(rst("ID"))
        else
            set thm = new GeneralTheme
        end if
        rst.Close
        set rst = nothing


        
        set getGeneralThemeFromDate = thm
    end function


    ' ############################################################################
    'SAVE new or edit
    public sub saveGeneralTheme(thm)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        dim sIDTheme: sIDTheme= "NULL"
        if thm.IDTheme <> "" then sIDTheme = thm.IDTheme

        if thm.ID > -1 then
            ' Not a new GeneralTheme
            SQL = "UPDATE GeneralTheme " & _
            " SET  " & _
            " Name = '" & replace(thm.Name, "'", "''") & "'" & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() " & _
            " , IDTheme = " & sIDTheme & " " & _
            " WHERE id = " & thm.ID
            
        else
            isNew = true
            ' New GeneralTheme
            SQL = "INSERT INTO GeneralTheme (IDClient, WYear " & _
            " , WMonth, WHalf, Name " & _
            " , LastUpdatedBy, LastUpdatedDate " & _
            " , IDTheme " & _
            " ) " & _
            " VALUES (" & thm.IDClient & ", " & thm.WYear & " " & _
            " , " & thm.WMonth & ", " & thm.WHalf & ", '" & Replace(thm.Name, "'", "''") & "' " & _
            " , " & session("IDEmpleado") & ", GETDATE() " & _
            " , " & sIDTheme & " " & _
            " )"
        
        end if

        ObjConnectionSQL.Execute SQL
        
        if thm.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassGeneralTheme", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            thm.ID = NewID
        end if
        
        
        
        
    end sub
    
    
    ' ############################################################################
    sub deleteGeneralTheme(id)
        
        dim SQL
        SQL = "DELETE FROM GeneralTheme WHERE ID = " & ID
        
        ObjConnectionSQL.Execute SQL
        
    end sub
    
</script>