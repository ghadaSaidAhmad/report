<script runat=server language="vbscript">

    ' ############################################################################
    class ThemeClientBrand
        
        public ID
        public IDBrand
        public IDClient
        public WYear
        public WMonth
        public WHalf
        public Name
        public Status
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            IDBrand = 0
            IDClient = 0
            WYear = 1900
            WMonth = 1
            WHalf = 1
            Name = ""
            Status = ""
        end sub
        

        ' Devuelve el nombre del color según el Status
        public property get BGColor
            dim ret
            ret = ""
            Select Case Status
                case "Draft":
                    ret = "yellow"
                case "Approved":
                    ret = "lightgreen"
            End Select


            BGColor = ret
        end property
        
        ' Devuelve el nombre del color según el Status
        public property get FGColor
            dim ret
            ret = ""
            Select Case Status
                case "Draft":
                    ret = "red"
                case "Approved":
                    ret = "black"
            End Select


            FGColor = ret
        end property

    end class
    
    
    ' ############################################################################
    ' READ Theme data
    public function getThemeClientBrand(id)
        dim SQL
        dim rst
        
        dim thm
        set thm = new ThemeClientBrand
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM ThemeClientBrand WHERE ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            thm.ID = id
            thm.IDBrand = rst("IDBrand")
            thm.IDClient = rst("IDClient")
            thm.WYear = rst("WYear")
            thm.WMonth = rst("WMonth")
            thm.WHalf = rst("WHalf")
            thm.Name = rst("Name")
            thm.Status = rst("Status")
            
        else
            Err.Raise 555, "ClassThemeClientBrand", "Theme not found"
        end if
        rst.Close
        set rst = nothing
        
        set getThemeClientBrand = thm
    end function
    
    
    
    public function getThemesClientBrand(WYear, WMonth, IDClient, IDBrand)
        dim SQL
        dim rst
        dim thm1, thm2
        dim arrThemes(1) ' Array con un tema por quincena
        
        
        set thm1 = new ThemeClientBrand
        set thm2 = new ThemeClientBrand
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, WHalf FROM ThemeClientBrand " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                set thm1 = getThemeClientBrand(rst("id"))
            elseif rst("WHalf") = 2 then
                set thm2 = getThemeClientBrand(rst("id"))
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrThemes(0) = thm1
        set arrThemes(1) = thm2
        
        getThemesClientBrand = arrThemes
    end function
    
    

    ' ############################################################################
    'SAVE new or edit
    public sub saveThemeClientBrand(thm)
        dim SQL
        dim rst
        dim NewID
        
        if thm.ID > -1 then
            ' Not a new Theme
            SQL = "UPDATE ThemeClientBrand " & _
            " SET  " & _
            " Name = '" & replace(thm.Name, "'", "''") & "'" & _
            " WHERE id = " & thm.ID
            
        else
            ' New Theme
            SQL = "INSERT INTO ThemeClientBrand (IDBrand, IDClient, WYear, " & _
            " WMonth, WHalf, " & _
            " Name) " & _
            " VALUES (" & thm.IDBrand & ", " & thm.IDClient & ", " & thm.WYear & ", " & _
            " " & thm.WMonth & ", " & thm.WHalf & ", " & _
            " '" & Replace(thm.Name, "'", "''") & "')"
        
        end if
        ObjConnectionSQL.Execute SQL
        
        if thm.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassThemeClientBrand", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            thm.ID = NewID
        end if
        
    end sub
    
</script>