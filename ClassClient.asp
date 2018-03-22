<script runat=server language="vbscript">

    ' ############################################################################
    class Client
        
        public IDClient
        public Name
        public ShortName
        public indBaja
        public SiebelCode
        public ImageFileNameH
        public ImageFileNameV
        public activateForms
        
        'Init method
        public sub Class_Initialize()
            IDClient = -1
            Name = ""
            ShortName = ""
            indBaja = 0
            SiebelCode = ""
            ImageFileNameH = ""
            ImageFileNameV = ""
            activateForms = 0
        end sub
        
        public property get activatedForms
            if activateForms = 0 then
                activatedForms = false
            else
                activatedForms = true
            end if
        end property
        
    end class
    
    ' ############################################################################
    ' READ Client data
    public function getClient(id)
        dim SQL
        dim rst
        
        dim cli
        set cli = new Client
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT cli.* " & _
        " FROM Client cli " & _
        " WHERE cli.IDClient = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            cli.IDClient = rst("IDClient")
            cli.Name = rst("Name")
            cli.ShortName = rst("ShortName")
            cli.indBaja = rst("indBaja")
            cli.SiebelCode = rst("SiebelCode")
            cli.ImageFileNameH = rst("ImageFileNameH")
            cli.ImageFileNameV = rst("ImageFileNameV")
            cli.activateForms = rst("activateForms")
            
        else
            Err.Raise 555, "ClassClient", "Client not found"
        end if
        rst.Close
        set rst = nothing
        
        set getClient = cli
    end function
    
    
    
    ' ############################################################################
    ' tipoOrden = "NOMBRE" o "ORDEN"
    public function getClients(tipoOrden)
        dim SQL
        dim rst
        dim cli
        dim arrClients()
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT IDClient FROM Client WHERE indBaja=0"
        if tipoOrden = "NOMBRE" then
            SQL = SQL & " ORDER BY Name"
        else
            SQL = SQL & " ORDER BY Orden"
        end if
        
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        dim nClients
        nClients = 0
        while not rst.EOF
            
            redim preserve arrClients(nClients)
            set cli = getClient(rst("IDClient"))
            set arrClients(nClients) = cli
            
            nClients = nClients + 1
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getClients = arrClients
    end function
    
    

    ' ############################################################################
    'SAVE new or edit
    public sub saveClient(cli)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        if cli.IDClient > -1 then
            ' Not a new Client
            SQL = "UPDATE Client " & _
            " SET  " & _
            " Name = '" & replace(cli.Name, "'", "''") & "'" & _
            " , ShortName = '" & replace(cli.ShortName, "'", "''") & "'" & _
            " , indBaja = '" & replace(cli.indBaja, "'", "") & "'" & _
            " , SiebelCode = '" & replace(cli.SiebelCode, "'", "''") & "'" & _
            " , ImageFileNameH = '" & replace(cli.ImageFileNameH, "'", "''") & "'" & _
            " , ImageFileNameV = '" & replace(cli.ImageFileNameV, "'", "''") & "'" & _
            " , activateForms = '" & replace(cli.activateForms, "'", "") & "'" & _
            " WHERE IDClient = " & cli.IDClient
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Client (Name, ShortName, indBaja, SiebelCode, ImageFileNameH, ImageFileNameV, activateForms " & _
            " ) " & _
            " VALUES ('" & Replace(cli.Name, "'", "''") & "', '" & Replace(cli.ShortName, "'", "''") & "', '" & Replace(cli.indBaja, "'", "") & "', '" & Replace(cli.SiebelCode, "'", "''") & "', '" & Replace(cli.ImageFileNameH, "'", "''") & "', '" & Replace(cli.ImageFileNameV, "'", "''") & "', '" & Replace(cli.activateForms, "'", "") & "' " & _
            " )"
        
        end if

        ObjConnectionSQL.Execute SQL
        
        if cli.IDClient < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassClient", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            cli.IDClient = NewID
        end if
        
        
        
        
    end sub
    
    
    ' ############################################################################
    sub deleteClient(id)
        
        dim SQL
    	SQL = "UPDATE Client SET indBaja=1 WHERE IDClient=" & id
        ObjConnectionSQL.Execute SQL
        
    end sub
    
</script>