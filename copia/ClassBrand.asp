<script runat=server language="vbscript">

    ' ############################################################################
    class Brand
        
        public IDBrand
        public Name
        public ShortName
        public indBaja
        public SiebelCode
        public ImageFileNameH
        public ImageFileNameV
        
        'Init method
        public sub Class_Initialize()
            IDBrand = -1
            Name = ""
            ShortName = ""
            indBaja = 0
            SiebelCode = ""
            ImageFileNameH = ""
            ImageFileNameV = ""
        end sub

    end class
    
    ' ############################################################################
    ' READ Brand data
    public function getBrand(id)
        dim SQL
        dim rst
        
        dim bra
        set bra = new Brand
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT bra.* " & _
        " FROM Brand bra " & _
        " WHERE bra.IDBrand = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            bra.IDBrand = rst("IDBrand")
            bra.Name = rst("Name")
            bra.ShortName = rst("ShortName")
            bra.indBaja = rst("indBaja")
            bra.SiebelCode = rst("SiebelCode")
            bra.ImageFileNameH = rst("ImageFileNameH")
            bra.ImageFileNameV = rst("ImageFileNameV")
            
        else
            Err.Raise 555, "ClassBrand", "Brand not found"
        end if
        rst.Close
        set rst = nothing
        
        set getBrand = bra
    end function
    
    
    
    ' ############################################################################
    public function getBrands()
        dim SQL
        dim rst
        dim bra
        dim arrBrands()
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT IDBrand FROM Brand WHERE indBaja=0 ORDER BY Name"
        
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        dim nBrands
        nBrands = 0
        while not rst.EOF
            
            redim preserve arrBrands(nBrands)
            set bra = getBrand(rst("IDBrand"))
            set arrBrands(nBrands) = bra
            
            nBrands = nBrands + 1
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getBrands = arrBrands
    end function
    
    

    ' ############################################################################
    'SAVE new or edit
    public sub saveBrand(bra)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        if bra.IDBrand > -1 then
            ' Not a new Brand
            SQL = "UPDATE Brand " & _
            " SET  " & _
            " Name = '" & replace(bra.Name, "'", "''") & "'" & _
            " , ShortName = '" & replace(bra.ShortName, "'", "''") & "'" & _
            " , indBaja = '" & replace(bra.indBaja, "'", "''") & "'" & _
            " , SiebelCode = '" & replace(bra.SiebelCode, "'", "''") & "'" & _
            " , ImageFileNameH = '" & replace(bra.ImageFileNameH, "'", "''") & "'" & _
            " , ImageFileNameV = '" & replace(bra.ImageFileNameV, "'", "''") & "'" & _
            " WHERE IDBrand = " & bra.IDBrand
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Brand (Name, ShortName, indBaja, SiebelCode, ImageFileNameH, ImageFileNameV " & _
            " ) " & _
            " VALUES ('" & Replace(bra.Name, "'", "''") & "', '" & Replace(bra.ShortName, "'", "''") & "', '" & Replace(bra.indBaja, "'", "''") & "', '" & Replace(bra.SiebelCode, "'", "''") & "', '" & Replace(ImageFileNameH, "'", "''") & "', '" & Replace(ImageFileNameV, "'", "''") & "' " & _
            " )"
        
        end if

        ObjConnectionSQL.Execute SQL
        
        if bra.IDBrand < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassBrand", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            bra.IDBrand = NewID
        end if
        
        
        
        
    end sub
    
    
    ' ############################################################################
    sub deleteBrand(id)
        
        dim SQL
        SQL = "DELETE FROM Brand WHERE IDBrand = " & id
        
        ObjConnectionSQL.Execute SQL
        
    end sub
    
</script>