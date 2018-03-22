<script runat=server language="vbscript">

    ' ############################################################################
    class Brand
        
        public IDBrand
        public Name
        public ShortName
        public indBaja
        public SiebelCode
        public Orden
        public ImageFileNameH
        public ImageFileNameV
        public idForm
        public FormName, FormDate
        public arrNShops(9)
        
        'Init method
        public sub Class_Initialize()
            IDBrand = -1
            Name = ""
            ShortName = ""
            indBaja = 0
            SiebelCode = ""
            Orden = 0
            ImageFileNameH = ""
            ImageFileNameV = ""
            idForm = -1
            
            FormName = ""
            FormDate = ""
            
            dim iSubcat
            for iSubcat = 0 to 9
                arrNShops(iSubcat) = ""
            next
            
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
        SQL = "SELECT bra.*, " & _
        " f.Name AS [FormName], CONVERT(varchar,f.DateFrom,103) AS [FormDate] " & _
        " FROM Brand bra " & _
		" LEFT JOIN Form f ON bra.idForm = f.idForm " & _
        " WHERE bra.IDBrand = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            bra.IDBrand = rst("IDBrand")
            bra.Name = rst("Name")
            bra.ShortName = rst("ShortName")
            bra.indBaja = rst("indBaja")
            bra.SiebelCode = rst("SiebelCode")
            bra.Orden = rst("Orden")
            bra.ImageFileNameH = rst("ImageFileNameH")
            bra.ImageFileNameV = rst("ImageFileNameV")
            if isNull(rst("idForm")) then
                bra.idForm = -1
            else
                bra.idForm = rst("idForm")
            end if
            
            bra.FormName = rst("FormName")
            bra.FormDate = rst("FormDate")
            
            dim iSubcat
            for iSubcat = 0 to 9
                bra.arrNShops(iSubcat) = rst("NShops" & iSubcat)
            next
            
        else
            Err.Raise 555, "ClassBrand", "Brand not found"
        end if
        rst.Close
        set rst = nothing
        
        set getBrand = bra
    end function
    
    
    
    ' ############################################################################
    ' tipoOrden = "NOMBRE" o "ORDEN"
    public function getBrands(tipoOrden)
        dim SQL
        dim rst
        dim bra
        dim arrBrands()
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT IDBrand FROM Brand WHERE indBaja=0 "
        if tipoOrden = "NOMBRE" then
            SQL = SQL & " ORDER BY Name"
        else
            SQL = SQL & " ORDER BY Orden"
        end if
        
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
            " Name = N'" & replace(bra.Name, "'", "''") & "'" & _
            " , ShortName = N'" & replace(bra.ShortName, "'", "''") & "'" & _
            " , indBaja = '" & replace(bra.indBaja, "'", "''") & "'" & _
            " , SiebelCode = N'" & replace(bra.SiebelCode, "'", "''") & "'" & _
            " , Orden = '" & replace(bra.Orden, "'", "") & "'" & _
            " , ImageFileNameH = '" & replace(bra.ImageFileNameH, "'", "''") & "'" & _
            " , ImageFileNameV = '" & replace(bra.ImageFileNameV, "'", "''") & "'" & _
            " , NShops0 = N'" & replace(bra.arrNShops(0), "'", "''") & "' " & _
            " , NShops1 = N'" & replace(bra.arrNShops(1), "'", "''") & "' " & _
            " , NShops2 = N'" & replace(bra.arrNShops(2), "'", "''") & "' " & _
            " , NShops3 = N'" & replace(bra.arrNShops(3), "'", "''") & "' " & _
            " , NShops4 = N'" & replace(bra.arrNShops(4), "'", "''") & "' " & _
            " , NShops5 = N'" & replace(bra.arrNShops(5), "'", "''") & "' " & _
            " , NShops6 = N'" & replace(bra.arrNShops(6), "'", "''") & "' " & _
            " , NShops7 = N'" & replace(bra.arrNShops(7), "'", "''") & "' " & _
            " , NShops8 = N'" & replace(bra.arrNShops(8), "'", "''") & "' " & _
            " , NShops9 = N'" & replace(bra.arrNShops(9), "'", "''") & "' " & _
            " WHERE IDBrand = " & bra.IDBrand
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Brand (Name, ShortName, indBaja, SiebelCode, Orden, ImageFileNameH, ImageFileNameV " & _
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
            " VALUES (N'" & Replace(bra.Name, "'", "''") & "', N'" & Replace(bra.ShortName, "'", "''") & "', '" & Replace(bra.indBaja, "'", "''") & "', N'" & Replace(bra.SiebelCode, "'", "''") & "', '" & Replace(bra.Orden, "'", "") & "', '" & Replace(bra.ImageFileNameH, "'", "''") & "', '" & Replace(bra.ImageFileNameV, "'", "''") & "'" & _
            " , N'" & Replace(bra.arrNShops(0), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(1), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(2), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(3), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(4), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(5), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(6), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(7), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(8), "'", "''") & "' " & _
            " , N'" & Replace(bra.arrNShops(9), "'", "''") & "' " & _
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
    	SQL = "UPDATE Brand SET indBaja=1 WHERE IDBrand=" & id
        ObjConnectionSQL.Execute SQL
        
    end sub
    
    
    ' ############################################################################
    public Function getFormBrands(idForm, queryType)
        dim SQL, rst, bra
        dim arrBrands()
        dim nBrands
        
        
        if queryType = "ASSIGNED_TO_FORM" then
            SQL = "SELECT IDBrand FROM Brand WHERE idForm = " & idForm & " ORDER BY Name "
        elseif queryType = "NOT_ASSIGNED" then
            SQL = "SELECT IDBrand FROM Brand WHERE (idForm IS NULL OR idForm = -1) ORDER BY Name "
        elseif queryType = "ASSIGNED_TO_OTHER_FORM" then
            SQL = "SELECT IDBrand FROM Brand WHERE idForm IS NOT NULL AND idForm <> -1 AND idForm <> '" & idForm & "' ORDER BY Name "
        end if
        
        set rst = ObjConnectionSQL.Execute(SQL)
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
        
        getFormBrands = arrBrands
    end Function
    
    
    
    class NumAndBrand
        public num
        public Brand
    end class
    
    ' ############################################################################
    public Function getFormAssignedBrandPromotions(idForm)
        dim SQL, rst, bra, nab
        dim arrBrands()
        dim nBrands
        
        
        
        SQL = "SELECT a.IDBrand, COUNT(a.id) AS N " & _
        " FROM Activity a " & _
        " INNER JOIN Brand b ON a.IDBrand = b.IDBrand " & _
        " WHERE a.idForm = '" & idForm & "' " & _
        " AND (a.IDForm <> b.IDForm  OR b.IDForm IS NULL )" & _
        " GROUP BY a.IDBrand, b.Name " & _
        " ORDER BY b.Name "


        set rst = ObjConnectionSQL.Execute(SQL)
        nBrands = 0
        while not rst.EOF
            
            set nab = new NumAndBrand
            
            set bra = getBrand(rst("IDBrand"))

            nab.num = rst("N")
            set nab.Brand = bra
            
            redim preserve arrBrands(nBrands)
            set arrBrands(nBrands) = nab
            
            nBrands = nBrands + 1
            
            rst.MoveNext
        wend
        rst.Close

        set rst = nothing
        
        getFormAssignedBrandPromotions = arrBrands
    end Function
    
    
    
</script>