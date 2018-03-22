<script runat=server language="vbscript">
    
    function getNR(WYear, WMonth, IDClient, IDBrand)
        dim NR: NR = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT NR " & _
        " FROM NetRevenue " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            NR = rst("NR")
        end if
        
        getNR = NR
    end function
    
    
    function getNRBrand(WYear, WMonth, IDBrand)
        dim NR: NR = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT SUM(NR) " & _
        " FROM NetRevenue " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDBrand = " & IDBrand
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            NR = rst("NR")
        end if
        
        getNRBrand = NR
    end function


    function getNRClient(WYear, WMonth, IDClient)
        dim NR: NR = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT SUM(NR) " & _
        " FROM NetRevenue " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            NR = rst("NR")
        end if
        
        getNRClient = NR
    end function
    
</script>