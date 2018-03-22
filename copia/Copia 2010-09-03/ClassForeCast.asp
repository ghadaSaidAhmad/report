<script runat=server language="vbscript">
    
    function getForecast(WYear, WMonth, IDClient, IDBrand)
        dim Forecast: Forecast = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT FC " & _
        " FROM ForeCast " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            Forecast = rst("FC")
        end if
        
        getForecast = Forecast
    end function
    
    
    function getForecastBrand(WYear, WMonth, IDBrand)
        dim Forecast: Forecast = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT SUM(FC) " & _
        " FROM ForeCast " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDBrand = " & IDBrand
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            Forecast = rst("FC")
        end if
        
        getForecastBrand = Forecast
    end function


    function getForecastClient(WYear, WMonth, IDClient)
        dim Forecast: Forecast = 0
        
        dim SQL
        dim rst
        
        
        SQL = "SELECT SUM(FC) " & _
        " FROM ForeCast " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient
        set rst = ObjConnectionSQL.Execute(SQL)
        if NOT rst.EOF then
            Forecast = rst("FC")
        end if
        
        getForecastClient = Forecast
    end function
    
</script>