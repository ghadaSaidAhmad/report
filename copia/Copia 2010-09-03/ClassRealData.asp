<script runat=server language="vbscript">

    ' ############################################################################
    class RealData
        
        public IDClient
        public IDBrand
        public WYear
        public WMonth
        public WHalf
        public NShops
        public PercentComplaint
        
        'Init method
        public sub Class_Initialize()
            WYear = 1900
            WMonth = 1
            WHalf = 1
            
        end sub

    end class
    
    
    ' ############################################################################
    ' READ RealData data
    public function getRealData(WYear, WMonth, WHalf, IDClient, IDBrand)
        dim SQL
        dim rst
        
        dim thm
        set thm = new RealData
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * " & _
        " FROM RealData " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & " AND WHalf = " & WHalf & " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            thm.IDClient = rst("IDClient")
            thm.IDBrand = rst("IDBrand")
            thm.WYear = rst("WYear")
            thm.WMonth = rst("WMonth")
            thm.WHalf = rst("WHalf")

            thm.PercentComplaint = rst("PercentComplaint")
            thm.NShops = rst("NShops")
        else
            ' No falla porque quizá no esté. Pero estará en blanco
            '''''''  Err.Raise 555, "ClassRealData", "Error reading Real Data"
        end if
        rst.Close
        set rst = nothing
        
        set getRealData = thm
    end function
    
    
    
    ' ############################################################################
    public function getRealDatas(IDClient, IDBrand, WYear, WMonth)
        dim SQL
        dim rst
        dim rd1, rd2
        dim arrRealData(1) ' Array con un registro por quincena
        
        set rd1 = new RealData
        set rd2 = new RealData
        
        rd1.IDClient = IDClient
        rd1.IDBrand = IDBrand
        rd1.WYear = WYear
        rd1.WMonth = WMonth
        rd1.WHalf = 1

        rd2.IDClient = IDClient
        rd2.IDBrand = IDBrand
        rd2.WYear = WYear
        rd2.WMonth = WMonth
        rd2.WHalf = 2

        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * " & _
        " FROM RealData " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                rd1.PercentComplaint = rst("PercentComplaint")
                rd1.NShops = rst("NShops")
            elseif rst("WHalf") = 2 then
                rd2.PercentComplaint = rst("PercentComplaint")
                rd2.NShops = rst("NShops")
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrRealData(0) = rd1
        set arrRealData(1) = rd2
        
        getRealDatas = arrRealData
    end function
    
        
</script>