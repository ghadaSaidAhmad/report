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
        
        public NShops0, NShops1, NShops2, NShops3, NShops4, NShops5, NShops6, NShops7, NShops8, NShops9
        public PercentComplaint0, PercentComplaint1, PercentComplaint2, PercentComplaint3, PercentComplaint4, PercentComplaint5, PercentComplaint6, PercentComplaint7, PercentComplaint8, PercentComplaint9

        public arrNShops(9), arrPercentComplaint(9)
        
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

            thm.NShops = rst("NShops")
            thm.PercentComplaint = rst("PercentComplaint")

            thm.NShops0 = rst("NShops0")
            thm.NShops1 = rst("NShops1")
            thm.NShops2 = rst("NShops2")
            thm.NShops3 = rst("NShops3")
            thm.NShops4 = rst("NShops4")
            thm.NShops5 = rst("NShops5")
            thm.NShops6 = rst("NShops6")
            thm.NShops7 = rst("NShops7")
            thm.NShops8 = rst("NShops8")
            thm.NShops9 = rst("NShops9")

            thm.PercentComplaint0 = rst("PercentComplaint0")
            thm.PercentComplaint1 = rst("PercentComplaint1")
            thm.PercentComplaint2 = rst("PercentComplaint2")
            thm.PercentComplaint3 = rst("PercentComplaint3")
            thm.PercentComplaint4 = rst("PercentComplaint4")
            thm.PercentComplaint5 = rst("PercentComplaint5")
            thm.PercentComplaint6 = rst("PercentComplaint6")
            thm.PercentComplaint7 = rst("PercentComplaint7")
            thm.PercentComplaint8 = rst("PercentComplaint8")
            thm.PercentComplaint9 = rst("PercentComplaint9")
            
            dim iSubcat
            for iSubcat = 0 to 9
                thm.arrNShops(iSubcat) = rst("NShops" & iSubcat)
                thm.arrPercentComplaint(iSubcat) = rst("PercentComplaint" & iSubcat)
            next
            
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
        dim iSubcat
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

                rd1.NShops0 = rst("NShops0")
                rd1.NShops1 = rst("NShops1")
                rd1.NShops2 = rst("NShops2")
                rd1.NShops3 = rst("NShops3")
                rd1.NShops4 = rst("NShops4")
                rd1.NShops5 = rst("NShops5")
                rd1.NShops6 = rst("NShops6")
                rd1.NShops7 = rst("NShops7")
                rd1.NShops8 = rst("NShops8")
                rd1.NShops9 = rst("NShops9")

                rd1.PercentComplaint0 = rst("PercentComplaint0")
                rd1.PercentComplaint1 = rst("PercentComplaint1")
                rd1.PercentComplaint2 = rst("PercentComplaint2")
                rd1.PercentComplaint3 = rst("PercentComplaint3")
                rd1.PercentComplaint4 = rst("PercentComplaint4")
                rd1.PercentComplaint5 = rst("PercentComplaint5")
                rd1.PercentComplaint6 = rst("PercentComplaint6")
                rd1.PercentComplaint7 = rst("PercentComplaint7")
                rd1.PercentComplaint8 = rst("PercentComplaint8")
                rd1.PercentComplaint9 = rst("PercentComplaint9")
                
                for iSubcat = 0 to 9
                    rd1.arrNShops(iSubcat) = rst("NShops" & iSubcat)
                    rd1.arrPercentComplaint(iSubcat) = rst("PercentComplaint" & iSubcat)
                next
                
            elseif rst("WHalf") = 2 then
                rd2.PercentComplaint = rst("PercentComplaint")
                rd2.NShops = rst("NShops")

                rd2.NShops0 = rst("NShops0")
                rd2.NShops1 = rst("NShops1")
                rd2.NShops2 = rst("NShops2")
                rd2.NShops3 = rst("NShops3")
                rd2.NShops4 = rst("NShops4")
                rd2.NShops5 = rst("NShops5")
                rd2.NShops6 = rst("NShops6")
                rd2.NShops7 = rst("NShops7")
                rd2.NShops8 = rst("NShops8")
                rd2.NShops9 = rst("NShops9")

                rd2.PercentComplaint0 = rst("PercentComplaint0")
                rd2.PercentComplaint1 = rst("PercentComplaint1")
                rd2.PercentComplaint2 = rst("PercentComplaint2")
                rd2.PercentComplaint3 = rst("PercentComplaint3")
                rd2.PercentComplaint4 = rst("PercentComplaint4")
                rd2.PercentComplaint5 = rst("PercentComplaint5")
                rd2.PercentComplaint6 = rst("PercentComplaint6")
                rd2.PercentComplaint7 = rst("PercentComplaint7")
                rd2.PercentComplaint8 = rst("PercentComplaint8")
                rd2.PercentComplaint9 = rst("PercentComplaint9")

                for iSubcat = 0 to 9
                    rd2.arrNShops(iSubcat) = rst("NShops" & iSubcat)
                    rd2.arrPercentComplaint(iSubcat) = rst("PercentComplaint" & iSubcat)
                next

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